# -*- coding: utf-8 -*-
"""
对账单批量作业 —— 流水线编排（阶段 A → B → C）。

进度通过 progress 回调往外抛，网页端据此渲染进度条与逐商户状态表。
"""
from __future__ import annotations

import csv
import io
import time
from collections import Counter
from datetime import datetime

from openpyxl import load_workbook

from app.config import ALLOW_REAL_SEND, MAIL_SENDER
from app.core import email_html
from app.core.mailer import MailError, send_mail
from app.core.xlsx_style import reorder_rows_by_key
from app.features.petty_cash import engine as E


def _progress(stage: str, **extra):
    return {"stage": stage, **extra}


def inspect(tool_bytes: bytes) -> dict:
    """上传《商户对账工具》后，列出可用城市 + 基础校验信息。"""
    wb_t = E.open_t(io.BytesIO(tool_bytes))
    cities = E.list_city_sheets(wb_t)
    missing = [s for s in (E.STATEMENT_SHEET, E.ADJUST_SHEET, E.INVOICE_SHEET) if s not in wb_t.sheetnames]
    return {
        "cities": cities,
        "sheets": wb_t.sheetnames,
        "missingSheets": missing,
        "hasDataSheet": E.DATA_SHEET in wb_t.sheetnames,
    }


def run(options: E.RunOptions, m_bytes: bytes, t_bytes: bytes, progress, is_cancelled=None) -> dict:
    """
    执行全流程，返回 {results, files, summary, preview_html}。
    progress: callable(dict)  —— 抛进度事件
    is_cancelled: callable() -> bool —— 用户取消检查
    """
    started = time.time()
    is_cancelled = is_cancelled or (lambda: False)

    progress(_progress("初始化：打开两份源文件"))
    wb_m = E.open_m(io.BytesIO(m_bytes))
    wb_t = E.open_t(io.BytesIO(t_bytes))

    templates = E.load_templates(wb_t)
    month_value = E.compute_month_value(options.month_mode, options.ym)

    # ---- A11：M 表按 A 列降序（顺手把行重排好，之后记录的绝对行号才是最终行号）----
    ws_m = wb_m[E.M_SHEET]
    reorder_rows_by_key(
        ws_m,
        key_index=E.M_COL_NAME,
        reverse=True,
        first_row=2,
        last_row=ws_m.max_row,
        ncols=E.NCOL_M,
    )

    # ---- A12/A13：读取循环数据（按开关跳过已有进度的行）----
    merchants = E.read_merchants(wb_m, skip_done=options.skip_done)
    if not merchants:
        return {
            "results": [],
            "files": [],
            "summary": {
                "city": options.city,
                "ym": options.ym,
                "dryRun": options.dry_run,
                "total": 0,
                "noFlow": 0,
                "sent": 0,
                "files": 0,
                "failed": 0,
                "skipped": 0,
                "note": "没有需要处理的商户：名单里所有行的「进度」列都已有值（skipDone 打开时即如此）。"
                        "如需全量重跑，请关闭「跳过已处理行」。",
            },
            "preview_html": "",
        }

    limit = options.max_merchants or 0
    if limit and len(merchants) > limit:
        merchants = merchants[:limit]

    # ---- B-1：按商户把城市明细分桶，避免每轮重复扫描 560 行 ----
    progress(_progress("读取城市明细", total=len(merchants)))
    city_rows = E.read_city_rows(wb_t, options.city)
    buckets: dict[str, list[list]] = {}
    for row in city_rows:
        buckets.setdefault(E._s(row[E.CITY_COL_SHOP - 1]), []).append(row)

    # ---- 增减项一次性读完，循环内只做匹配 ----
    adjust_all = []
    ws_adjust = wb_t[E.ADJUST_SHEET]
    for r in range(2, ws_adjust.max_row + 1):
        if E._s(ws_adjust.cell(row=r, column=E.ADJUST_COL_CITY).value) != options.city:
            continue
        adjust_all.append([ws_adjust.cell(row=r, column=c).value for c in range(1, E.ADJUST_NCOL + 1)])
    adjust_index: dict[str, list[list]] = {}
    for row in adjust_all:
        adjust_index.setdefault(E._s(row[E.ADJUST_COL_SHOP - 1]), []).append(row)

    results: list[E.MerchantResult] = []
    output_files: list[tuple[str, bytes]] = []
    preview_html = ""
    counters = Counter()

    total = len(merchants)
    for idx, merchant in enumerate(merchants, start=1):
        if is_cancelled():
            progress(_progress("已取消", total=total, processed=idx - 1))
            break

        progress(_progress("处理中", total=total, processed=idx - 1, current=merchant.name))

        rows = buckets.get(merchant.name, [])
        n = len(rows)

        if n == 0:
            # 分支 1：本月无流水（B8a/B8b）—— 不发信、不产文件，只回写状态
            result = E.MerchantResult(
                merchant=merchant.name,
                to="、".join(merchant.to_list),
                cc="、".join(merchant.cc_list),
                status=E.STATUS_NO_FLOW,
                detail="城市明细中找不到该商户的记录",
                rows=0,
            )
            results.append(result)
            counters["noFlow"] += 1
            if not options.dry_run:
                ws_m.cell(row=merchant.excel_row, column=E.M_COL_PROGRESS).value = E.STATUS_NO_FLOW
            progress(_progress("处理中", total=total, processed=idx, row=result.as_dict()))
            continue

        statement_rows = [E.build_statement_row(r, month_value) for r in rows]
        adjust_rows = adjust_index.get(merchant.name, [])
        m = len(adjust_rows)

        file_name = E.ATTACH_TEMPLATE.format(merchant=merchant.name, ym=options.ym)
        try:
            xlsx_bytes = E.build_statement_workbook(
                templates, statement_rows, adjust_rows, numeric_format=options.numeric_format
            )
        except Exception as exc:                                  # noqa: BLE001
            entry = E.MerchantResult(
                merchant=merchant.name, to="、".join(merchant.to_list), cc="、".join(merchant.cc_list),
                status=E.STATUS_FAILED, detail=f"生成对账单失败：{exc}", rows=n, invoice_rows=m,
            )
            results.append(entry)
            counters["failed"] += 1
            if not options.dry_run:
                ws_m.cell(row=merchant.excel_row, column=E.M_COL_PROGRESS).value = E.STATUS_FAILED
            progress(_progress("处理中", total=total, processed=idx, row=entry.as_dict()))
            continue

        output_files.append((file_name, xlsx_bytes))

        invoice = E.build_invoice_context(
            [merchant.name, merchant.invoice_info, merchant.invoice_type, merchant.remark,
             merchant.to_raw, merchant.cc_raw, ""],
            statement_rows, templates.invoice_intro, templates.invoice_tip,
        )
        body_html = email_html.wrap_html(email_html.render_invoice_html(invoice))
        body_text = email_html.render_invoice_text(invoice)
        subject = E.SUBJECT_TEMPLATE.format(merchant=merchant.name, ym=options.ym)
        if not preview_html:
            preview_html = body_html

        to_list, cc_list = merchant.to_list, merchant.cc_list
        to_hdr, cc_hdr = ", ".join(to_list), ", ".join(cc_list)

        detail_parts = [f"{n} 条明细", f"{m} 条增减项"]
        if not to_list:
            detail_parts.append("收件人格式异常")

        if options.dry_run:
            status = E.STATUS_DONE if to_list else E.STATUS_FAILED
            if not to_list:
                detail_parts.append("无有效收件人，正式发送时会记为「未完成」")
            else:
                detail_parts.append("预览模式：未发送")
        elif not to_list:
            status = E.STATUS_FAILED
        else:
            try:
                send_mail(
                    password=options.smtp_password,
                    to=to_hdr,
                    cc=cc_hdr,
                    subject=subject,
                    html_body=body_html,
                    text_body=body_text,
                    attachment_bytes=xlsx_bytes,
                    attachment_name=file_name,
                    sender=MAIL_SENDER,
                )
                status = E.STATUS_DONE
            except MailError as exc:
                status = E.STATUS_FAILED
                detail_parts.append(str(exc))

        entry = E.MerchantResult(
            merchant=merchant.name,
            to="、".join(to_list),
            cc="、".join(cc_list),
            status=status,
            detail="；".join(detail_parts),
            rows=n,
            invoice_rows=m,
            file_name=file_name,
        )
        results.append(entry)
        counters["sent" if status == E.STATUS_DONE else "failed"] += 1

        if not options.dry_run:
            ws_m.cell(row=merchant.excel_row, column=E.M_COL_PROGRESS).value = status
        progress(_progress("处理中", total=total, processed=idx, row=entry.as_dict()))

    # ---- 阶段 C：收尾。M 表状态保存（预览模式不动源文件）----
    files: list[tuple[str, bytes]] = list(output_files)

    if not options.dry_run:
        progress(_progress("保存《不开票商户明细》", total=total, processed=total))
        m_buf = io.BytesIO()
        wb_m.save(m_buf)
        files.append((f"不开票商户明细-{options.city}.xlsx", m_buf.getvalue()))

    files.append((f"发信结果-{options.city}-{options.ym}.csv", _results_csv(results)))
    if preview_html:
        files.append(("_邮件正文预览.html", preview_html.encode("utf-8")))

    summary = {
        "city": options.city,
        "ym": options.ym,
        "dryRun": options.dry_run,
        "skipDone": options.skip_done,
        "total": total,
        "noFlow": counters["noFlow"],
        "sent": counters["sent"],
        "failed": counters["failed"],
        "files": len(output_files),
        "adjustMerchants": len(adjust_all),
        "elapsed": round(time.time() - started, 1),
        "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "sender": MAIL_SENDER,
    }
    progress(_progress("完成", total=total, processed=total))
    return {"results": results, "files": files, "summary": summary, "preview_html": preview_html}


def _results_csv(results: list[E.MerchantResult]) -> bytes:
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(["商户名称", "收件人", "抄送", "状态", "明细行数", "增减项条数", "附件名", "说明"])
    for r in results:
        writer.writerow([r.merchant, r.to, r.cc, r.status, r.rows, r.invoice_rows, r.file_name, r.detail])
    return buf.getvalue().encode("utf-8-sig")      # BOM：Excel 打开不乱码
