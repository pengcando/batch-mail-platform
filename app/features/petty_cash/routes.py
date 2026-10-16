# -*- coding: utf-8 -*-
"""对账单批量作业 —— HTTP 接口。"""
from __future__ import annotations

import threading
import zipfile
from datetime import datetime
from pathlib import Path

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse

from app.config import ALLOW_REAL_SEND, JOBS_DIR, MAIL_SENDER, MAX_MERCHANTS, SMTP_HOST, SMTP_PORT
from app.core import mailer
from app.core.jobs import store
from app.features.petty_cash import engine as E
from app.features.petty_cash import pipeline

router = APIRouter()

MAX_UPLOAD_BYTES = 60 * 1024 * 1024          # 单文件 60MB 上限


async def _read_upload(file: UploadFile) -> bytes:
    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail=f"文件「{file.filename}」是空的")
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=400, detail=f"文件「{file.filename}」超过 60MB")
    return data


@router.get("/meta")
def meta():
    """给前端渲染表单用的静态元信息。"""
    return {
        "sender": MAIL_SENDER,
        "smtp": {"host": SMTP_HOST, "port": SMTP_PORT, "secure": True},
        "allowRealSend": ALLOW_REAL_SEND,
        "maxMerchants": MAX_MERCHANTS,
        "statuses": {
            "done": E.STATUS_DONE,
            "noFlow": E.STATUS_NO_FLOW,
            "failed": E.STATUS_FAILED,
        },
    }


@router.post("/inspect")
async def inspect_tool(toolFile: UploadFile = File(...)):
    """上传《商户对账工具》后，回出可选城市列表（用于城市下拉联动校验）。"""
    data = await _read_upload(toolFile)
    try:
        info = pipeline.inspect(data)
    except E.SourceError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:                                     # noqa: BLE001
        raise HTTPException(status_code=400, detail=f"解析失败：{exc}") from exc
    return JSONResponse(info)


@router.post("/verify-smtp")
async def verify_smtp(password: str = Form("")):
    """只验证授权码能否登录，不发任何邮件。"""
    if not password:
        raise HTTPException(status_code=400, detail="请先填写邮箱授权码")
    ok, message = mailer.verify_login(password)
    return {"ok": ok, "message": message}


@router.post("/run")
async def run_job(
    merchantFile: UploadFile = File(...),
    toolFile: UploadFile = File(...),
    city: str = Form(...),
    ym: str = Form(...),
    smtpPassword: str = Form(""),
    dryRun: bool = Form(True),
    skipDone: bool = Form(True),
    monthMode: str = Form("run_date"),
    numericFormat: bool = Form(False),
):
    m_bytes = await _read_upload(merchantFile)
    t_bytes = await _read_upload(toolFile)

    city = (city or "").strip()
    ym = (ym or "").strip()
    if not city:
        raise HTTPException(status_code=400, detail="请选择城市")
    if len(ym) != 7 or ym[4] != "-":
        raise HTTPException(status_code=400, detail="年月格式应为 YYYY-MM，例如 2026-09")

    if not dryRun:
        if not ALLOW_REAL_SEND:
            raise HTTPException(status_code=403, detail="当前服务已禁用正式发送（ALLOW_REAL_SEND=0）")
        if not smtpPassword.strip():
            raise HTTPException(status_code=400, detail="正式发送必须填写邮箱授权码")

    # 先做一次轻量校验，避免把明显错误留到后台线程才报
    try:
        wb_t = E.open_t(__import__("io").BytesIO(t_bytes))
        if city not in wb_t.sheetnames:
            available = "、".join(s["name"] for s in E.list_city_sheets(wb_t))
            raise HTTPException(status_code=400, detail=f"工具文件里没有「{city}」工作表，可用：{available}")
        wb_m = E.open_m(__import__("io").BytesIO(m_bytes))
        E.read_merchants(wb_m, skip_done=True)
    except E.SourceError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    options = E.RunOptions(
        city=city,
        ym=ym,
        smtp_password=smtpPassword,
        dry_run=dryRun,
        skip_done=skipDone,
        month_mode=monthMode,
        numeric_format=numericFormat,
        max_merchants=MAX_MERCHANTS,
    )

    job = store.create("petty-cash", f"{city} · {ym}", total=0)
    store.update(job.id, status="running", stage="排队中")

    def worker():
        def on_progress(evt: dict):
            processed = evt.get("processed")
            fields = {"stage": evt.get("stage", "")}
            if evt.get("total"):
                fields["total"] = evt["total"]
            if processed is not None:
                fields["processed"] = processed
            if evt.get("current"):
                fields["current"] = evt["current"]
            store.update(job.id, **fields)
            if evt.get("row"):
                store.append_row(job.id, evt["row"])

        try:
            outcome = pipeline.run(
                options, m_bytes, t_bytes, on_progress,
                is_cancelled=lambda: store.is_cancelled(job.id),
            )
            summary = outcome["summary"]

            zip_name = f"{city}-对账单批量作业-{ym}.zip"
            zip_path = None
            if outcome["files"]:
                # 只有真正产出了文件才打包；否则留空，避免给一个 22 字节的空 zip
                zip_path = JOBS_DIR / f"{job.id}.zip"
                with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                    for name, data in outcome["files"]:
                        zf.writestr(name, data)

            summary["previewHtml"] = outcome.get("preview_html", "")[:200000]
            cancelled = store.is_cancelled(job.id)
            store.update(
                job.id,
                status="cancelled" if cancelled else "done",
                stage="已取消" if cancelled else "完成",
                summary=summary,
                zip_path=str(zip_path) if zip_path else None,
                zip_name=zip_name if zip_path else None,
            )
        except Exception as exc:                                 # noqa: BLE001
            store.update(job.id, status="failed", stage="失败", error=f"{type(exc).__name__}: {exc}")

    threading.Thread(target=worker, name=f"petty-cash-{job.id}", daemon=True).start()
    return {"jobId": job.id}


@router.get("/jobs/{job_id}")
def job_status(job_id: str):
    job = store.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="任务不存在")
    return store.snapshot(job)


@router.post("/jobs/{job_id}/cancel")
def job_cancel(job_id: str):
    job = store.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="任务不存在")
    store.update(job_id, cancel_requested=True)
    return {"ok": True}


@router.get("/jobs/{job_id}/download")
def job_download(job_id: str):
    job = store.get(job_id)
    if job is None or not job.zip_path:
        raise HTTPException(status_code=404, detail="结果不存在或尚未生成")
    path = Path(job.zip_path)
    if not path.exists():
        raise HTTPException(status_code=404, detail="结果文件已被清理")
    stamp = datetime.now().strftime("%Y%m%d")
    return FileResponse(
        path,
        media_type="application/zip",
        filename=job.zip_name or f"result-{stamp}.zip",
    )
