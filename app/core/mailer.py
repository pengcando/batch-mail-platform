# -*- coding: utf-8 -*-
"""
SMTP 发信。

- 发件人默认 sender@example.com，请通过环境变量 MAIL_SENDER 改成你自己的邮箱
- 主机默认 smtp.example.com，端口 465，SSL（各家邮箱服务商略有差异，按需改）
- 授权码由使用者在网页端输入，只在内存里用完即弃，绝不落盘、不写日志
"""
import smtplib
import ssl
from email.message import EmailMessage
from email.utils import formataddr, formatdate

from app.config import MAIL_SENDER, SMTP_HOST, SMTP_PORT, SMTP_USE_SSL


class MailError(Exception):
    pass


def send_mail(
    *,
    password: str,
    to: str,
    cc: str,
    subject: str,
    html_body: str,
    text_body: str,
    attachment_bytes: bytes,
    attachment_name: str,
    sender: str = MAIL_SENDER,
    host: str = SMTP_HOST,
    port: int = SMTP_PORT,
    timeout: int = 60,
):
    msg = EmailMessage()
    msg["From"] = formataddr(("对账单通知", sender))
    if to:
        msg["To"] = to
    if cc:
        msg["Cc"] = cc
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=True)

    msg.set_content(text_body or "请使用支持 HTML 的邮件客户端查看本邮件。")
    msg.add_alternative(html_body, subtype="html")

    msg.add_attachment(
        attachment_bytes,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=attachment_name,
    )

    recipients = [x for x in (to, cc) if x]
    if not recipients:
        raise MailError("没有有效收件人")

    try:
        context = ssl.create_default_context()
        if SMTP_USE_SSL:
            server = smtplib.SMTP_SSL(host, port, timeout=timeout, context=context)
        else:
            server = smtplib.SMTP(host, port, timeout=timeout)
            server.starttls(context=context)
        with server:
            server.login(sender, password)
            server.send_message(msg, from_addr=sender, to_addrs=recipients)
    except smtplib.SMTPAuthenticationError as exc:
        raise MailError(f"授权码校验失败：{exc}") from exc
    except Exception as exc:                                     # noqa: BLE001
        raise MailError(f"{type(exc).__name__}: {exc}") from exc


def verify_login(password: str, sender: str = MAIL_SENDER, host: str = SMTP_HOST, port: int = SMTP_PORT):
    """只做一次登录校验，用于「测试 SMTP 连接」按钮；不会发任何邮件。"""
    try:
        context = ssl.create_default_context()
        server = smtplib.SMTP_SSL(host, port, timeout=20, context=context)
        with server:
            server.login(sender, password)
        return True, "SMTP 登录成功"
    except Exception as exc:                                     # noqa: BLE001
        return False, f"{type(exc).__name__}: {exc}"
