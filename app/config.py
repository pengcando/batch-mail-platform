# -*- coding: utf-8 -*-
"""全局配置：可被环境变量覆盖，便于后续上服务器时改而不动代码。"""
import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent      # mail-platform/
WEB_DIR = BASE_DIR / "web"
STORAGE_DIR = BASE_DIR / "storage"

# 任务产物的存放目录（zip、生成的 xlsx）；服务器上可按需挂载到磁盘
JOBS_DIR = STORAGE_DIR / "jobs"
UPLOADS_DIR = STORAGE_DIR / "uploads"

for _d in (STORAGE_DIR, JOBS_DIR, UPLOADS_DIR):
    _d.mkdir(parents=True, exist_ok=True)

# ---- 邮件固定参数 ----
# 发件人、SMTP 主机请按你自己的邮箱服务商改，或用同名环境变量覆盖（推荐）：
#   set MAIL_SENDER=your@company.com
#   set SMTP_HOST=smtp.your-provider.com
MAIL_SENDER = os.getenv("MAIL_SENDER", "sender@example.com")
SMTP_HOST = os.getenv("SMTP_HOST", "smtp.example.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "465"))
SMTP_USE_SSL = os.getenv("SMTP_USE_SSL", "1") == "1"

# 单次任务最多处理多少商户（防止误操作把整批发出去；0 = 不限制）
MAX_MERCHANTS = int(os.getenv("MAX_MERCHANTS", "0"))

# 是否允许「正式发送」模式（上服务器后建议保持 1；本地调试可设 0 彻底锁死）
ALLOW_REAL_SEND = os.getenv("ALLOW_REAL_SEND", "1") == "1"
