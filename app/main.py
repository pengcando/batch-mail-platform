# -*- coding: utf-8 -*-
"""
邮件发送平台 —— 应用入口。

启动：
    python -m uvicorn app.main:app --host 0.0.0.0 --port 8000
或直接双击 run.bat。

设计为「外壳 + 功能插件」结构：main 只负责挂载功能路由与静态页面，
每个业务功能是一个独立模块，后续加功能不需要动这个文件。
"""
from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles

from app.config import WEB_DIR
from app.features import register

app = FastAPI(title="邮件发送平台", version="0.1.0")

# 功能路由（必须在静态资源挂载之前注册，否则会被 "/" 通配吃掉）
register(app)


@app.get("/api/features")
def api_features():
    from app.features import feature_list

    return [
        {k: v for k, v in meta.items() if k != "router"}
        for meta in feature_list()
    ]


@app.get("/api/health")
def health():
    return {"ok": True}


@app.exception_handler(Exception)
async def unhandled(request, exc):                               # noqa: ANN001
    if isinstance(exc, HTTPException):
        return JSONResponse({"detail": exc.detail}, status_code=exc.status_code)
    return JSONResponse({"detail": f"{type(exc).__name__}: {exc}"}, status_code=500)


app.mount("/", StaticFiles(directory=str(WEB_DIR), html=True), name="web")
