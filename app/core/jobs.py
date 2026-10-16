# -*- coding: utf-8 -*-
"""
极简内存任务仓库 —— 支撑网页端的「进度条 + 逐商户状态」实时反馈。

单体部署够用；将来上服务器要多实例时，把这里换成 Redis / 数据库即可，
对外接口（create / get / update / append_row）保持不变。
"""
import threading
import time
import uuid
from dataclasses import dataclass, field
from typing import Any


@dataclass
class Job:
    id: str
    feature: str
    title: str
    status: str = "pending"          # pending | running | done | failed | cancelled
    stage: str = ""
    total: int = 0
    processed: int = 0
    current: str = ""                # 当前正在处理的商户，进度条上实时显示
    rows: list = field(default_factory=list)
    summary: dict = field(default_factory=dict)
    error: str = ""
    zip_path: str = ""
    zip_name: str = ""
    cancel_requested: bool = False
    created_at: float = field(default_factory=time.time)
    updated_at: float = field(default_factory=time.time)


class JobStore:
    def __init__(self):
        self._jobs: dict[str, Job] = {}
        self._lock = threading.Lock()

    def create(self, feature: str, title: str, total: int = 0) -> Job:
        job = Job(id=uuid.uuid4().hex[:16], feature=feature, title=title, total=total)
        with self._lock:
            self._jobs[job.id] = job
        return job

    def get(self, job_id: str) -> Job | None:
        with self._lock:
            return self._jobs.get(job_id)

    def update(self, job_id: str, **fields: Any) -> Job | None:
        with self._lock:
            job = self._jobs.get(job_id)
            if job is None:
                return None
            for key, value in fields.items():
                setattr(job, key, value)
            job.updated_at = time.time()
            return job

    def append_row(self, job_id: str, row: dict) -> Job | None:
        with self._lock:
            job = self._jobs.get(job_id)
            if job is None:
                return None
            job.rows.append(row)
            job.processed = len(job.rows)
            job.updated_at = time.time()
            return job

    def is_cancelled(self, job_id: str) -> bool:
        with self._lock:
            job = self._jobs.get(job_id)
            return bool(job and job.cancel_requested)

    def snapshot(self, job: Job) -> dict:
        return {
            "id": job.id,
            "feature": job.feature,
            "title": job.title,
            "status": job.status,
            "stage": job.stage,
            "total": job.total,
            "processed": job.processed,
            "current": job.current,
            "rows": list(job.rows),
            "summary": dict(job.summary),
            "error": job.error,
            "hasResult": bool(job.zip_path),
            "zipName": job.zip_name,
            "elapsed": round(time.time() - job.created_at, 1),
        }


store = JobStore()
