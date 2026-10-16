# -*- coding: utf-8 -*-
"""
功能注册表。

新增功能只需两步：
  1. 在 app/features/ 下建一个包，其中 routes.py 暴露 `router`，
     并在包的 __init__.py 里写一个 FEATURE 字典（id/name/desc/icon/status）；
  2. 把它加进下面的 _MODULES 列表。

前端侧同理：在 web/js/features/ 下放一个模块并在 web/js/app.js 的注册表里登记。
"""
from importlib import import_module

_MODULES = [
    "app.features.petty_cash",
]


def feature_list() -> list[dict]:
    out = []
    for module_path in _MODULES:
        package = import_module(module_path)
        router_module = import_module(f"{module_path}.routes")
        meta = dict(getattr(package, "FEATURE", {}))
        meta.setdefault("id", module_path.rsplit(".", 1)[-1])
        meta.setdefault("name", meta["id"])
        meta.setdefault("desc", "")
        meta.setdefault("status", "ready")
        out.append({**meta, "router": router_module.router})
    return out


def register(app):
    for meta in feature_list():
        app.include_router(meta["router"], prefix=f"/api/{meta['id']}", tags=[meta["name"]])
