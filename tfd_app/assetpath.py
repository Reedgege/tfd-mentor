# -*- coding: utf-8 -*-
"""运行时资源路径解析 · 导师版
==================================================================

gui.py 要读二维码/窗口图标，watermark.py 要读水印图片，都放在
``tfd_app/assets/``。但**打包后布局会变**（Nuitka 把包编进 exe、数据目录被拷到
别处），写死一条 ``<模块目录>/assets/xxx`` 就会在打包版里静默读不到。

历史事故（v1.1.0 及更早）：``build.py`` 根本没把 ``tfd_app/assets`` 带进产物，
而读取方又只有那条写死路径 —— 结果打包版「二维码不显示、水印图片丢失」，
因为每处都写了 ``if os.path.isfile(...)`` 守卫，**全程静默、不报错**，很难发现。

本模块做两件事：
  1. 提供 ``find_asset(name)``：按几种可能的布局依次探测，返回第一个真实存在的
     路径；都没有返回 ``None``（调用方静默跳过 —— 资源只是锦上添花，绝不影响启动）。
  2. 提供 ``REQUIRED_ASSETS``：**必须随产物携带**的资源清单，供 build.py 作为
     唯一事实来源（避免"清单里漏一个"再次静默复发），并供测试断言文件确实存在。
"""
from __future__ import annotations

import os
import sys

# 必须随产物携带的资源（build.py 据此生成 --include-data-files=…）。
REQUIRED_ASSETS = (
    "icon.png",            # 窗口图标（Tk 用 PNG）
    "icon.ico",            # Windows 窗口/文件图标
    "icon.icns",           # macOS
    "qrcode.png",          # 咨询/客服二维码
    "miniapp_qrcode.png",  # 小程序二维码
    "watermark.png",       # 试用水印图（正文穿插）
    "watermark_bg.png",    # 试用水印背景图（页眉 VML）
    "advisor_background.png",  # 导师版主界面背景图（Backdrop 铺底，缺则降级为纯色）
)


def _asset_dir() -> str:
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")


def _runtime_bases() -> list:
    """打包后「可执行文件所在目录」的候选（源码运行时基本用不上）。"""
    bases = []
    raws = []
    argv = getattr(sys, "argv", None) or []
    if argv:
        raws.append(argv[0])
    raws.append(getattr(sys, "executable", "") or "")
    for raw in raws:
        try:
            if raw:
                bases.append(os.path.dirname(os.path.abspath(raw)))
        except Exception:
            continue
    try:
        bases.append(os.getcwd())
    except Exception:
        pass
    out = []
    for b in bases:
        if b and b not in out:
            out.append(b)
    return out


def candidates(name: str) -> list:
    """按「最可能命中」的顺序返回候选路径（含不存在的）。"""
    here = os.path.dirname(os.path.abspath(__file__))
    cands = [
        os.path.join(here, "assets", name),          # 源码运行 / Nuitka 包内布局
        os.path.join(here, "..", "assets", name),    # 资源被放在包外一层
    ]
    for base in _runtime_bases():
        cands += [
            os.path.join(base, "assets", name),          # 资源与 exe 同级
            os.path.join(base, "tfd_app", "assets", name),
            os.path.join(base, name),
        ]
    out = []
    for c in cands:
        try:
            c = os.path.normpath(os.path.abspath(c))
        except Exception:
            continue
        if c not in out:
            out.append(c)
    return out


def find_asset(name: str):
    """返回第一个真实存在的资源路径；都没有则 None（调用方静默跳过）。"""
    for path in candidates(name):
        try:
            if os.path.isfile(path):
                return path
        except Exception:
            continue
    return None


def missing_assets() -> list:
    """返回在**源码树**里缺失的资源名（供测试/构建前自检）。"""
    out = []
    for name in REQUIRED_ASSETS:
        if find_asset(name) is None:
            out.append(name)
    return out
