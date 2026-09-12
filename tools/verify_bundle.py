# -*- coding: utf-8 -*-
"""构建后自检：运行时资源真的进了产物吗？
==================================================================
用法：``python tools/verify_bundle.py [dist_dir]``（默认 ``dist``）

为什么需要它
------------
v1.1.0 及更早的 ``build.py`` 没有把 ``tfd_app/assets`` 带进产物，而 gui.py /
watermark.py 读取资源时都写了 ``if os.path.isfile(...)`` 守卫 —— 结果打包版
「**二维码不显示、水印图丢失**」，还**全程静默、不报错**，直到人工打开软件才可能
发现。这个脚本把这类"看着正常、功能缺件"的失败变成**构建红灯**。

清单来源：``tfd_app/assetpath.py`` 的 ``REQUIRED_ASSETS``（唯一事实来源）。
"""
import os
import sys

# Windows runner 控制台默认 cp1252：直接 print 中文（全角冒号「：」等）会抛
# UnicodeEncodeError，**让自检自己把构建搞红**（v1.1.1 Windows 首跑就栽在这）。
# 与 build.py 同款兜底：强制 UTF-8，编不出的字符替换掉，绝不让日志编码左右构建结果。
for _stream in ("stdout", "stderr"):
    try:
        getattr(sys, _stream).reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tfd_app.assetpath import REQUIRED_ASSETS  # noqa: E402


def main(dist="dist"):
    if not os.path.isdir(dist):
        print("找不到产物目录：%s" % dist)
        return 1
    found = set()
    for _root, _dirs, files in os.walk(dist):
        found.update(files)
    missing = [n for n in REQUIRED_ASSETS if n not in found]
    if missing:
        print("产物缺少运行时资源：%s" % ", ".join(missing))
        print("（说明 build.py 的 --include-data-files 没生效 —— 见 tfd_app/assetpath.py）")
        return 1
    print("OK：%d 个运行时资源均已打包进 %s" % (len(REQUIRED_ASSETS), dist))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1] if len(sys.argv) > 1 else "dist"))
