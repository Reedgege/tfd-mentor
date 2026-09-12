# -*- coding: utf-8 -*-
"""运行时资源回归（导师版）
==================================================================
背景：v1.1.0 及更早的 build.py 没把 ``tfd_app/assets`` 带进产物，而读取方都写了
``if os.path.isfile(...)`` 守卫 → 打包版「**二维码不显示、水印图丢失**」且**全程
静默**。本测试在**打包前**就把这类缺件钉住：

  1. ``REQUIRED_ASSETS`` 每个文件在源码树里真实存在（缺一个 = 会打个缺件的包）；
  2. ``assetpath.find_asset`` 能定位到它们（多候选路径逻辑没写坏）；
  3. gui 的 ICON / QRCODE / MINIAPP_QRCODE 与 watermark 的两张水印图路径，
     都指向真实文件（导入即求值，等于校验了资源接线）。

用法：``python tests/test_assets.py``（无 tkinter 的环境会自动跳过第 3 节）
"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from tfd_app import assetpath          # noqa: E402

fails = []


def check(cond, label):
    print(("  ok  " if cond else "  FAIL ") + label)
    if not cond:
        fails.append(label)


print("=== 1. 必需资源清单 ===")
missing = assetpath.missing_assets()
check(not missing, "REQUIRED_ASSETS 全部存在（缺失：%s）" % (missing or "无"))
for name in assetpath.REQUIRED_ASSETS:
    p = assetpath.find_asset(name)
    check(p is not None and os.path.isfile(p), "find_asset(%r) -> %s" % (name, p))

print("=== 2. 构建脚本的清单与资源一致 ===")
build_py = os.path.join(ROOT, "build.py")
src = open(build_py, encoding="utf-8").read()
check("REQUIRED_ASSETS" in src, "build.py 以 assetpath.REQUIRED_ASSETS 为唯一事实来源")
check("--include-data-files=" in src, "build.py 会把资源带进产物")

print("=== 3. gui / watermark 的资源接线 ===")
try:
    from tfd_app import gui as gui_mod          # 需要 tkinter
    from tfd_app import watermark as wm_mod
except ImportError as e:                        # pragma: no cover - 无 tkinter 的环境
    print("  skip 本机无 tkinter，跳过资源接线检查：%s" % e)
else:
    check(os.path.isfile(gui_mod.ICON), "gui.ICON -> %s" % gui_mod.ICON)
    check(os.path.isfile(gui_mod.QRCODE), "gui.QRCODE -> %s" % gui_mod.QRCODE)
    check(os.path.isfile(gui_mod.MINIAPP_QRCODE),
          "gui.MINIAPP_QRCODE -> %s" % gui_mod.MINIAPP_QRCODE)
    check(os.path.isfile(wm_mod._IMG_SRC), "watermark 水印图 -> %s" % wm_mod._IMG_SRC)
    check(os.path.isfile(wm_mod._BG_IMG_SRC),
          "watermark 背景水印图 -> %s" % wm_mod._BG_IMG_SRC)

print()
if fails:
    print("==== 资源回归失败 %d 项 ====" % len(fails))
    for f in fails:
        print("  - " + f)
    raise SystemExit(1)
print("==== 资源回归全部通过 ====")
