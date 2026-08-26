# 论文格式医生 · 导师版 (Thesis Format Doctor — Mentor)

把「论文格式医生」改造成**导师专用**的桌面程序：导师上传**学生的论文**与**学校模板**，
软件检查格式问题，生成一份**带批注的论文副本**（只在问题处加 Word 批注气泡、不改学生正文），
可直接转发给学生自行修改；同时也提供「一键修正」生成改好的版本。

- 🪟 **Windows**：双击 `tfd-mentor.exe`
- 🍎 **macOS**：双击 `tfd-mentor`
- 🐧 **Linux**：运行 `tfd-mentor`

**原生图形窗口（Tkinter），不依赖浏览器、完全离线、零第三方依赖、文件不离开你的电脑**。

> 兼容老格式：检测到 `.doc` / WPS `.wps` 时，若本机装有 LibreOffice 会自动转成 `.docx` 再处理；否则提示你先用 Word/WPS「另存为 .docx」。

---

## 它能做什么

| 功能 | 说明 |
| --- | --- |
| 🩺 **格式体检** | 对照学校模板画像，逐项诊断标题样式、三线表、参考文献、页边距、全角标点等，按高危/中危/低危分级 |
| 💬 **批注副本（核心）** | 不改学生正文，只在格式问题处加 Word 批注气泡（标注"这里要改成什么"），导师转发副本给学生自行修改 |
| 🎯 **一键修正** | 需要时代替学生把论文按模板改好，生成「已修正版本」 |
| 📋 **检查报告** | 可导出一份独立的检查报告（Word），汇总全部问题与修改建议 |
| 📐 **抽取模板画像** | 上传学校模板，自动抽取其格式规范（以模板批注为准） |

> 模板驱动：一切格式要求来自学校模板画像，软件不写死任何学校的具体格式值。

---

## 使用方法

1. 下载对应系统的安装包（见 [Releases](../../releases)）。
2. 双击运行（Windows 可能弹出「未知发布者」提示，点「仍要运行」即可——这是没花钱买代码签名证书的正常现象）。
3. 程序弹出原生图形窗口：
   - 选择「学生论文 .docx / .doc / .wps」（必选）
   - 选择「学校模板 .docx」（用于抽取格式画像，必选）
   - 点「生成批注副本」→ 得到一份带批注、未改正文的副本，转发学生；
     或点「一键修正」→ 得到改好的版本。
4. 关闭窗口即退出。

---

## 技术说明

- 界面：原生 **Tkinter** 窗口，无浏览器、无本地服务、纯本地运行。
- 引擎：模板驱动，纯 Python 标准库解析 docx（zipfile + ElementTree），零第三方依赖。
- 打包：PyInstaller `--onefile --noconsole`，见 `build.py` 与 `.github/workflows/build.yml`。
- 授权：试用 2 次 + 激活码解锁（激活码体系与标准版一致）。

### 本地自行构建（可选）

```bash
pip install pyinstaller
python build.py            # 按当前系统构建到 dist/
```

> Linux 构建前需 `sudo apt-get install -y tk tcl`；Windows / macOS 自带 Tk。

### 发版（GitHub Actions 自动构建三平台）

推送一个 `v*` 标签即可触发：

```bash
git tag v1.0.0
git push origin v1.0.0
```

Actions 会在 Windows / macOS / Ubuntu 三个 runner 上分别构建，并把产物发布到 GitHub Release（按平台后缀 `-windows` / `-macos` / `-linux` 区分）。

---

## 目录结构

```
tfd-mentor/
├── main.py                 # 启动入口（检查 Tkinter → 拉起 GUI）
├── build.py                # 跨平台 PyInstaller 构建脚本
├── pyproject.toml
├── tfd_app/
│   ├── gui.py              # 原生 Tkinter 界面（导师视角）
│   ├── engine.py           # 引擎封装层（吃路径、返回文本 + .doc/.wps 自动转换）
│   ├── core/               # 引擎（标准库，零依赖）
│   └── assets/
└── .github/workflows/build.yml
```

---

## 许可

MIT License。品牌「论文格式医生」归属 reedskill（公众号【芦苇不熬夜】）。
