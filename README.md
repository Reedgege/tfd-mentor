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
| 📂 **批量 / 文件夹导入** | 支持逐个多选，也能一键「选择文件夹」自动收齐其中的全部 Word 文档（.docx / .doc / .wps），下方实时显示「已选择 N 篇论文」 |
| 🔁 **①+② 都要** | 第③步可同时选「批注副本 + 已修正版本」，一次跑完导出两套，不重来 |
| 📋 **检查报告** | 可导出一份独立的检查报告（Word），汇总全部问题与修改建议 |
| 📐 **抽取模板画像** | 上传学校模板，自动抽取其格式规范（以模板批注为准） |

> 模板驱动：一切格式要求来自学校模板画像，软件不写死任何学校的具体格式值。

---

## 使用方法

1. 下载对应系统的安装包（见 [Releases](../../releases)）。
2. 双击运行（Windows 可能弹出「未知发布者」提示，点「仍要运行」即可——这是没花钱买代码签名证书的正常现象）。
3. 程序弹出原生图形窗口：
   - **导入论文（必选）**：点「选择论文…」逐个或批量选文件；若学生论文都在同一文件夹，
     直接点「选择文件夹…」即可自动收齐其中的所有 Word 文档，下方实时显示「已选择 N 篇论文」。
   - **选择学校模板（推荐）**：用于抽取格式画像，越贴近学校要求；可勾选「🔒 记住此模板」，
     下次打开自动载入，免重复导入。
   - **第③步选交付方式**：
     - ① 只批注·不改原稿（生成带批注副本，转发学生自行修改）
     - ② 一键修正·直接改好（生成改好的版本）
     - ③ ①+② 都要（一次跑完，同时导出「批注副本」与「已修正版本」）
4. 关闭窗口即退出。

---

## 技术说明

- 界面：原生 **Tkinter** 窗口，无浏览器、无本地服务、纯本地运行。
- 引擎：模板驱动，纯 Python 标准库解析 docx（zipfile + ElementTree），零第三方依赖。
- 打包：Nuitka 编译为原生二进制（`--standalone`），再经 NSIS 打安装包，见 `build.py` 与 `.github/workflows/build.yml`。
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
├── build.py                # 跨平台 Nuitka 构建脚本（产出 standalone 文件夹 + NSIS 安装包）
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

---

## 关于杀软误报（重要）

本程序是**纯 Python + Tkinter 开源项目**，用 **Nuitka 编译为原生二进制**再经 NSIS 打安装包。**个别杀软（尤其 Windows Defender）可能误报为病毒**——这是 Python 打包软件的普遍现象，并非程序含恶意代码：

- 误报根因：早期用 PyInstaller 单文件（onefile）模式，启动时临时解压到 `%TEMP%\_MEIxxxx\` 再运行（典型病毒投放器行为），命中杀软启发式。
- **v1.0.26 起改用 Nuitka + 安装包彻底解决**：Nuitka 把 Python 编译成 C → 原生机器码，产物在 Defender 眼里等同普通 C++ 程序，**运行时不自解压**，实测报毒率仅 2/71；再套一层标准 NSIS 安装包（所有正规软件的做法），用户双击安装、桌面/开始菜单得一个图标，背后文件夹不可见。
- 若个别杀软仍拦截：在 Defender 弹窗点「**更多信息 → 仍要运行**」即可；如需彻底解决，可在 [Microsoft 安全智能提交页](https://www.microsoft.com/en-us/wdsi/filesubmission) 提交安装包申诉（选 "Should not be detected"），通常 1-5 个工作日加入白名单。
- 你也可以在 [VirusTotal](https://www.virustotal.com) 上传安装包自查：正常情况 0~2 / 72 引擎报毒，均为同类启发式误报。
