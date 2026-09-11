# AGENTS.md — tfd-mentor（论文格式医生 · 导师版）

> 这是给 Codex / AI 编码助手的**上下文文件**。
> 老板（芦苇）不懂代码 —— 你改完必须交 **diff + 大白话说明**，经「技术专家团」代码把关（fan-zhengma）后才算完成，最终由老板真机验收。

## 一、项目是什么

面向**导师**的桌面工具：导师上传学生论文 + 学校模板，产出**带 Word 批注气泡的论文副本**（只在问题处加批注、**不改学生正文**），可直接转发给学生自行修改；同时也提供一键修正。

- **当前版本**：`VERSION` 文件 = `v1.0.32`（⚠️ 注意有 `v` 前缀，与学生版 `tfd-desktop` 的纯数字格式不同，别搞混）
- 技术同源 `tfd-desktop`，本仓库自带 `tests/`，改动后务必跑

核心功能：格式体检、**批注副本（导师版最大差异点）**、一键修正、批量/文件夹导入（显示"已选择 N 篇论文"）、③可同时导出「批注副本 + 已修正版本」。
另有「总结」——把全组论文问题汇总成逐篇明细 + 共性问题总表。

## 二、技术栈与硬约束（⚠️ 红线）

- **Python ≥ 3.10，零第三方依赖**：docx 处理靠 `zipfile` + XML 手写。
  🔴 **禁止引入任何第三方库** —— 会摧毁"打包极小、跨平台无依赖坑"的优势。
- **界面**：原生 **Tkinter**；**打包**：**Nuitka**（`build.py`）+ NSIS（`installer.nsi`）。
- 🔴 **防杀软红线**：**禁止自写 base64 / XOR 混淆**（`tfd-desktop` 曾因此被 Defender 判 `Sabsik.TE.A!ml`）。要加固走代码签名证书或 PyArmor。
- 🔴 **批注副本禁止改学生正文**：只能新增批注气泡标注"这里要改成什么"。这是导师版产品的立身之本，动了等于产品变质。

## 三、目录与入口

| 路径 | 说明 |
|---|---|
| `main.py` | 程序入口 |
| `reedmentor.py` / `reedmentor_gui.py` | 早期/别称入口，改动前先确认是否仍被打包引用 |
| `tfd_app/gui.py` `engine.py` | Tkinter 界面 / 引擎编排 |
| `tfd_app/core/` | 核心处理模块 |
| `tfd_app/license.py` `trial.py` `watermark.py` | 激活、试用、水印 |
| `tests/test_headless_regression.py` | 无头回归测试 |
| `tests/test_isolation.py` | 隔离性测试（离线/隐私相关） |

## 四、常用命令

```bash
python main.py                              # 本地运行调试
python -m unittest discover -s tests        # 跑回归 + 隔离性测试
python -m pytest tests                      # 若装了 pytest 也可
python build.py                             # 本地打包（正式三平台走 GitHub Actions）
```

## 五、不可破坏的对外承诺（源自 MEMORY.md，违反 = L3 拦截）

1. **论文绝对不上传**：本机处理，仅激活联网校验一次。
2. **模板驱动**：学校模板画像优先级 **批注说明 > 样式定义 > 通用规范**。
3. **一次付费永久使用**：终身 + 一机一码。
4. **批注气泡说明"改了哪、为什么"**，方便对照导师意见。
5. 交付物完整：批注副本 / 修正版 + 报告 + 全组总结。

## 六、改动铁律（每次迭代必守）

1. **小步快跑**：一次只改一件事。
2. **先备份**：改前 `git commit` 或打包留存。
3. **跑测试**：涉及批量/路径/离线的改动必须过 `tests/`（`test_isolation.py` 专治隐私承诺被破坏）。
4. **白话交底**：diff + 大白话说明一起交。
5. **先测后上**，再动生产。
6. **真机验收**：老板用**真实学生论文**（一整组，多个学生）走流程，他说 OK 才算上线。

## 七、本仓库驱动 Codex 的标准命令

```bash
export DEEPSEEK_API_KEY="$(cat '/d/AgentSpace/config/deepseek-key.txt' | tr -d '\r\n')" && \
codex exec --skip-git-repo-check "你要它干的事"                        # 只读分析/问代码
codex exec --skip-git-repo-check -s workspace-write "你要它干的事"    # 真改代码：必须加 -s workspace-write
```

⚠️ **沙箱坑**：Codex 默认沙箱为 `read-only`，**不写盘——所有 shell 命令会被策略拦截**（报 `exec_command failed ... rejected: blocked by policy`）。
**要让它真正落盘改代码，必须显式加 `-s workspace-write`**；只做只读问答时用默认即可。取值：`read-only | workspace-write | danger-full-access`（后者不要碰）。

在本目录运行，改动即落本仓库。API key 走本地 keyfile，**不进聊天、不写进 config**。

## 八、事实来源（冲突以此为准）

- `D:\AgentSpace\MEMORY.md` —— 全局事实卡（功能 / 版本 / 价格 / 联系方式唯一来源）
- `D:\AgentSpace\技术专家团_接手须知.md` —— 技术专家团 briefing
- `D:\AgentSpace\迭代与需求清单.md` —— 需求入口
