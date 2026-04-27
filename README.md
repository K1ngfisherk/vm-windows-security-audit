# VM Windows 安全审查 Skill

[English](README.en.md)

`vm-windows-security-audit` 是一个通用 agent skill 脚手架，用于根据安全检查表审查 VMware 中的 Windows 虚拟机。

它使用 `vmrun`、VMware Tools 和来宾机内的 PyAutoGUI 操作真实 Windows 图形界面，保存原始截图证据，并可按需生成带截图证据的 Excel 检查报告。

## 功能

- 分析 `.xlsx` 安全检查表，并把检查行映射到 GUI 审查动作。
- 打开 `secpol.msc`、`gpedit.msc`、`lusrmgr.msc`、`services.msc`、`eventvwr.msc`、`fsmgmt.msc`、`regedit` 等 Windows 管理工具。
- 按行保存截图证据，例如 `row11_gpedit_rdp_client_connection_encryption_level.png`。
- 当检查项要求图形化验证时，优先使用 GUI 截图，不用命令输出替代。
- 对未命中已知模式的行，支持基于上下文的自适应 GUI 操作。
- 来宾机缺少 Python 或 PyAutoGUI 环境时，可使用仓库内置离线包自动准备审查所需依赖。
- 可选：复制原始表格并把截图直接嵌入 `检查情况` 列。

## 主流程

```text
vmrun + VMware Tools + guest-side PyAutoGUI + Excel
```

VMware MCP 可以作为可选辅助，但不是必要依赖。

## 目录结构

```text
vm-windows-security-audit-skill-sample/
├── SKILL.md
├── README.md
├── README.en.md
├── agents/openai.yaml
├── offline/
│   ├── python-3.9.13-amd64.exe
│   ├── requirements-guest-py39.txt
│   ├── install_guest_offline.ps1
│   ├── MANIFEST.json
│   └── wheelhouse/
├── references/
│   ├── check_patterns.json
│   ├── evidence_rules.md
│   ├── gui_adaptive_rules.md
│   └── troubleshooting.md
└── scripts/
    ├── analyze_checklist.py
    ├── guest_gui_runner.py
    ├── guest_preflight.py
    ├── guest_setup_pyautogui.ps1
    ├── host_orchestrator.ps1
    ├── workbook_embed_excel_com.ps1
    └── workbook_output.py
```

## 环境要求

宿主机：

- Windows 宿主机，安装 VMware Workstation 或兼容的 `vmrun.exe`。
- 来宾机中 VMware Tools 可用。
- 如需高保真嵌图报告，宿主机需要安装 Microsoft Excel。

来宾机：

- Windows 虚拟机，有可交互桌面会话。
- 有权限执行审查的来宾账号。
- Python 和 GUI 自动化依赖可由脚本自动检测和安装。
- 仓库内置 `offline/` 离线包，包含 Python 3.9.13 x64 安装器和 PyAutoGUI 依赖 wheelhouse。
- 在线环境可复用已有 Python 或使用 winget；无网络系统可直接使用 `offline/install_guest_offline.ps1`。

自动准备的来宾依赖包括：

```text
pyautogui
pillow
pyperclip
pygetwindow
```

环境准备脚本：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/guest_setup_pyautogui.ps1
```

无网络来宾机可使用：

```powershell
powershell -ExecutionPolicy Bypass -File offline/install_guest_offline.ps1
```

## 输出规范

```text
<source_workbook_stem>_<task_label>.xlsx
Windows完整检查_<task_label>_证据
Windows完整检查_<task_label>_执行计划.json
```

截图命名：

```text
rowNN_<tool>_<check_key>.png
```

最终证据目录根目录只保留已验收的最终截图。运行日志、stdout/stderr、runner 结果、校验 JSON、联系表、临时脚本和错误诊断截图必须进入 `Windows完整检查_<task_label>_证据/tmp`，任务完成后默认删除该 `tmp` 目录。

已安装补丁检查的最终截图必须停留在“已安装更新”页面，并能看到 Windows 更新/KB 列表区域，不能只截“卸载或更改程序”的软件列表页。

## 说明

- 生成 workbook、超链接、压缩包都是可选行为，只在需要时执行。
- 当前 GUI 导航主要围绕 Windows Server 2012 风格 MMC 窗口整理，其他 Windows 版本或显示缩放可能需要调整。
