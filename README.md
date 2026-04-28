# VM Windows 安全审查 Skill

[English](README.en.md)

`vm-windows-security-audit` 用于根据 Windows 安全检查表审查 VMware 虚拟机。它会读取 `.xlsx` 检查表，在虚拟机中打开对应的 Windows 管理界面，逐项采集截图证据，并可按需生成带截图的 Excel 报告。

## 适用场景

- Windows Server 安全基线检查。
- 需要按照表格逐项截图取证的审查任务。
- VMware Workstation 虚拟机中的 Windows 系统检查。
- 需要保留原始 GUI 证据，而不是只导出命令结果的任务。

## 功能概览

- 自动分析 `.xlsx` 检查表，识别检查项、预期结果和操作说明。
- 支持常见 Windows 管理界面，包括本地安全策略、组策略、本地用户和组、服务、事件查看器、共享文件夹、注册表和控制面板。
- 按检查表行号保存截图证据，例如 `row11_gpedit_rdp_client_connection_encryption_level.png`。
- 对截图进行基础有效性和页面内容检查，减少截图页面与检查项不匹配的问题。
- 支持离线准备虚拟机中的 Python 和 GUI 自动化依赖。
- 可选生成 Excel 报告，把截图嵌入到检查结果列。

## 使用前准备

宿主机需要：

- Windows 系统。
- VMware Workstation 或兼容的 `vmrun.exe`。
- 可访问目标虚拟机的 `.vmx` 文件。
- 可访问源检查表 `.xlsx`。
- 如果需要生成嵌图版 Excel 报告，需要安装 Microsoft Excel。

虚拟机需要：

- Windows 系统。
- VMware Tools 正常运行。
- 已登录到桌面。
- 提供用于检查的 guest 用户名和密码。

## 工作流程

1. 读取检查表并生成检查计划。
2. 检查虚拟机环境和桌面状态。
3. 准备虚拟机中的运行环境。
4. 逐项打开对应的 Windows 管理界面。
5. 采集并校验截图证据。
6. 保存最终截图。
7. 如果用户需要，生成带截图的 Excel 报告。

## 输出内容

默认输出在源检查表同级目录，文件名会包含任务标签。

```text
Windows完整检查_<task_label>_证据\
<source_workbook_stem>_<task_label>.xlsx
```

证据截图命名格式：

```text
rowNN_<tool>_<check_key>.png
```

最终证据目录中主要包含可用于交付的截图文件。报告文件只在用户要求生成时创建。

## 当前覆盖范围

已内置的常见检查包括：

- 本地安全策略：密码策略、账户锁定策略、审核策略、安全选项。
- 组策略：远程桌面加密级别、空闲会话限制、连接数限制。
- 本地用户和组：用户列表、默认账户、Guest 属性、Administrators 成员。
- 服务：Remote Registry 等服务配置。
- 事件查看器：System 日志及属性页。
- 共享文件夹、LSA 注册表项、已安装更新页面。
- 身份信息、管理员认证方式等需要可见窗口展示的辅助检查。

## 目录结构

```text
vm-windows-security-audit/
├── SKILL.md
├── README.md
├── README.en.md
├── agents/
├── offline/
├── references/
└── scripts/
```

## 说明

- 原始检查表不会被直接修改。
- Excel 报告、超链接和压缩包属于可选输出。
- 不同 Windows 版本、语言和显示缩放可能导致部分 GUI 路径需要适配。
