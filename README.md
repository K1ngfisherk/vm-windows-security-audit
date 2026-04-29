# VM Windows 安全审查 Skill

[English](README.en.md)

`vm-windows-security-audit` 用于根据 Windows 安全检查表审查 VMware 虚拟机，也支持 Linux/Unix 主机执行命令检查。默认走更快的无截图模式；只有用户提到“截图、证据、取证、screenshot、evidence”等关键词时，才启用截图采集。Linux/Unix 先问客户要 SSH IP、账户和密码/密钥并优先 SSH；只有用户明确说没有 SSH 或提供不了连接信息时，才走 `vmrun`。

## 适用场景

- Windows Server 安全基线检查。
- Linux/Unix 安全基线检查、主机核查、应急响应核验。
- 需要快速完成、不要求截图证据的审查任务。
- 需要按照表格逐项截图取证的审查任务。
- VMware Workstation 虚拟机中的 Windows 系统检查。
- 需要保留原始 GUI 证据，而不是只导出命令结果的任务。
- 需要通过 SSH 命令输出保留截图证据的任务。

## 功能概览

- 自动分析 `.xlsx` 检查表，识别检查项、预期结果和操作说明。
- 支持常见 Windows 管理界面，包括本地安全策略、组策略、本地用户和组、服务、事件查看器、共享文件夹、注册表和控制面板。
- Windows 截图证据必须是原生 GUI 页面或对话框；不能用 PowerShell、cmd、Windows Terminal、命令输出或渲染 PNG 替代。原生页面无法采集时会停止并询问用户。
- 使用 `scripts/audit_mode.py` 判断是否需要截图；默认不截图。
- Linux/Unix 无截图时用 `scripts/run-linux-commands.ps1`，先 SSH；`vmrun` 需要用户明确说没有 SSH 或给不了 IP/账户/密码后才允许。
- 需要截图时按检查表行号保存截图证据，例如 `row11_gpedit_rdp_client_connection_encryption_level.png`。
- 对截图进行基础有效性和页面内容检查，减少截图页面与检查项不匹配的问题。
- 支持离线准备虚拟机中的 Python 和 GUI 自动化依赖。
- 可选生成 Excel 报告，把截图嵌入到检查结果列。
- 识别到 Linux/Unix 检测时，先向客户确认 SSH IP、账户和密码/密钥，默认用 SSH 输出命令文本与 `manifest.json`；客户明确说没有 SSH 或提供不了连接信息时，再用 VMware Tools/`vmrun` 直接执行命令；仅在用户要求截图/证据时使用 Windows Terminal + OpenSSH 生成每条命令的截图。
- Linux/Unix SSH 取证也可输出复制版 Excel；命令可通过 `row`/`workbookRow` 或 `row05` 这类标记映射回检查表行。
- 最终交付前会把截图重命名为中文检查项名称，例如 `row05_应对登录操作系统和数据库系统的用户进行身份标识和鉴别.png`。
- 涉及访谈管理员的检查项不会执行 GUI 或命令操作，会在 `检查情况` 写入 `涉及访谈管理员，未进行操作`，在 `结果` 写入 `未检查`。

## 使用前准备

宿主机需要：

- Windows 系统。
- VMware Workstation 或兼容的 `vmrun.exe`。
- 可访问目标虚拟机的 `.vmx` 文件。
- 可访问源检查表 `.xlsx`。
- 如果需要生成嵌图版 Excel 报告，需要安装 Microsoft Excel。
- Linux/Unix 无截图 SSH 支持密码和密钥：密码路径使用 Python/Paramiko，密钥或 agent 路径使用 OpenSSH。截图取证需要 Windows Terminal（`wt.exe`）、OpenSSH 客户端（`ssh.exe`）、目标 SSH 连通性和有效凭据或密钥。

虚拟机需要：

- Windows 系统。
- VMware Tools 正常运行。
- 已登录到桌面。
- 提供用于检查的 guest 用户名和密码。

## 工作流程

Windows VMware 检查：

1. 读取检查表并生成检查计划。
2. 检查虚拟机环境和桌面状态。
3. 准备虚拟机中的运行环境。
4. 逐项打开对应的 Windows 管理界面。
5. 默认不采集截图；如果命中截图/证据关键词，再采集并校验截图证据。
6. 保存最终截图或无截图命令输出。
7. 如果用户需要，生成 Excel 报告。

Linux/Unix SSH 检查：

1. 准备命令 JSON，可参考 `scripts/sample-commands.json`。
2. 默认运行 `scripts/run-linux-commands.ps1`，使用 SSH 输出每条命令的 `.txt` 和 `manifest.json`。
3. 只有用户说没有 SSH 或提供不了 IP/账户/密码时，才带 `-AllowVmrunFallback` 走 `vmrun`。
4. 用户要求截图/证据时，优先改用 `scripts/capture-ssh-evidence.ps1` 保留 SSH 截图流程。
5. 截图模式下校验输出目录中的截图和 `manifest.json`。
6. 如果需要表格输出，运行 `scripts/ssh_workbook_plan.py` 生成计划；截图模式再用 `scripts/finalize_evidence_names.py` 最终中文重命名，最后用 `scripts/workbook_output.py` 或 `scripts/workbook_embed_excel_com.ps1` 输出复制版 Excel。需要依赖 `runner_result.json` 的报告会先保留本机 `tmp`，写表并校对通过后再清理。

## 输出内容

默认输出在源检查表同级目录，文件名会包含任务标签。

```text
Windows完整检查_<task_label>_证据\
<source_workbook_stem>_<task_label>.xlsx
```

Linux/Unix SSH 证据输出目录：

```text
<OutputRoot>\<HostName>_<timestamp>\
```

证据截图命名格式：

```text
rowNN_<检查项中文>.png
```

最终证据目录中主要包含可用于交付的截图文件；无截图模式则输出逐条命令文本和 `manifest.json`。报告文件只在用户要求生成时创建。
如果 Linux/Unix 检查需要 Excel，输出为原检查表的复制版。`检查情况` 只写命令关键结果和简要描述，不写“已通过 SSH”“检查方式”“证据文件”等过程性文字；`结果` 从原表固定选项中选择，通常为 `符合`、`不符合`、`不适用`、`未检查`。

## 当前覆盖范围

已内置的常见检查包括：

- 本地安全策略：密码策略、账户锁定策略、审核策略、安全选项。
- 组策略：远程桌面加密级别、空闲会话限制、连接数限制。
- 本地用户和组：用户列表、默认账户、Guest 属性、Administrators 成员。
- 服务：Remote Registry 等服务配置。
- 事件查看器：System 日志及属性页。
- 共享文件夹、LSA 注册表项、已安装更新页面。
- 身份信息等需要可见窗口展示的辅助检查；访谈管理员类检查项按未操作、未检查写入报告。
- Linux/Unix 命令输出或截图结果映射回检查表行并复用同一套 Excel 输出脚本。

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
