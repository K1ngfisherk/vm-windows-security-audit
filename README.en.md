# VM Windows Security Audit Skill

[中文](README.md)

`vm-windows-security-audit` audits Windows VMware virtual machines against Windows security checklist spreadsheets and also supports Linux/Unix command checks. It defaults to the faster no-screenshot path; screenshots are collected only when the user mentions keywords such as screenshot, evidence, or capture evidence. For Linux/Unix, ask the customer for SSH IP plus account credentials first and use SSH by default; use `vmrun` only when the user says SSH is unavailable or cannot be provided.

## Use Cases

- Windows Server security baseline checks.
- Linux/Unix security baseline checks, host inspection, and incident-response verification.
- Fast audit tasks that do not require screenshots.
- Spreadsheet-driven audit tasks that require screenshot evidence.
- Windows system checks inside VMware Workstation virtual machines.
- Reviews where original GUI evidence is preferred over command-only output.
- Tasks where SSH command output needs screenshot evidence.

## Features

- Analyze `.xlsx` checklists and identify check items, expected results, and operation guidance.
- Support common Windows management surfaces, including Local Security Policy, Group Policy, Local Users and Groups, Services, Event Viewer, Shared Folders, Registry Editor, and Control Panel.
- Windows screenshot evidence must be native GUI pages or dialogs; PowerShell, cmd, Windows Terminal, command output, or rendered PNG substitutes are not accepted. If native GUI capture is blocked, the workflow stops and asks the user.
- Use `scripts/audit_mode.py` to decide whether screenshots are requested; default is no screenshots.
- For Linux/Unix without screenshots, use `scripts/run-linux-commands.ps1` with SSH first; `vmrun` fallback requires the user to say SSH is unavailable or they cannot provide IP/account/password.
- When screenshots are requested, save row-level evidence with stable names such as `row11_gpedit_rdp_client_connection_encryption_level.png`.
- Perform basic screenshot usability and page-content checks to reduce mismatched evidence.
- Support offline preparation of Python and GUI automation dependencies inside the VM.
- Optionally generate an Excel report with screenshots embedded into the result column.
- When Linux/Unix checks are detected, ask the customer for SSH IP, account, and password/key first. Default to SSH command text output plus `manifest.json`; if the customer says SSH is unavailable or cannot provide connection details, use VMware Tools/`vmrun` direct commands. Use Windows Terminal + OpenSSH screenshots only when screenshots/evidence are requested.
- Linux/Unix SSH evidence can also produce a copied Excel workbook; commands map back to checklist rows through `row`/`workbookRow` or markers such as `row05`.
- Before delivery, screenshots are renamed to Chinese checklist-item filenames, for example `row05_应对登录操作系统和数据库系统的用户进行身份标识和鉴别.png`.
- Administrator-interview checklist rows do not run GUI or command operations; the report writes `涉及访谈管理员，未进行操作` in `检查情况` and `未检查` in `结果`.

## Before You Start

Host requirements:

- Windows host.
- VMware Workstation or compatible `vmrun.exe`.
- Access to the target VM `.vmx` file.
- Access to the source checklist `.xlsx`.
- Microsoft Excel if an embedded-image Excel report is required.
- Linux/Unix no-screenshot SSH supports passwords and keys: the password path uses Python/Paramiko, while key or agent authentication uses OpenSSH. Screenshot evidence needs Windows Terminal (`wt.exe`), OpenSSH client (`ssh.exe`), network access to the SSH target, and valid credentials or key-based access.

Guest requirements:

- Windows guest system.
- VMware Tools running.
- Logged-in desktop session.
- Guest username and password for the audit.

## Workflow

Windows VMware checks:

1. Read the checklist and build an audit plan.
2. Check the VM environment and desktop state.
3. Prepare the runtime inside the guest.
4. Open the corresponding Windows management interface for each row.
5. Do not capture screenshots by default; if screenshot/evidence keywords are present, capture and validate screenshot evidence.
6. Save final screenshots or no-screenshot command output.
7. Generate an Excel report if requested.

Linux/Unix SSH checks:

1. Prepare a command JSON file; use `scripts/sample-commands.json` as the format reference.
2. By default, run `scripts/run-linux-commands.ps1` over SSH and write per-command `.txt` outputs plus `manifest.json`.
3. Only when the user says SSH is unavailable or cannot provide IP/account/password, rerun with `-AllowVmrunFallback` to use `vmrun`.
4. If screenshots/evidence are requested, prefer `scripts/capture-ssh-evidence.ps1` to preserve the SSH screenshot flow.
5. In screenshot mode, verify screenshots and `manifest.json` in the output folder.
6. If workbook output is needed, run `scripts/ssh_workbook_plan.py`; in screenshot mode, run `scripts/finalize_evidence_names.py` for final Chinese filenames, then reuse `scripts/workbook_output.py` or `scripts/workbook_embed_excel_com.ps1` to write a copied workbook. Reports that need `runner_result.json` keep host `tmp` until the workbook is written and validated, then remove it.

## Outputs

By default, outputs are written next to the source workbook and include the task label in their names.

```text
Windows完整检查_<task_label>_证据\
<source_workbook_stem>_<task_label>.xlsx
```

Linux/Unix SSH evidence output directory:

```text
<OutputRoot>\<HostName>_<timestamp>\
```

Screenshot naming format:

```text
rowNN_<Chinese checklist item>.png
```

The final evidence directory mainly contains deliverable screenshot files; no-screenshot mode outputs per-command text files and `manifest.json`. Report workbooks are created only when requested.
When Linux/Unix checks need Excel output, the workbook is a copy of the source checklist. `检查情况` contains only concise key command results, not process wording such as `已通过 SSH`, `检查方式`, or `证据文件`; `结果` is selected from the workbook's fixed choices, usually `符合`, `不符合`, `不适用`, and `未检查`.

## Current Coverage

Built-in checks currently cover:

- Local Security Policy: password policy, account lockout policy, audit policy, and security options.
- Group Policy: Remote Desktop encryption level, idle session limit, and connection limit.
- Local Users and Groups: user lists, default accounts, Guest properties, and Administrators membership.
- Services: Remote Registry and other service settings.
- Event Viewer: System log and properties pages.
- Shared Folders, LSA registry values, and Installed Updates.
- Helper checks that display identity details in a visible window; administrator-interview rows are written as not operated and not checked.
- Linux/Unix command output or screenshot evidence mapped back into checklist rows through the shared Excel output scripts.

## Layout

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

## Notes

- The source checklist is not modified directly.
- Excel reports, hyperlinks, and zip packages are optional outputs.
- Some GUI paths may need adaptation for different Windows versions, languages, or display scaling.
