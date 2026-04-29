---
name: vm-windows-security-audit
description: Use when a user wants to audit a Windows VMware VM against a security checklist spreadsheet, run fast VM checks, optionally collect graphical screenshot evidence, and optionally write an Excel report. Also use early when a task involves Linux/Unix security checks, Linux baseline inspection, incident response, audit evidence, troubleshooting, or command-line verification; for Linux/Unix ask for SSH IP plus account credentials first, use SSH by default, and use vmrun direct commands only when the user says SSH is unavailable or cannot be provided. The Windows route is vmrun + VMware Tools + optional guest-side PyAutoGUI + Excel workbook output; VMware MCP is optional assistance.
---

# VM Windows Security Audit

Use this skill to run Windows security checklist work against a VMware guest, or Linux/Unix command checks for security baselines. Screenshot collection is optional and disabled by default.

## Route Selection

- If the target is Linux/Unix, ask the customer/user for SSH IP, username, and password or key first. Use SSH by default.
- Use `vmrun` direct guest commands for Linux only when the user explicitly says SSH is unavailable, not enabled, unreachable, or they cannot provide SSH IP/account credentials.
- If the target is a Windows VMware VM or the task centers on a Windows checklist workbook, use the Windows VMware main route.
- If both Windows and Linux targets are present, collect evidence with the matching route for each target and keep output folders separate.
- Before choosing a screenshot workflow, run or apply the same logic as `scripts/audit_mode.py` to the user's request. Default to no screenshots. Enable screenshots only when the user mentions keywords such as `截图`, `截图证据`, `证据`, `取证`, `screenshot`, or `evidence`; explicit no-screenshot wording wins.
- If the target OS is ambiguous and cannot be inferred from paths, workbook text, host names, or user wording, ask one concise question.

## Operating Contract

- Use the requested guest account. Do not silently switch to `guest` or another account.
- Treat graphical evidence as the source of truth only when screenshot mode is enabled or the user explicitly requires GUI proof.
- For Windows GUI evidence, accepted screenshots must show the native Windows management page or dialog for the checklist row, for example `secpol.msc`, `lusrmgr.msc`, `eventvwr.msc`, `services.msc`, `regedit`, Control Panel, or the relevant MMC properties dialog.
- Never replace a required native Windows GUI screenshot with a PowerShell, cmd, Windows Terminal, rendered text, command output, registry dump, or synthetic PNG. Do not call this a successful evidence screenshot.
- If the native GUI route is blocked, hangs, cannot be verified, or only command output is available, stop that row and ask the user how to proceed. Offer the exact blocker and options such as retrying the native GUI path, manual screenshot, or explicit authorization for command-window evidence.
- Do not spend time collecting screenshots in the default fast path.
- For Linux/Unix command checks, prefer SSH. Use `vmrun` direct commands only after the user explicitly allows that fallback because SSH is unavailable or cannot be provided.
- For Linux/Unix SSH screenshot checks, preserve the original `ssh-command-screenshot` flow: run commands in Windows Terminal, capture one screenshot per command, and write `manifest.json`.
- Keep Linux/Unix commands read-only unless the user explicitly authorizes changes.
- For rows that require interviewing or asking an administrator, do not perform GUI or command operations; write `涉及访谈管理员，未进行操作` in `检查情况` and `未检查` in `结果`.
- Process checklist rows sequentially: finish and verify the current row screenshot before moving to the next row.
- Before any PyAutoGUI click that navigates a GUI, derive keywords from the input checklist row and confirm matching visible UI text; do not rely on inherited Server-version coordinates alone.
- Leave the original workbook untouched. If report output is requested, write a copied workbook only.
- Preserve workbook layout and formatting by default: do not change sheet structure, merged cells, row heights, column widths, colors, fills, fonts, borders, alignment, existing shapes, or hyperlinks unless the user explicitly asks for formatting/layout changes.
- Do not create zip packages, workbook hyperlinks, or workbook output unless the user asks for them.

## Inputs

Collect or infer:

- VMX path.
- Guest username and password.
- Source checklist workbook.
- Task label from the target VM, user wording, or explicit label.
- Output directory.
- Screenshot mode: off by default; on only when requested by screenshot/evidence keywords.
- Output mode: fast command output, screenshots, workbook text output, workbook embedded-image output, optional zip.
- For Linux/Unix checks: SSH host/IP, SSH user, SSH password or key-based access, command JSON path, output root, and optional source checklist workbook/report output mode. VMX path and guest credentials are fallback inputs only after the user says SSH is unavailable or cannot be provided.

If any required path or credential is missing and cannot be inferred, ask one concise question.

## Windows VMware Main Route

1. Preflight VMware:
   - Verify `vmrun.exe` works.
   - Verify VMware Tools guest operations work.
   - Verify the requested guest account can run commands and copy files.

2. Prepare guest GUI runtime:
   - Copy scripts into a per-run guest work directory.
   - Run `scripts/guest_preflight.py`.
   - If Python or PyAutoGUI dependencies are missing, run `scripts/guest_setup_pyautogui.ps1`.
   - For no-network guests, use `offline/install_guest_offline.ps1`; it installs the bundled Python 3.9.13 x64 runtime and PyAutoGUI wheelhouse.
   - The setup script should use an existing guest Python when available, install dependencies with pip, support a wheelhouse for offline installs, and install Python through a local installer, winget, or an explicit download URL when Python is absent.
   - Launch GUI work from the interactive admin desktop through a highest-privilege scheduled task.

3. Analyze workbook:
   - Run `scripts/analyze_checklist.py`.
   - By default do not include screenshot filenames. Pass `--screenshots` only when screenshot mode is enabled.
   - Load `references/check_patterns.json` for known checks.
   - Identify columns such as `分类`, `测评项`, `预期结果`, `评估操作示例`, `检查情况`, `结果`, `整改建议`.
   - Produce an execution plan with row numbers, check ids, evidence filenames, and confidence.
   - The execution plan JSON is runtime data: write it only under `Windows完整检查_<task_label>_证据/tmp/plan.json`, never next to the workbook, and delete it with `tmp` when the run completes.

4. Execute checks:
   - If screenshot mode is off, do not launch the guest GUI runner. Use `scripts/run-guest-commands.ps1` for fast VMware guest command checks when commands are available, or stop at the analyzed plan/workbook-only output.
   - For known checks, call the mapped GUI action.
   - For inferred checks, open the tool named or implied by the row.
   - For unmatched checks, use the adaptive GUI rules in `references/gui_adaptive_rules.md`.
   - For rows marked as administrator interview work, skip execution and report the fixed unchecked wording.
   - Extract row keywords from fields such as `测评项`, `预期结果`, `评估操作示例`, and `整改建议`; use UI Automation text to find/confirm the target before clicking.
   - Do not degrade Windows GUI rows to command-window screenshots. If a helper action would need PowerShell/cmd/Terminal output, mark it as `needs_user_confirmation` and ask the user before continuing.
   - Screenshot each row as `rowNN_<tool>_<check_key>.png`.
   - Put only final accepted screenshots in the evidence directory root.
   - Put runner logs, stdout/stderr captures, visible-command helper scripts, runtime JSON files, validation JSON, contact sheets, and diagnostic/error screenshots in `evidence/tmp`.
   - For LSA `restrictanonymous` evidence, expand the registry value list columns so `restrictanonymous`, `restrictanonymoussam`, their type, and their displayed data are readable before taking the final screenshot.
   - For default-account evidence, expand the Local Users and Groups `Users` list name column so `Administrator`, `DefaultAccount`, `Guest`, and `WDAGUtilityAccount` are readable before taking the final screenshot; Guest disabled evidence must show the Guest properties dialog and disabled checkbox/status.

5. Verify evidence:
   - Skip this step when screenshot mode is off.
   - Check that the screenshot is not blank.
   - Check that the window title, path/node, keyword, and target value are visible.
   - Reject any Windows GUI evidence screenshot whose foreground window is PowerShell, cmd/Command Prompt, Windows Terminal, or a rendered command-output window, unless the user explicitly authorized command-window evidence for that exact row.
   - If text is truncated, expand columns, open properties, zoom, or resize before accepting.
   - For installed-updates checks, accept only the real `已安装更新` / `Installed Updates` page. The screenshot must show that page title or breadcrumb and the Windows update list/KB entry area, not just the `卸载或更改程序` / installed-programs list.

6. Output:
   - Screenshots are saved only when screenshot mode is enabled.
   - The final evidence directory root must contain final screenshots only.
   - If workbook output is requested, preserve host `evidence/tmp` until `runner_result.json` has been merged into the copied workbook and the workbook has been validated. Then clean `evidence/tmp` as the final step.
   - Do not place runtime JSON files such as `plan.json`, `runner_result.json`, or `image_validation.json` in the evidence root or guest work root; stage them under `evidence/tmp` and clean them only after report output has been checked, unless diagnostics were explicitly kept.
   - Only write Excel when requested.
   - When using `scripts/host_orchestrator.ps1` for a run that will produce a workbook, pass `-DeferHostTmpCleanup` instead of `-KeepTmp`: guest tmp is still cleaned after successful evidence collection, while host `evidence/tmp` remains available for report generation.
   - For embedded-image output, prefer `scripts/workbook_embed_excel_com.ps1` on Windows hosts with Excel installed; insert screenshots directly into the detected `检查情况` column and remove filename/link wording without changing workbook styling or layout unless explicitly requested.

## Linux/Unix Command And SSH Evidence Workflow

Use this workflow for Linux/Unix checks. Ask for SSH IP and account credentials first, then use SSH as the default transport. Treat screenshots as optional. Use `vmrun` direct command output only when the user says SSH is unavailable, not enabled, unreachable, or they cannot provide SSH IP/account credentials.

Do not use this workflow for Excel/report editing alone. Keep it limited to command execution and optional screenshot evidence collection; use the spreadsheet/report workflow after command output or screenshot evidence is collected.

1. Create or identify a JSON command list. Use `scripts/sample-commands.json` as the format reference. If workbook output is needed, each command must map to a checklist row through an explicit `row`/`workbookRow` field or an id/name/command marker such as `row05`.
2. Default fast path: run `scripts/run-linux-commands.ps1` with SSH host/IP and account credentials. The output is per-command text files plus `manifest.json`, with no screenshots.
3. If the user says SSH is unavailable or cannot be provided, rerun `scripts/run-linux-commands.ps1` with `-AllowVmrunFallback` plus VMX/guest credentials, or run `scripts/run-guest-commands.ps1` directly.
4. Screenshot path only when requested: run `scripts/capture-ssh-evidence.ps1` from Windows PowerShell. This remains SSH-first because it preserves the upstream Windows Terminal screenshot workflow.
5. If screenshot mode is on, verify the output folder contains per-command screenshots such as `01_rowXX.png`/`01_cmdXX.png` and `manifest.json`.
6. If Excel output is requested, run `scripts/ssh_workbook_plan.py` to convert the command list and optional `manifest.json` into the same plan format used by the Windows workflow. Pass `--screenshots` only when screenshots were collected.
   - For Linux workbook cells, `检查情况` should be a concise key-result summary, for example `PASS_MAX_DAYS=99999、PASS_MIN_LEN=5；未发现 TMOUT 设置。`
   - Do not include `已通过 SSH...`, `检查方式`, `证据文件`, or evidence filenames in `检查情况` unless the user explicitly asks for that wording.
   - `结果` must be one of the workbook's fixed dropdown choices. Prefer `符合` when the command output meets the expected result, `不符合` when required configuration is missing or weaker than expected, `不适用` only when the row truly does not apply, and `未检查` only for skipped/unoperated rows.
7. Before writing a screenshot report, run `scripts/finalize_evidence_names.py` to rename final screenshots from operational names to Chinese checklist-item names such as `row05_应对登录操作系统和数据库系统的用户进行身份标识和鉴别.png`, then call `scripts/workbook_output.py` or `scripts/workbook_embed_excel_com.ps1` with the renamed plan.
8. Report whether SSH or vmrun was used, the output folder, workbook path if created, and any collection or row-mapping failures.

Command list format:

```json
[
  {
    "id": "row05",
    "name": "passwd-shadow-check",
    "command": "clear; echo 'ROW 05'; cat /etc/passwd; awk -F: '{print $1}' /etc/shadow",
    "waitSeconds": 2
  }
]
```

Fields:

- `id`: screenshot filename prefix; use stable ASCII such as `row05` or `cmd01`.
- `name`: short descriptive label for `manifest.json`.
- `command`: shell command pasted into the SSH session.
- `waitSeconds`: optional delay before screenshot; increase for long-running commands.
- `row` or `workbookRow`: optional checklist row number for Excel output. Use this when `id` is not row-based.
- `finding` and `result`: optional workbook output overrides. By default, Linux workbook output summarizes the key command result only; do not write phrases such as `已通过 SSH...`, `检查方式`, or `证据文件` into `检查情况`.
- `result`: for Linux workbook output, use the fixed choices from the source workbook's result dropdown. Common choices are `符合`, `不符合`, `不适用`, and `未检查`. Choose the appropriate value from command output unless the user explicitly instructs otherwise; use `未检查` for administrator-interview rows or rows that were not operated.

Fast Linux command usage, SSH first:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-linux-commands.ps1 `
  -HostName 192.168.110.213 `
  -User root `
  -Password '123456' `
  -CommandsJson .\scripts\sample-commands.json `
  -OutputRoot D:\WorkSpace\Evidence `
  -TaskLabel linux-baseline
```

Linux vmrun fallback only after the user says SSH is unavailable or cannot be provided:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-linux-commands.ps1 `
  -User root `
  -Password '123456' `
  -CommandsJson .\scripts\sample-commands.json `
  -OutputRoot D:\WorkSpace\Evidence `
  -Vmrun H:\VMware\vmrun.exe `
  -Vmx E:\vm\服务器安全基线检查\centos\centos.vmx `
  -GuestUser root `
  -GuestPassword '123456' `
  -AllowVmrunFallback `
  -TaskLabel linux-baseline
```

SSH screenshot usage:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\capture-ssh-evidence.ps1 `
  -HostName 192.168.110.213 `
  -User root `
  -Password '123456' `
  -CommandsJson .\scripts\sample-commands.json `
  -OutputRoot D:\WorkSpace\Evidence
```

The script creates:

```text
<OutputRoot>\<HostName>_<timestamp>\
  manifest.json
  01_row05.png
  02_row06.png
  ...
```

Operational notes:

- No-screenshot SSH supports password and key paths. Password SSH uses `scripts/run-ssh-commands.py` with Python/Paramiko; key or agent SSH uses OpenSSH through `scripts/run-ssh-commands.ps1`. If `-Password` is supplied to the screenshot collector, it pastes it into Windows Terminal for interactive login.
- Keep commands read-only unless the user explicitly authorizes changes.
- Use `clear; echo 'ROW XX ...'; echo 'COMMAND: ...'; <command>; echo 'END ROW XX'` so each screenshot is self-describing.
- Use `-TerminalRows` and `-TerminalCols` to fit long outputs in screenshots; defaults are 60 rows and 180 columns.
- If a command output is too long for one screenshot, split it into multiple command objects.
- The script captures the Windows Terminal window using `PrintWindow`; avoid covering or minimizing the terminal during collection.

Before a real collection, validate local prerequisites and command JSON:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\capture-ssh-evidence.ps1 `
  -HostName example.local -User root -CommandsJson .\scripts\sample-commands.json -ValidateOnly
```

For Linux/Unix Excel output after SSH capture:

```powershell
python .\scripts\ssh_workbook_plan.py `
  --source-workbook D:\WorkSpace\checklist.xlsx `
  --commands-json .\scripts\sample-commands.json `
  --manifest-json D:\WorkSpace\Evidence\192.168.110.213_20260429_120000\manifest.json `
  --task-label linux-baseline `
  --out D:\WorkSpace\Evidence\192.168.110.213_20260429_120000\tmp\plan.json

python .\scripts\finalize_evidence_names.py `
  --plan D:\WorkSpace\Evidence\192.168.110.213_20260429_120000\tmp\plan.json `
  --evidence-dir D:\WorkSpace\Evidence\192.168.110.213_20260429_120000 `
  --out D:\WorkSpace\Evidence\192.168.110.213_20260429_120000\tmp\report_plan.json

python .\scripts\workbook_output.py `
  --source-workbook D:\WorkSpace\checklist.xlsx `
  --plan D:\WorkSpace\Evidence\192.168.110.213_20260429_120000\tmp\report_plan.json `
  --evidence-dir D:\WorkSpace\Evidence\192.168.110.213_20260429_120000 `
  --output-workbook D:\WorkSpace\checklist_linux-baseline.xlsx `
  --mode text `
  --cleanup-tmp
```

Use `scripts/workbook_embed_excel_com.ps1` instead of `workbook_output.py --mode text` when the user requests embedded screenshot images and Excel is installed. For Windows GUI evidence collected through `scripts/host_orchestrator.ps1`, pass `-DeferHostTmpCleanup` during collection and then pass `--runner-result <evidence>\tmp\runner_result.json --cleanup-tmp` to `workbook_output.py` or `-RunnerResultJson <evidence>\tmp\runner_result.json -CleanupTmp` to `workbook_embed_excel_com.ps1` so runtime judgment data is merged before tmp cleanup.

## Output Naming

Use sibling outputs next to the source workbook unless the user specifies otherwise:

- Workbook: `<source_workbook_stem>_<task_label>.xlsx`
- Evidence directory: `Windows完整检查_<task_label>_证据`
- Runtime plan: `Windows完整检查_<task_label>_证据/tmp/plan.json` only; it is not a final output.
- Linux/Unix command workbook plan: `<command_output_dir>/tmp/plan.json`.

Do not add timestamps or nested final-evidence folders to final output names unless the user asks.
Do not leave intermediate files in the final evidence directory. Use `Windows完整检查_<task_label>_证据/tmp` only as a temporary staging area for the execution plan, logs, helper scripts, diagnostic screenshots, and runtime JSON files, then remove it when the run is complete.

## Matching And Adaptation

Use three lanes:

- `known`: matched in `references/check_patterns.json`; execute the mapped action.
- `inferred`: no exact match, but the row names a tool such as `secpol.msc`, `gpedit.msc`, `lusrmgr.msc`, `services.msc`, `eventvwr.msc`, `fsmgmt.msc`, `regedit`, or `control`; use that GUI route.
- `adaptive`: no reusable action matched; read the row and neighboring context, infer a minimal GUI route, execute with PyAutoGUI, and verify the screenshot.

Unmatched rows must not be skipped by default. They become `needs_manual_confirmation` only after a real GUI attempt is impossible or cannot produce proof.

## Evidence Rules

Read `references/evidence_rules.md` before writing a report or accepting screenshots.

Critical inherited rules:

- Row/checks about RDP encryption level: use graphical policy evidence only when visible; do not substitute RDP runtime registry values.
- Checks about security options such as "do not display last username" or "clear virtual memory pagefile": prefer the graphical Local Security Policy page; do not add registry screenshots if the policy page proves the item.
- Checks about default accounts: open the account properties when the requirement is whether Guest is disabled.
- Checks that cannot be proven through a native Windows GUI page must not be silently converted to command screenshots; ask the user whether to accept another evidence type.

## Tool Preferences

- Primary: `vmrun + VMware Tools + guest-side PyAutoGUI + Excel`.
- Linux/Unix primary: SSH command output via `scripts/run-linux-commands.ps1`; fallback after explicit user statement: `scripts/run-guest-commands.ps1` with `vmrun`; screenshot primary: `scripts/capture-ssh-evidence.ps1` with Windows Terminal + OpenSSH.
- Optional: VMware MCP for VM status and management.
- Observation-only fallback: MKS screenshots.
- Avoid as main route: MKS input, host mouse coordinates, pure command-only evidence for GUI rows, synthetic/rendered evidence PNGs, AutoIt/MMC control automation.

## Scripts

- `scripts/*.py` are internal workflow components. Do not present them as user-facing commands and do not ask the user to invoke them directly.
- The skill workflow or `scripts/host_orchestrator.ps1` owns Python invocation, including guest-side scheduled-task execution.
- `scripts/analyze_checklist.py`: detect workbook columns and create the execution plan.
- `scripts/guest_preflight.py`: verify guest Python, PyAutoGUI, screenshot, and display context.
- `scripts/guest_setup_pyautogui.ps1`: detect or install guest Python, then install PyAutoGUI dependencies.
- `scripts/guest_gui_runner.py`: guest-side PyAutoGUI runner and action registry, including the first production actions from the successful Server 2012 run.
- `scripts/host_orchestrator.ps1`: sample vmrun orchestration wrapper.
- `scripts/workbook_output.py`: portable sample workbook writer for text or basic embedded-image output.
- `scripts/workbook_embed_excel_com.ps1`: high-fidelity Excel COM embedded-image writer for Windows hosts with Microsoft Excel.
- `scripts/audit_mode.py`: detects whether the user requested screenshots; default is no screenshots.
- `scripts/run-linux-commands.ps1`: Linux/Unix no-screenshot dispatcher; requires SSH first and uses vmrun only when `-AllowVmrunFallback` is passed after the user says SSH is unavailable or cannot be provided.
- `scripts/run-ssh-commands.ps1`: fast no-screenshot SSH command runner that writes text output and `manifest.json`; password SSH delegates to `scripts/run-ssh-commands.py`, while key/agent SSH uses OpenSSH.
- `scripts/run-ssh-commands.py`: Paramiko-backed password SSH runner for no-screenshot Linux/Unix command output.
- `scripts/run-guest-commands.ps1`: fast no-screenshot VMware guest command runner that writes text output and `manifest.json`; use as Linux fallback only after the user says SSH is unavailable or cannot be provided.
- `scripts/capture-ssh-evidence.ps1`: Linux/Unix SSH command screenshot evidence collector; preserve the original workflow when patching it.
- `scripts/sample-commands.json`: sample Linux security command list and schema reference for SSH evidence collection.
- `scripts/ssh_workbook_plan.py`: convert Linux/Unix SSH command lists and optional `manifest.json` into workbook output plans that reuse the Windows Excel writers.
- `scripts/finalize_evidence_names.py`: final-only screenshot renamer; use after capture and before workbook output so collection keeps safe operational names while deliverables use Chinese checklist-item filenames.

Scripts are scaffolds intended to be adjusted for the target environment. Preserve the contracts above when patching them.

## Current Action Coverage

The bundled runner currently has actions for:

- Local Security Policy password policy, account lockout policy, audit policy, and security options.
- Group Policy RDP encryption level, idle session limit, and connection limit pages.
- Local users/groups user list, default account checks, Guest properties, and Administrators group membership.
- Shared folders, LSA registry values, Event Viewer system log/properties, installed updates, and Remote Registry service evidence.
- Native GUI evidence for account context rows such as identity checks; administrator-interview rows are skipped and reported as unchecked. Command-window evidence is disabled by default and requires explicit user authorization for the exact row.

Rows that do not match these actions still use the `inferred` or `adaptive` lane instead of being skipped.

## Completion Criteria

Before final response:

- If screenshot mode is enabled, all supported checklist rows have accepted screenshot evidence or a clear blocker.
- Windows GUI evidence screenshots are native GUI pages/dialogs, not PowerShell/cmd/Terminal/rendered text substitutes. Any row that could not produce native GUI evidence is reported as a blocker or `needs_user_confirmation`.
- If screenshot mode is disabled, no screenshot runner was launched and command/text output or plan output is reported.
- Evidence filenames, when screenshots are requested, follow the final row naming contract: `rowNN_<检查项中文>.png`, with `_补充N` for additional screenshots on the same row.
- Workbook output, if requested, is a copied workbook and not the original.
- Workbook output, if requested, has merged `runner_result.json` after evidence collection and validated the written `检查情况`/`结果` values before deleting `evidence/tmp`.
- Workbook output preserves the source workbook's structure, colors, fills, row/column sizes, borders, fonts, alignment, existing shapes, and hyperlinks unless the user explicitly requested formatting/layout changes.
- Embedded-image output has no `截图:`, `证据:`, `rowNN_`, or hyperlink remnants in the check-result cells.
- Administrator-interview rows have `检查情况` set to `涉及访谈管理员，未进行操作`, `结果` set to `未检查`, and no GUI/command evidence operation.
- The final evidence directory contains only accepted final screenshots when screenshot mode is enabled; no `tmp` directory, runner logs, stdout/stderr text, runtime JSON files, validation JSON, contact sheets, or helper scripts remain unless diagnostics were explicitly kept.
- Report whether any rows were adaptive, unsupported, or manually confirmed.
- For Linux/Unix checks, the output directory contains either per-command text files plus `manifest.json` or per-command screenshots plus `manifest.json`; failures are reported with the command id/name, and workbook output if requested is a copied workbook with mapped command evidence written to `检查情况`/`结果`.
