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
- If the target is a database service running in Docker, use the Docker database route. For MySQL/MariaDB containers, collect Docker metadata with `docker inspect`, run service checks through `docker exec`, and keep container evidence separate from host/VM evidence.
- Before choosing a screenshot workflow, run or apply the same logic as `scripts/audit_mode.py` to the user's request. Default to no screenshots. Enable screenshots only when the user mentions keywords such as `截图`, `截图证据`, `证据`, `取证`, `screenshot`, or `evidence`; explicit no-screenshot wording wins. Treat screenshot collection and report embedding as separate choices: screenshot keywords alone mean standalone auxiliary evidence, and images go into a workbook/document only when the user asks to put/insert/embed screenshots there.
- If the target OS is ambiguous and cannot be inferred from paths, workbook text, host names, or user wording, ask one concise question.

## Operating Contract

- Use the requested guest account. Do not silently switch to `guest` or another account.
- For Windows Server GUI screenshot evidence, require the actual logged-in Windows desktop account before running. Pass it explicitly as `-InteractiveGuestUser`; if it was not provided, ask the user for it instead of falling back to `Guest`, `guest`, or an inferred account.
- Treat graphical evidence as the source of truth only when screenshot mode is enabled or the user explicitly requires GUI proof.
- For Windows GUI evidence, accepted screenshots must show the native Windows management page or dialog for the checklist row, for example `secpol.msc`, `lusrmgr.msc`, `eventvwr.msc`, `services.msc`, `regedit`, Control Panel, or the relevant MMC properties dialog.
- Never replace a required native Windows GUI screenshot with a PowerShell, cmd, Windows Terminal, rendered text, command output, registry dump, or synthetic PNG. Do not call this a successful evidence screenshot.
- If the native GUI route is blocked, hangs, cannot be verified, or only command output is available, stop that row and ask the user how to proceed. Offer the exact blocker and options such as retrying the native GUI path, manual screenshot, or explicit authorization for command-window evidence.
- Do not spend time collecting screenshots in the default fast path.
- For Linux/Unix command checks, prefer SSH. Use `vmrun` direct commands only after the user explicitly allows that fallback because SSH is unavailable or cannot be provided.
- For Linux/Unix SSH screenshot checks, preserve the original `ssh-command-screenshot` flow: run commands in Windows Terminal, capture one screenshot per command, and write `manifest.json`.
- For Docker/database checks, distinguish evidence types explicitly:
  - `original_screenshots`: real Windows Terminal, VM console, remote desktop, or system screenshots.
  - `screenshots`: rendered command-output images, if any. These are auxiliary renderings and must not be called original screenshots.
- When the user asks for original screenshots, never substitute rendered text images, synthetic PNGs, or command-output renderings. Capture the real terminal, desktop, remote session, or VM window.
- For long Docker/database command output, filter or split to the fields that prove the conclusion. Do not force unreadable full output into one screenshot.
- Keep Linux/Unix commands read-only unless the user explicitly authorizes changes.
- For rows that require interviewing or asking an administrator, do not perform GUI or command operations; write `涉及访谈管理员，未进行操作` in `检查情况` and `未检查` in `结果`.
- Process checklist rows sequentially: finish and verify the current row screenshot before moving to the next row.
- Before any PyAutoGUI click that navigates a GUI, derive keywords from the input checklist row and confirm matching visible UI text; do not rely on inherited Server-version coordinates alone.
- Leave the original workbook untouched. If report output is requested, write a copied workbook only.
- Preserve workbook layout and formatting by default: do not change sheet structure, merged cells, row heights, column widths, colors, fills, fonts, borders, alignment, existing shapes, or hyperlinks unless the user explicitly asks for formatting/layout changes.
- For baseline workbook output, only modify the `检查情况` and `结果` columns by default. Do not modify, clear, restyle, overwrite, hide, resize, or otherwise touch `整改建议`.
- Treat `整改建议` as read-only context only: it may be read to understand a checklist row, but its cell values and visual style must remain exactly as in the source workbook.
- Do not write `证据: xxx.png`, screenshot filenames, image paths, or raw evidence paths into result cells unless the user explicitly asks for evidence paths as table text.
- Put evidence paths, screenshot type descriptions, raw-output folders, and command mappings in `manifest.json` or a separate `检查说明.txt`, not in checklist result cells.
- Do not create zip packages, workbook hyperlinks, or workbook output unless the user asks for them.

## Inputs

Collect or infer:

- VMX path.
- Guest username and password.
- For Windows GUI screenshots: actual logged-in Windows desktop username and password for the interactive scheduled task. This is usually the same administrator account, but it must still be explicit; do not infer or fall back to Guest.
- Source checklist workbook.
- Task label from the target VM, user wording, or explicit label.
- Output directory.
- Screenshot mode: off by default; on only when requested by screenshot/evidence keywords.
- Output mode: fast command output, standalone screenshot evidence, workbook/document text output, workbook/document embedded-image output, optional zip.
- For Linux/Unix checks: SSH host/IP, SSH user, SSH password or key-based access, command JSON path, output root, and optional source checklist workbook/report output mode. VMX path and guest credentials are fallback inputs only after the user says SSH is unavailable or cannot be provided.
- For Docker database checks: Docker host access method, container name or id, database type, database credential source, source checklist workbook if report output is requested, output root, and whether original terminal screenshots are required.

If any required path or credential is missing and cannot be inferred, ask one concise question.

## Windows VMware Main Route

1. Preflight VMware:
   - Verify `vmrun.exe` works.
   - Verify VMware Tools guest operations work.
   - Verify the requested guest account can run commands and copy files.

2. Prepare guest GUI runtime:
   - Copy scripts into a per-run guest work directory.
   - Run `scripts/guest_preflight.py`.
   - If Python or PyAutoGUI dependencies are missing, run `scripts/guest_setup_pyautogui.ps1`. It must write a manifest with the absolute managed `python.exe` path; do not depend on PATH after installing Python.
   - For no-network guests, use `offline/install_guest_offline.ps1`; it installs the bundled Python 3.9.13 x64 runtime and PyAutoGUI wheelhouse.
   - Guest-side temporary files and managed environments belong under the current guest user's profile, for example `C:\Users\<username>\CodexVmAudit`, not `C:\Windows\Temp` or a root-level directory.
   - When the offline bundle is available, always install/use the bundled Python in that managed guest user work directory, even if the guest already has Python. This keeps the bundled cp39 wheelhouse compatible with the interpreter and avoids relying on customer PATH state. Online/non-bundled setup may still use an existing guest Python as the base interpreter for a per-run venv.
   - After successful evidence collection, clean guest runtime directories and the managed Python/venv environment. If the guest already had Python, clean only the per-run venv/dependencies and staged files, not the customer's base Python. Preserve the environment only for diagnostics, such as with `-KeepTmp` or `-KeepGuestPythonEnv`.
   - Launch GUI work from the interactive admin desktop through a highest-privilege scheduled task using the explicit `-InteractiveGuestUser` account. If the account was not provided, stop and ask for the Windows Server login account before invoking `vmrun` GUI evidence.

3. Analyze workbook:
   - Run `scripts/analyze_checklist.py`.
   - By default do not include screenshot filenames. Pass `--screenshots` only when screenshot mode is enabled.
   - Load `references/check_patterns.json` for known checks.
   - Identify columns such as `分类`, `测评项`, `预期结果`, `评估操作示例`, `检查情况`, `结果`, `整改建议`.
   - Mark only `检查情况` and `结果` as writable output columns. `整改建议` must be tracked as read-only context and must not be changed.
   - Produce an execution plan with row numbers, check ids, evidence filenames, and confidence.
   - The execution plan JSON is runtime data: write it only under `<label>安全检查证据/tmp/plan.json`, never next to the workbook, and delete it with `tmp` after the workbook and evidence outputs have both been checked.

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
   - Table/report output has three states for both Windows and Linux:
     1. No screenshot wording: do not collect screenshots; write detection results to the requested workbook or document only.
     2. Screenshot/evidence wording without an instruction to place images in the report: collect screenshots only as auxiliary evidence in a standalone folder named `<label>安全检查证据`; write workbook/document text results without screenshot filenames or images.
     3. Screenshot/evidence wording plus an instruction to place images in the report/document/table: first collect the same standalone evidence folder, then embed screenshots into the specified report location; if no exact location is specified, embed them in the detected `检查情况` column.
   - The final evidence directory root must contain final screenshots only.
   - If workbook output is requested, preserve host `evidence/tmp` until `runner_result.json` has been merged into the copied workbook, the workbook has been validated, and the evidence directory has been checked. Then clean `evidence/tmp` as the final cleanup step.
   - Do not place runtime JSON files such as `plan.json`, `runner_result.json`, or `image_validation.json` in the evidence root or guest work root; stage them under `evidence/tmp` and clean them only after both report output and evidence output have been checked, unless diagnostics were explicitly kept.
   - Only write Excel when requested.
   - When using `scripts/host_orchestrator.ps1` for a run that will produce a workbook, pass `-DeferHostTmpCleanup` instead of `-KeepTmp`: host `evidence/tmp` remains available for report generation, guest tmp/runtime is preserved for follow-up collection, and `deferred_guest_cleanup.json` is written under host `evidence/tmp` so the deferred guest cleanup scope is explicit.
   - For embedded-image output, prefer `scripts/workbook_embed_excel_com.ps1` on Windows hosts with Excel installed; insert screenshots directly into the requested location or the detected `检查情况` column and remove filename/link wording without changing workbook styling or layout unless explicitly requested.

## Linux/Unix Command And SSH Evidence Workflow

Use this workflow for Linux/Unix checks. Ask for SSH IP and account credentials first, then use SSH as the default transport. Treat screenshots as optional. Use `vmrun` direct command output only when the user says SSH is unavailable, not enabled, unreachable, or they cannot provide SSH IP/account credentials.

Do not use this workflow for Excel/report editing alone. Keep it limited to command execution and optional screenshot evidence collection; use the spreadsheet/report workflow after command output or screenshot evidence is collected.

1. Create or identify a JSON command list. Use `scripts/linux-baseline-commands.json` for the bundled Linux checklist and `scripts/sample-commands.json` as the minimal schema reference. If workbook output is needed, each command must map to a checklist row through an explicit `row`/`workbookRow` field or an id/name/command marker such as `row05`.
2. Default fast path: run `scripts/run-linux-commands.ps1` with SSH host/IP and account credentials. The output is per-command text files plus `manifest.json`, with no screenshots.
3. If the user says SSH is unavailable or cannot be provided, rerun `scripts/run-linux-commands.ps1` with `-AllowVmrunFallback` plus VMX/guest credentials, or run `scripts/run-guest-commands.ps1` directly.
4. Screenshot path only when requested: run `scripts/capture-ssh-evidence.ps1` from Windows PowerShell. This remains SSH-first because it preserves the upstream Windows Terminal screenshot workflow. When a command can produce long output, keep `command` as the complete raw collection command and add `screenshotCommand` or `focusCommand` that prints only the row's key fields for the screenshot.
5. If screenshot mode is on, verify the output folder contains per-command screenshots such as `01_rowXX.png`/`01_cmdXX.png` and `manifest.json`.
6. If Excel output is requested, run `scripts/ssh_workbook_plan.py` to convert the command list and optional `manifest.json` into the same plan format used by the Windows workflow. Pass `--manifest-json` from the raw SSH command collection so summaries are based on complete text output. If screenshots were collected separately, also pass `--screenshot-manifest-json` and `--screenshots`; the text workbook writer still keeps images out of `检查情况` unless an embedded-image report was explicitly requested.
   - For Linux workbook cells, `检查情况` should be a concise key-result summary, for example `PASS_MAX_DAYS=99999、PASS_MIN_LEN=5；未发现 TMOUT 设置。`
   - Do not include `已通过 SSH...`, `检查方式`, `证据文件`, or evidence filenames in `检查情况` unless the user explicitly asks for that wording.
   - `结果` must be one of the workbook's fixed dropdown choices. Prefer `符合` when the command output meets the expected result, `不符合` when required configuration is missing or weaker than expected, `不适用` only when the row truly does not apply, and `未检查` only for skipped/unoperated rows.
   - Do not write to `整改建议` for Linux/Unix reports.
7. Before delivering screenshot evidence or writing a screenshot report, run `scripts/finalize_evidence_names.py` to rename final screenshots from operational names to concise names such as `row05_身份鉴别.png` or `row06_口令策略.png`. Prefer `evidenceLabel`/`shortName` from the command JSON; do not use the full checklist sentence as the filename. Call `scripts/workbook_output.py --mode text` for text-only reports, or `scripts/workbook_embed_excel_com.ps1` / `scripts/workbook_output.py --mode embed-images` only when the user requested screenshots inside the report.
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
- `command`: complete read-only shell command used for raw text collection and workbook inference.
- `screenshotCommand` or `focusCommand`: optional shorter read-only command used only by `capture-ssh-evidence.ps1` to show the row's key evidence in one readable terminal screenshot. Use it for rows such as password policy where `cat /etc/login.defs` is complete but too long for one screenshot.
- `evidenceLabel` or `shortName`: optional concise final screenshot label, for example `口令策略`; keep it short and do not copy the full checklist sentence.
- `waitSeconds`: optional delay before screenshot; increase for long-running commands.
- `screenshotWaitSeconds`: optional screenshot-only delay override.
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
  -OutputRoot D:\WorkSpace\Evidence `
  -TaskLabel linux-baseline
```

The script creates:

```text
<OutputRoot>\<label>安全检查证据\
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
  --manifest-json D:\WorkSpace\Evidence\linux-baseline_raw\manifest.json `
  --screenshot-manifest-json D:\WorkSpace\Evidence\linux-baseline安全检查证据\manifest.json `
  --screenshots `
  --task-label linux-baseline `
  --out D:\WorkSpace\Evidence\linux-baseline安全检查证据\tmp\plan.json

python .\scripts\finalize_evidence_names.py `
  --plan D:\WorkSpace\Evidence\linux-baseline安全检查证据\tmp\plan.json `
  --evidence-dir D:\WorkSpace\Evidence\linux-baseline安全检查证据 `
  --out D:\WorkSpace\Evidence\linux-baseline安全检查证据\tmp\report_plan.json

python .\scripts\workbook_output.py `
  --source-workbook D:\WorkSpace\checklist.xlsx `
  --plan D:\WorkSpace\Evidence\linux-baseline安全检查证据\tmp\report_plan.json `
  --evidence-dir D:\WorkSpace\Evidence\linux-baseline安全检查证据 `
  --output-workbook D:\WorkSpace\linux-baseline安全检查报告_20260430_153000.xlsx `
  --mode text `
  --cleanup-tmp
```

Use `scripts/workbook_embed_excel_com.ps1` instead of `workbook_output.py --mode text` when the user requests embedded screenshot images and Excel is installed. For Windows GUI evidence collected through `scripts/host_orchestrator.ps1`, pass `-DeferHostTmpCleanup` during collection and then pass `--runner-result <evidence>\tmp\runner_result.json --cleanup-tmp` to `workbook_output.py` or `-RunnerResultJson <evidence>\tmp\runner_result.json -CleanupTmp` to `workbook_embed_excel_com.ps1` so runtime judgment data is merged before tmp cleanup. Do not use `--include-evidence-filenames` unless the user explicitly asks for screenshot filenames as text.

## Docker MySQL Baseline Workflow

Use this workflow when the checked MySQL/MariaDB service runs inside a Docker container, for example a Pikachu target container such as `pikachu-local`. The same evidence and Excel rules apply: original screenshots must be real terminal/window captures, raw command output stays in text files, and workbook cells contain conclusions only.

1. Identify the target:
   - Confirm or infer the container name/id, database type, Docker host access, source checklist workbook, and output label.
   - Prefer read-only inspection commands. Do not change container configuration or database settings unless the user explicitly requests remediation.

2. Create the evidence directory:
   - `raw/`: original command output text, one file per check module.
   - `original_screenshots/`: real terminal, VM console, remote desktop, or system screenshots when original screenshots are requested.
   - `screenshots/`: optional rendered command-output images only when generated; never label these as original screenshots.
   - `manifest.json`: command id, module name, command text, collection time, raw output path, screenshot path, screenshot type, and conclusion mapping.
   - `检查说明.txt`: explain evidence source, screenshot type distinction, workbook path, and any skipped or unsupported checks.

3. Collect Docker status and metadata:

```powershell
docker ps --filter name=pikachu-local
docker inspect pikachu-local
```

Check whether the container is running, image name, port mappings, `Privileged`, network mode, mounts, environment variables, and whether MySQL port `3306` is exposed to the host.

4. Collect MySQL process and filesystem evidence inside the container:

```sh
whoami
id
ps aux | grep -E "mysqld|USER" | grep -v grep
ls -la /etc/mysql
ls -la /var/lib/mysql
ls -la /var/log/mysql
```

Check whether `mysqld` runs as root, whether data/log directories have overly broad permissions, and whether ownership belongs to a dedicated database user.

5. Check MySQL accounts and empty passwords:

```sql
select version(), current_user();
show databases;
select
  user,
  host,
  plugin,
  case when authentication_string='' then 'EMPTY' else 'SET' end as password_state,
  account_locked
from mysql.user
order by user,host;
```

Check empty-password users, empty-password `root`, remote `root` entries such as `%`, default accounts, locked status, and whether the application uses `root` for business access.

6. Check MySQL security variables:

```sql
show variables where Variable_name in (
  'bind_address',
  'port',
  'general_log',
  'slow_query_log',
  'log_bin',
  'max_user_connections',
  'interactive_timeout',
  'wait_timeout',
  'log_error',
  'general_log_file',
  'slow_query_log_file'
);

show global variables like '%connect%';
```

Check whether `bind_address` is `0.0.0.0`, whether default port `3306` is used, whether `general_log`, `slow_query_log`, and `log_bin` are enabled as expected, whether per-user connections are limited, whether timeouts are excessive, and whether error logging is configured.

7. Check root grants:

```sql
show grants for `root`@`localhost`;
```

Check high-risk permissions such as `FILE`, `GRANT OPTION`, `SHUTDOWN`, `PROCESS`, broad administrative privileges, and least-privilege alignment.

8. Check application database connection configuration with key-only output:

```sh
echo "## /app/inc/config.inc.php"
grep -nE "DBHOST|DBUSER|DBPW|DBNAME|DBPORT" /app/inc/config.inc.php 2>/dev/null

echo

echo "## /app/inc/mysql.inc.php"
grep -nE "function connect|mysqli_connect" /app/inc/mysql.inc.php 2>/dev/null
```

Use key-only output for screenshots and summaries. Check database host, port, database name, hard-coded credentials, empty passwords, and application use of `root`.

9. Check running services and minimal deployment:

```sh
service --status-all 2>&1 | head -80
ps aux | head -40
```

Check unnecessary services, combined Web/MySQL/supervisord processes, and minimal-installation alignment.

10. Check history files:

```sh
find / -name .mysql_history -o -name mysql.history -o -name .bash_history 2>/dev/null | while read f; do
  echo "--- $f"
  ls -l "$f"
  wc -c "$f" 2>/dev/null
  head -20 "$f" 2>/dev/null
done
```

Check whether `.mysql_history`, `mysql.history`, or `.bash_history` contains sensitive commands, plaintext passwords, or leakage risk.

11. Recommended module ids for command lists and manifests:
   - `docker_status`: `docker ps`, container state, and port mappings.
   - `docker_inspect`: image, user, network, privileged mode, ports, mounts, and environment.
   - `mysql_process_files`: `whoami`, `id`, process list, MySQL directory permissions.
   - `mysql_users_passwords`: `mysql.user`, default accounts, empty passwords, locked status.
   - `mysql_variables`: bind address, port, logs, connection limits, and timeouts.
   - `mysql_root_grants`: root high-risk grants.
   - `app_db_config`: key DB connection settings and connect call sites.
   - `services`: service/process minimization.
   - `history_files`: shell/MySQL history leakage.

12. Workbook output:
   - Start from a copied workbook based on the user-provided template.
   - Preserve original colors, fills, borders, fonts, row/column sizes, merged cells, sheet structure, existing shapes, and hyperlinks.
   - Fill only `检查情况` and `结果`; do not modify `整改建议`.
   - Keep screenshot paths and raw-output paths out of table cells; record them in `manifest.json` or `检查说明.txt`.

13. Verification:
   - Confirm expected raw text files exist under `raw/`.
   - Confirm every requested original screenshot is a real terminal/window capture under `original_screenshots/`, not a rendered command-output image.
   - Confirm long-output screenshots show only readable, conclusion-supporting fields.
   - Confirm `manifest.json` maps each module to commands, raw output, screenshots, screenshot type, and workbook conclusion.
   - Confirm workbook result cells contain no `证据:`, screenshot filenames, raw paths, image paths, or hyperlink remnants unless explicitly requested.
   - Confirm workbook formatting matches the source template except for intended result-cell values.
   - Confirm `整改建议` cell values and formatting are unchanged from the source workbook.
   - After both the workbook and evidence directory pass validation, delete `tmp` from the evidence directory unless the user explicitly asked to keep diagnostics.

## Output Naming

Use sibling outputs next to the source workbook unless the user specifies otherwise:

- Workbook/report: `<label>安全检查报告_<yyyyMMdd_HHmmss>.xlsx` where `<label>` is the task label, system name, VM name, host name, or another user-specified output label.
- Screenshot evidence directory: `<label>安全检查证据` by default. Do not append a timestamp unless the user explicitly requests timestamped evidence folders or a collision must be avoided.
- Runtime plan during screenshot collection: `<label>安全检查证据/tmp/plan.json` only; it is not a final output.
- Linux/Unix command workbook plan: `<command_output_dir>/tmp/plan.json`.

Use timestamps for workbook/report filenames by default. Do not add nested final-evidence folders unless the user asks.
Do not leave intermediate files in the final evidence directory. Process files such as scripts, runtime JSON, execution plans, logs, helper scripts, diagnostic screenshots, contact sheets, and validation JSON may live under `<label>安全检查证据/tmp` during collection and report generation. Delete that `tmp` directory only at the very end, after both the workbook/report and evidence directory have been checked successfully, unless the user explicitly asks to keep diagnostics.

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
- `scripts/guest_setup_pyautogui.ps1`: detect or install guest Python, create a managed venv, install PyAutoGUI dependencies, write the Python environment manifest, and clean the managed environment when called with `-Cleanup`.
- `scripts/guest_gui_runner.py`: guest-side PyAutoGUI runner and action registry, including the first production actions from the successful Server 2012 run.
- `scripts/host_orchestrator.ps1`: sample vmrun orchestration wrapper.
- `scripts/workbook_output.py`: portable sample workbook writer for text or basic embedded-image output.
- `scripts/workbook_embed_excel_com.ps1`: high-fidelity Excel COM embedded-image writer for Windows hosts with Microsoft Excel.
- `scripts/audit_mode.py`: detects whether the user requested screenshots; default is no screenshots.
- `scripts/run-linux-commands.ps1`: Linux/Unix no-screenshot dispatcher; requires SSH first and uses vmrun only when `-AllowVmrunFallback` is passed after the user says SSH is unavailable or cannot be provided.
- `scripts/run-ssh-commands.ps1`: fast no-screenshot SSH command runner that writes text output and `manifest.json`; password SSH delegates to `scripts/run-ssh-commands.py`, while key/agent SSH uses OpenSSH.
- `scripts/run-ssh-commands.py`: Paramiko-backed password SSH runner for no-screenshot Linux/Unix command output.
- `scripts/run-guest-commands.ps1`: fast no-screenshot VMware guest command runner that writes text output and `manifest.json`; use as Linux fallback only after the user says SSH is unavailable or cannot be provided.
- `scripts/capture-ssh-evidence.ps1`: Linux/Unix SSH command screenshot evidence collector; preserves full command metadata and uses `screenshotCommand`/`focusCommand` when present so long-output rows show the key evidence area.
- `scripts/linux-baseline-commands.json`: bundled Linux checklist command list; uses complete raw commands for workbook inference and focused screenshot commands for readable evidence.
- `scripts/sample-commands.json`: minimal Linux security command list and schema reference for SSH evidence collection.
- `scripts/ssh_workbook_plan.py`: convert Linux/Unix SSH command lists plus optional raw/screenshot manifests into workbook output plans that reuse the Windows Excel writers.
- `scripts/finalize_evidence_names.py`: final-only screenshot renamer; use after capture and before workbook output so collection keeps safe operational names while deliverables use concise row labels.

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
- Evidence filenames, when screenshots are requested, follow the final row naming contract: `rowNN_<简写>.png`, with `_补充N` for additional screenshots on the same row, and live under `<label>安全检查证据`.
- Workbook output, if requested, is a copied workbook and not the original.
- Workbook output, if requested, has merged `runner_result.json` after evidence collection and validated the written `检查情况`/`结果` values. `整改建议` must remain unchanged. Delete `evidence/tmp` only after both workbook/report validation and evidence-directory validation pass.
- Workbook output preserves the source workbook's structure, colors, fills, row/column sizes, borders, fonts, alignment, existing shapes, and hyperlinks unless the user explicitly requested formatting/layout changes.
- Text workbook/document output has no screenshot filenames unless the user explicitly asked for filenames as text. Embedded-image output has no `截图:`, `证据:`, `rowNN_`, or hyperlink remnants in the check-result cells.
- Administrator-interview rows have `检查情况` set to `涉及访谈管理员，未进行操作`, `结果` set to `未检查`, and no GUI/command evidence operation.
- The final evidence directory contains only accepted final screenshots when screenshot mode is enabled; no `tmp` directory, runner logs, stdout/stderr text, runtime JSON files, validation JSON, contact sheets, or helper scripts remain unless diagnostics were explicitly kept.
- Report whether any rows were adaptive, unsupported, or manually confirmed.
- For Linux/Unix checks, the output directory contains either per-command text files plus `manifest.json` or per-command screenshots plus `manifest.json`; failures are reported with the command id/name, and workbook output if requested is a copied workbook with mapped command evidence written only to `检查情况`/`结果`.
