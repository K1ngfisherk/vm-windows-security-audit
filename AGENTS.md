# Agent Guide For `vm-windows-security-audit`

This file is a fast orientation layer for AI agents editing or using this skill.
`SKILL.md` remains the source of truth. Keep this file short and update it when
workflow contracts or script entry points change.

## Purpose

This skill audits security baseline checklists for:

- Windows VMware guests, using `vmrun`, VMware Tools, optional guest PyAutoGUI,
  native Windows GUI screenshots, and optional Excel report output.
- Linux/Unix hosts, using SSH-first command collection, optional SSH screenshot
  evidence, optional vmrun fallback, and optional Excel report output.
- Docker database services, especially MySQL/MariaDB containers, using
  `docker inspect`, `docker exec`, raw command-output capture, optional real
  terminal screenshots, and optional Excel report output.

Default mode is fast and no-screenshot. Screenshot collection is opt-in.

## Hard Rules

- Do not modify the original checklist workbook. Reports must be copied workbooks.
- In copied report workbooks, only `检查情况` and `结果` are writable by default.
  `整改建议` is strictly read-only: do not change its value, style, width,
  visibility, comments, hyperlinks, or formatting.
- Do not collect screenshots unless the user asks for screenshots, evidence,
  proof, capture, `截图`, `证据`, `取证`, `screenshot`, or similar wording.
- Explicit no-screenshot wording wins over screenshot/evidence wording.
- Screenshot collection and report embedding are separate. Screenshot keywords
  alone mean a standalone auxiliary folder named
  `<label>安全检查证据`. Embed images into a workbook/document
  only when the user asks to put/insert/embed screenshots there; if no exact
  placement is specified, use the `检查情况` column.
- Windows screenshot evidence must be a native Windows GUI page/dialog such as
  `secpol.msc`, `lusrmgr.msc`, `eventvwr.msc`, `services.msc`, `regedit`,
  Control Panel, or an MMC properties dialog.
- Never use PowerShell, cmd, Windows Terminal, command output, registry dumps,
  rendered text, or synthetic PNGs as a substitute for required Windows GUI
  evidence.
- If native Windows GUI evidence is blocked or cannot be verified, stop that row
  and ask the user before using any alternate evidence type.
- Command-window evidence is disabled by default and requires explicit
  authorization for the exact row.
- Administrator-interview rows are not operated. Write:
  - `检查情况`: `涉及访谈管理员，未进行操作`
  - `结果`: `未检查`
- Linux/Unix must ask for SSH IP, username, and password/key first. Use vmrun
  only after the user says SSH is unavailable or credentials cannot be provided.
- Keep Linux/Unix commands read-only unless the user explicitly authorizes
  changes.
- For Docker/database checks, collect metadata with `docker inspect` and service
  evidence with `docker exec`; keep commands read-only unless remediation is
  explicitly requested.
- Original screenshots must be real terminal, VM console, remote desktop, or
  system screenshots. Rendered command-output images belong under `screenshots`
  and must never be described as original screenshots.
- Long Docker/database outputs should be filtered or split so screenshots show
  readable key evidence, not unreadable full dumps.
- Runtime JSON, logs, diagnostics, helper scripts, and temp files belong under
  `evidence/tmp`, not the evidence root.
- If workbook output needs `runner_result.json`, defer host tmp cleanup until the
  workbook has merged and validated the collected results and the evidence
  directory has also been checked.

## Route Selection

Use `scripts/audit_mode.py` or the same keyword logic to decide screenshot mode.

Windows checklist route:

1. Run `scripts/analyze_checklist.py` to build `evidence/tmp/plan.json`.
2. If screenshots are off, do not launch the GUI runner.
3. If screenshots are on, run `scripts/host_orchestrator.ps1` and pass
   `-Screenshots`.
4. For workbook output, pass `-DeferHostTmpCleanup` during collection, then pass
   `runner_result.json` to the workbook writer and clean tmp last.

Linux/Unix route:

1. Ask for SSH IP, user, and password/key.
2. Default no-screenshot path: `scripts/run-linux-commands.ps1`.
3. Password SSH uses `scripts/run-ssh-commands.py` through Paramiko.
4. Key/agent SSH uses `scripts/run-ssh-commands.ps1` and OpenSSH.
5. vmrun fallback: `scripts/run-guest-commands.ps1`, only after explicit user
   statement that SSH is unavailable or cannot be provided.
6. Screenshot path, only when requested: `scripts/capture-ssh-evidence.ps1`.
7. Excel plan: `scripts/ssh_workbook_plan.py`.
8. Final screenshot Chinese names: `scripts/finalize_evidence_names.py`.
9. Workbook output: `scripts/workbook_output.py` or
   `scripts/workbook_embed_excel_com.ps1`.

Docker MySQL route:

1. Confirm container name/id, Docker host access, database type, credential
   source, checklist workbook, output root, and whether original screenshots are
   required.
2. Create `raw/`, optional `original_screenshots/`, optional `screenshots/`,
   `manifest.json`, and `检查说明.txt`.
3. Collect modules: `docker_status`, `docker_inspect`, `mysql_process_files`,
   `mysql_users_passwords`, `mysql_variables`, `mysql_root_grants`,
   `app_db_config`, `services`, and `history_files`.
4. For app DB config evidence, capture only key fields such as `DBHOST`,
   `DBUSER`, `DBPW`, `DBNAME`, `DBPORT`, `function connect`, and
   `mysqli_connect`.
5. Workbook output must be copied from the source template and contain
   conclusions only, with evidence paths kept in `manifest.json` or
   `检查说明.txt`.

## Script Map

- `scripts/audit_mode.py`: screenshot/report-embedding keyword decision.
- `scripts/analyze_checklist.py`: Windows workbook plan builder.
- `scripts/host_orchestrator.ps1`: Windows vmrun orchestration.
- `scripts/guest_gui_runner.py`: guest-side native GUI automation and screenshot
  validation.
- `scripts/guest_setup_pyautogui.ps1`: guest-side managed Python/venv setup,
  manifest writing, and cleanup for Windows GUI evidence.
- `scripts/workbook_output.py`: portable workbook writer.
- `scripts/workbook_embed_excel_com.ps1`: high-fidelity Excel COM image embedder.
- `scripts/run-linux-commands.ps1`: Linux/Unix dispatcher, SSH-first.
- `scripts/run-ssh-commands.ps1`: no-screenshot OpenSSH/key runner.
- `scripts/run-ssh-commands.py`: no-screenshot Paramiko/password runner.
- `scripts/run-guest-commands.ps1`: vmrun command fallback.
- `scripts/capture-ssh-evidence.ps1`: SSH screenshot collector.
- `scripts/ssh_workbook_plan.py`: Linux command output to workbook plan.
- `scripts/finalize_evidence_names.py`: final screenshot Chinese renamer.
- `references/check_patterns.json`: Windows row pattern to GUI action mapping.
- `references/evidence_rules.md`: acceptance rules for screenshots and reports.

## Workbook Output Rules

- Preserve workbook structure, sheet layout, merged cells, row heights, column
  widths, colors, fills, fonts, borders, alignment, existing shapes, and
  hyperlinks unless the user explicitly requests formatting changes.
- Only write to `检查情况` and `结果`. `整改建议` may be read as checklist context,
  but must remain exactly as it was in the source workbook.
- When the user does not specify names, place images/evidence under
  `<label>安全检查证据` and name the copied report
  `<label>安全检查报告_<yyyyMMdd_HHmmss>.xlsx`.
- Process files such as scripts, JSON, plans, logs, helper files, and diagnostic
  screenshots may live under `<label>安全检查证据/tmp` while working. Delete that
  `tmp` directory only after both the report workbook and evidence directory
  have passed final checks, unless diagnostics were explicitly kept.
- `检查情况` should contain the key result and concise description.
- Text workbook/document output should not include screenshot filenames or
  images unless explicitly requested. If screenshots were requested but not
  report embedding, keep them only in the standalone evidence folder.
- Do not write `证据: xxx.png`, screenshot paths, raw-output paths, or image
  filenames into result cells unless the user explicitly asks for evidence paths
  as table text.
- For Linux reports, do not write process wording such as `已通过 SSH`,
  `检查方式`, `证据文件`, or evidence filenames unless the user explicitly asks.
- `结果` must be selected from the workbook's fixed options, commonly `符合`,
  `不符合`, `不适用`, and `未检查`.
- Before completion, verify `整改建议` was not modified.
- Do not treat workflow statuses such as `planned`, `not_run`,
  `needs_user_confirmation`, or `validation_failed` as report results.

## Verification Before Completion

Run the smallest useful checks for the changed surface:

- Python syntax: `python -m py_compile <changed .py files>`.
- PowerShell syntax: use `[System.Management.Automation.PSParser]::Tokenize`.
- JSON validity for edited `.json` files.
- `git diff --check`.
- If changing workbook writers, test at least one merge/write/validation path.
- If changing Linux row summaries, test representative command output parsing.
- If changing screenshot validation, confirm forbidden command-window screenshots
  fail unless exact-row authorization is present.
