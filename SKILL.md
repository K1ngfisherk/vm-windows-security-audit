---
name: vm-windows-security-audit
description: Use when a user wants to audit a Windows VMware VM against a security checklist spreadsheet, collect graphical screenshot evidence, and optionally write an Excel report. The main route is vmrun + VMware Tools + guest-side PyAutoGUI + Excel workbook output; VMware MCP is optional assistance, not a required dependency.
---

# VM Windows Security Audit

Use this skill to run Windows security checklist work against a VMware guest with real GUI evidence.

## Operating Contract

- Use the requested guest account. Do not silently switch to `guest` or another account.
- Treat graphical evidence as the source of truth when the checklist asks for GUI inspection.
- Do not use command output to replace a required GUI screenshot.
- Process checklist rows sequentially: finish and verify the current row screenshot before moving to the next row.
- Leave the original workbook untouched. If report output is requested, write a copied workbook only.
- Do not create zip packages, workbook hyperlinks, or workbook output unless the user asks for them.

## Inputs

Collect or infer:

- VMX path.
- Guest username and password.
- Source checklist workbook.
- Task label from the target VM, user wording, or explicit label.
- Output directory.
- Output mode: screenshots only, workbook text output, workbook embedded-image output, optional zip.

If any required path or credential is missing and cannot be inferred, ask one concise question.

## Main Route

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
   - Load `references/check_patterns.json` for known checks.
   - Identify columns such as `分类`, `测评项`, `预期结果`, `评估操作示例`, `检查情况`, `结果`, `整改建议`.
   - Produce an execution plan with row numbers, check ids, evidence filenames, and confidence.

4. Execute checks:
   - For known checks, call the mapped GUI action.
   - For inferred checks, open the tool named or implied by the row.
   - For unmatched checks, use the adaptive GUI rules in `references/gui_adaptive_rules.md`.
   - Screenshot each row as `rowNN_<tool>_<check_key>.png`.
   - Put only final accepted screenshots in the evidence directory root.
   - Put runner logs, stdout/stderr captures, visible-command helper scripts, runtime JSON files, validation JSON, contact sheets, and diagnostic/error screenshots in `evidence/tmp`.

5. Verify evidence:
   - Check that the screenshot is not blank.
   - Check that the window title, path/node, keyword, and target value are visible.
   - If text is truncated, expand columns, open properties, zoom, or resize before accepting.
   - For installed-updates checks, accept only the real `已安装更新` / `Installed Updates` page. The screenshot must show that page title or breadcrumb and the Windows update list/KB entry area, not just the `卸载或更改程序` / installed-programs list.

6. Output:
   - Screenshots are always saved when checks run.
   - The final evidence directory root must contain final screenshots only.
   - Delete `evidence/tmp` at task completion unless the user explicitly asks to keep diagnostics for troubleshooting.
   - Do not place runtime JSON files such as `plan.json`, `runner_result.json`, or `image_validation.json` in the evidence root or guest work root; stage them under `evidence/tmp` and clean them with the rest of tmp.
   - Only write Excel when requested.
   - For embedded-image output, prefer `scripts/workbook_embed_excel_com.ps1` on Windows hosts with Excel installed; insert screenshots directly into the detected `检查情况` column and remove filename/link wording.

## Output Naming

Use sibling outputs next to the source workbook unless the user specifies otherwise:

- Workbook: `<source_workbook_stem>_<task_label>.xlsx`
- Evidence directory: `Windows完整检查_<task_label>_证据`
- Plan: `Windows完整检查_<task_label>_执行计划.json`

Do not add timestamps or nested final-evidence folders to final output names unless the user asks.
Do not leave intermediate files in the final evidence directory. Use `Windows完整检查_<task_label>_证据/tmp` only as a temporary staging area for logs, helper scripts, diagnostic screenshots, and runtime JSON files, then remove it when the run is complete.

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

## Tool Preferences

- Primary: `vmrun + VMware Tools + guest-side PyAutoGUI + Excel`.
- Optional: VMware MCP for VM status and management.
- Observation-only fallback: MKS screenshots.
- Avoid as main route: MKS input, host mouse coordinates, pure command-only evidence for GUI rows, AutoIt/MMC control automation.

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

Scripts are scaffolds intended to be adjusted for the target environment. Preserve the contracts above when patching them.

## Current Action Coverage

The bundled runner currently has actions for:

- Local Security Policy password policy, account lockout policy, audit policy, and security options.
- Group Policy RDP encryption level, idle session limit, and connection limit pages.
- Local users/groups user list, default account checks, Guest properties, and Administrators group membership.
- Shared folders, LSA registry values, Event Viewer system log/properties, installed updates, and Remote Registry service evidence.
- Visible command-window evidence for interview/account context rows such as identity and administrator authentication method.

Rows that do not match these actions still use the `inferred` or `adaptive` lane instead of being skipped.

## Completion Criteria

Before final response:

- All supported checklist rows have accepted screenshot evidence or a clear blocker.
- Evidence filenames follow the row naming contract.
- Workbook output, if requested, is a copied workbook and not the original.
- Embedded-image output has no `截图:`, `证据:`, `rowNN_`, or hyperlink remnants in the check-result cells.
- The final evidence directory contains only accepted final screenshots; no `tmp` directory, runner logs, stdout/stderr text, runtime JSON files, validation JSON, contact sheets, or helper scripts remain unless diagnostics were explicitly kept.
- Report whether any rows were adaptive, unsupported, or manually confirmed.
