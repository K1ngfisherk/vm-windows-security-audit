# Evidence Rules

## Screenshot Mode

- Screenshot collection is optional and off by default.
- Enable screenshots only when the user asks for screenshots, evidence, capture proof, or similar wording.
- When screenshot mode is off, do not create placeholder screenshot filenames; use command/text output or workbook-only results.

## Directory And Names

- Final evidence directory: `Windows完整检查_<task_label>_证据`.
- Put final screenshots directly in that directory when screenshot mode is enabled.
- Put all intermediate files under `Windows完整检查_<task_label>_证据/tmp`.
- Put runtime JSON files such as `plan.json`, `runner_result.json`, and `image_validation.json` under `tmp`; never leave them in the evidence root.
- If workbook output is requested, keep host `tmp` until `runner_result.json` has been merged into the copied workbook and the written `检查情况`/`结果` values have been validated; delete `tmp` only after that report check passes unless the user explicitly asks to keep diagnostics.
- Do not use timestamped final folders or nested `最终截图证据` folders unless requested.
- The final evidence directory root should contain accepted final screenshots only when screenshot mode is enabled, not logs, stdout/stderr captures, runtime JSON files, JSON diagnostics, helper scripts, contact sheets, or preview files.
- Screenshot filenames:
  - Lowercase ASCII.
  - Prefix with workbook row number: `rowNN_`.
  - Include a stable tool/source and check key.
  - End in `.png`.

Examples:

```text
row06_secpol_password_policy.png
row11_gpedit_rdp_client_connection_encryption_level.png
row16_lusrmgr_administrators_members.png
row16_services_mysql_logon_account.png
```

## Screenshot Acceptance

Accept a screenshot only when it proves the row:

- The target GUI window is visible.
- For Windows checklist GUI evidence, the foreground window is the native Windows management page or dialog for that row, such as Local Security Policy, Local Users and Groups, Event Viewer, Services, Registry Editor, Control Panel, or an MMC properties dialog.
- The relevant tree path, tab, policy name, account name, service name, or registry key is visible.
- The target value/status is visible.
- Important keywords are not truncated.
- The image is not black, blank, covered by a modal error, or on the wrong account/session.
- The image is not PowerShell, cmd/Command Prompt, Windows Terminal, rendered text, command output, or a synthetic PNG standing in for the native GUI page.

If text is truncated:

- Widen the column.
- Open the properties dialog.
- Resize or maximize the window.
- Scroll to keep the target row and value visible.
- Retake the screenshot.

## GUI Versus Command Evidence

- If the checklist row requests or implies GUI inspection, the final evidence must be GUI.
- Commands may help discover a service name, registry path, policy name, or installed component.
- Command output never replaces native Windows GUI evidence unless the user explicitly authorizes command-window evidence for that exact row after being told the native GUI blocker.
- If the native GUI route hangs, cannot be verified, or cannot produce proof, stop and ask the user before substituting another evidence type.

## Workbook Output

Do not write a workbook unless requested.

When writing a workbook:

- Copy the source workbook first.
- Preserve original formatting as much as possible.
- Detect the `检查情况` and `结果` columns instead of assuming fixed letters.
- Do not add hyperlinks by default.
- Do not write absolute evidence paths into cells.

When the user requests embedded evidence:

- Insert screenshots directly into the detected `检查情况` column.
- Keep a concise result sentence above the images.
- Remove filename/link wording such as `截图:`, `证据:`, `rowNN_...png`.
- Place multiple images side by side when practical.
- Adjust only the affected rows/columns needed for readability.
- Prefer Excel COM placement on Windows hosts with Excel installed, because it
  preserves the workbook's existing layout more reliably than generic `.xlsx`
  libraries for floating images.

## Special Case Rules

- RDP encryption level checks: use the graphical policy page when available. Do not use `MinEncryptionLevel`, `SecurityLayer`, or similar runtime values as replacement evidence.
- Default account checks: if checking whether Guest is disabled, open the Guest properties dialog.
- LSA `restrictanonymous` checks: expand the registry value list columns before capture so `restrictanonymous`, `restrictanonymoussam`, type, and displayed data are readable in the screenshot.
- Default account list checks: expand the Local Users and Groups `Users` list name column before capture so `Administrator`, `DefaultAccount`, `Guest`, and `WDAGUtilityAccount` are readable; the Guest disabled screenshot must show the Guest properties dialog and disabled checkbox/status.
- Security Options checks: for "do not display last username" and "clear virtual memory pagefile", use the Local Security Policy Security Options page when available. Do not add registry evidence if the GUI proves the item.
- Installed updates checks: use the `已安装更新` / `Installed Updates` page under Programs and Features. The final screenshot must show that page title or breadcrumb plus the Windows update list/KB entry area. Do not accept the plain `卸载或更改程序` / installed-programs list as evidence for patch location.
