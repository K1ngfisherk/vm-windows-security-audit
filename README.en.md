# VM Windows Security Audit Skill

[中文](README.md)

`vm-windows-security-audit` is a portable agent skill scaffold for auditing Windows VMware virtual machines against security checklist spreadsheets.

It uses `vmrun`, VMware Tools, and guest-side PyAutoGUI to operate the Windows GUI inside the VM, capture original screenshot evidence, and optionally generate an Excel report with evidence images embedded in the checklist.

## Features

- Analyze `.xlsx` security checklists and map rows to GUI audit actions.
- Open Windows management tools such as `secpol.msc`, `gpedit.msc`, `lusrmgr.msc`, `services.msc`, `eventvwr.msc`, `fsmgmt.msc`, and `regedit`.
- Collect row-level screenshot evidence with stable names such as `row11_gpedit_rdp_client_connection_encryption_level.png`.
- Prefer graphical evidence over command output when the checklist requires GUI verification.
- Support adaptive GUI handling for rows that do not match a known pattern.
- Automatically prepare guest Python and PyAutoGUI dependencies with the bundled offline environment package when the VM does not already have them.
- Optionally write a copied workbook and embed evidence images into the `检查情况` column.

## Main Workflow

```text
vmrun + VMware Tools + guest-side PyAutoGUI + Excel
```

VMware MCP can be used as an optional helper, but it is not required.

`scripts/*.py` files are internal workflow components, not user-facing entry points. Run audits through the skill workflow or `scripts/host_orchestrator.ps1`; the workflow invokes Python scripts on its own.

## Layout

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

## Requirements

Host:

- Windows host with VMware Workstation or compatible `vmrun.exe`.
- VMware Tools available in the guest.
- Microsoft Excel if using high-fidelity embedded-image workbook output.

Guest:

- Windows VM with an interactive desktop session.
- A valid guest account with permission to perform the audit.
- Python and GUI automation dependencies can be detected and installed by the setup script.
- The repository includes an `offline/` bundle with the Python 3.9.13 x64 installer and a PyAutoGUI dependency wheelhouse.
- Online guests can reuse an existing Python install or winget; offline guests can run `offline/install_guest_offline.ps1`.

Guest dependencies prepared by the setup script:

```text
pyautogui
pillow
pyperclip
pygetwindow
```

Prepare the guest environment:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/guest_setup_pyautogui.ps1
```

For offline guests:

```powershell
powershell -ExecutionPolicy Bypass -File offline/install_guest_offline.ps1
```

## Output Convention

```text
<source_workbook_stem>_<task_label>.xlsx
Windows完整检查_<task_label>_证据
Windows完整检查_<task_label>_执行计划.json
```

Screenshot names:

```text
rowNN_<tool>_<check_key>.png
```

The final evidence directory root should contain accepted final screenshots only. Logs, stdout/stderr captures, runtime JSON files such as `plan.json`, `runner_result.json`, and `image_validation.json`, contact sheets, temporary scripts, and diagnostic error screenshots must go under `Windows完整检查_<task_label>_证据/tmp`, and that `tmp` directory is deleted by default when the task completes.

Installed-update evidence must land on the actual `Installed Updates` page and show the Windows update/KB list area. Do not accept the plain installed-programs list as patch-location evidence.

## Notes

- Workbook output, hyperlinks, and zip packages are optional and should only be generated when requested.
- GUI navigation is currently tuned around Windows Server 2012 style MMC windows and may need adjustment for other Windows versions or display scaling.
