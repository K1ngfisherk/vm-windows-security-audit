# Troubleshooting

## Black Screenshots

Cause: PyAutoGUI ran outside the interactive desktop, often through plain `runProgramInGuest`.

Fix:

- Confirm the requested user is logged in.
- Pass the logged-in Windows account explicitly as `-InteractiveGuestUser`; do not let the scheduled task run as Guest or another low-privilege account.
- Launch the runner through a highest-privilege scheduled task in that user's interactive session.
- Run a screenshot smoke test before starting checklist rows.

## MKS Input Problems

MKS input is not the main route.

- Do not use MKS for mouse or keyboard input.
- MKS screenshot capture can be used for observation only.
- Prefer guest-side PyAutoGUI screenshots for final evidence.

## MMC Or Policy Windows Do Not Respond

- Run from an elevated scheduled task.
- Avoid relying on MMC control APIs.
- Use PyAutoGUI keyboard navigation and screenshots.
- If a value is truncated in a list view, open the policy or account properties dialog.

## Missing PyAutoGUI

Run `scripts/guest_setup_pyautogui.ps1` in the guest. When a local offline
installer is provided, it always installs/uses that bundled Python in a managed
work directory, even if the guest already has Python. Without an offline
installer, it may reuse an existing Python as the base interpreter for a per-run
venv, or install Python through winget/an explicit download URL. The script
writes a manifest containing the absolute venv `python.exe` path; use that path
instead of relying on PATH.

For online guests:

```powershell
python -m pip install -r requirements-guest.txt
```

For offline guests, copy a wheelhouse and use:

```powershell
python -m pip install --no-index --find-links <wheelhouse> -r requirements-guest.txt
```

For guests without Python, copy a Python installer into the guest and run:

```powershell
powershell -ExecutionPolicy Bypass -File guest_setup_pyautogui.ps1 -PythonInstaller C:\path\python-installer.exe -Wheelhouse C:\path\wheelhouse -EnvRoot C:\Users\Administrator\CodexVmAudit\pyenv -InstallDir C:\Users\Administrator\CodexVmAudit\python39 -Manifest C:\Users\Administrator\CodexVmAudit\evidence\tmp\python_env.json
```

The guest runner also uses `pygetwindow` to find and maximize GUI windows. If
screenshots are blank or the runner reports that a window cannot be found, make
sure both `pyautogui` and `pygetwindow` import successfully in the same guest
Python used by the scheduled task.

After a successful run, call the same script with `-Cleanup` and the manifest
path, or let `scripts/host_orchestrator.ps1` do it automatically. Cleanup
removes only the managed venv and managed Python directory recorded in the
manifest; it does not uninstall a customer Python that existed before the run.

## OptionalFeatures.exe Error

If Windows says it cannot find `OptionalFeatures.exe`, do not use that route for installed components on Server Core or restricted Server builds. Use `control appwiz.cpl` or Server Manager where available.

## Workbook Fidelity

If openpyxl output changes formatting too much, use Excel COM on the host for final report writing. Always keep the original workbook untouched and write a copy.

For embedded-image reports, prefer `scripts/workbook_embed_excel_com.ps1` when
Microsoft Excel is installed on the host. Use `scripts/workbook_output.py` as a
portable fallback when Excel COM is unavailable.
