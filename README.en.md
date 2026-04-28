# VM Windows Security Audit Skill

[中文](README.md)

`vm-windows-security-audit` audits Windows VMware virtual machines against Windows security checklist spreadsheets. It reads an `.xlsx` checklist, opens the corresponding Windows management interfaces inside the VM, captures row-level screenshot evidence, and can optionally generate an Excel report with embedded screenshots.

## Use Cases

- Windows Server security baseline checks.
- Spreadsheet-driven audit tasks that require screenshot evidence.
- Windows system checks inside VMware Workstation virtual machines.
- Reviews where original GUI evidence is preferred over command-only output.

## Features

- Analyze `.xlsx` checklists and identify check items, expected results, and operation guidance.
- Support common Windows management surfaces, including Local Security Policy, Group Policy, Local Users and Groups, Services, Event Viewer, Shared Folders, Registry Editor, and Control Panel.
- Save row-level evidence with stable names such as `row11_gpedit_rdp_client_connection_encryption_level.png`.
- Perform basic screenshot usability and page-content checks to reduce mismatched evidence.
- Support offline preparation of Python and GUI automation dependencies inside the VM.
- Optionally generate an Excel report with screenshots embedded into the result column.

## Before You Start

Host requirements:

- Windows host.
- VMware Workstation or compatible `vmrun.exe`.
- Access to the target VM `.vmx` file.
- Access to the source checklist `.xlsx`.
- Microsoft Excel if an embedded-image Excel report is required.

Guest requirements:

- Windows guest system.
- VMware Tools running.
- Logged-in desktop session.
- Guest username and password for the audit.

## Workflow

1. Read the checklist and build an audit plan.
2. Check the VM environment and desktop state.
3. Prepare the runtime inside the guest.
4. Open the corresponding Windows management interface for each row.
5. Capture and validate screenshot evidence.
6. Save final screenshots.
7. Generate an Excel report if requested.

## Outputs

By default, outputs are written next to the source workbook and include the task label in their names.

```text
Windows完整检查_<task_label>_证据\
<source_workbook_stem>_<task_label>.xlsx
```

Screenshot naming format:

```text
rowNN_<tool>_<check_key>.png
```

The final evidence directory mainly contains deliverable screenshot files. Report workbooks are created only when requested.

## Current Coverage

Built-in checks currently cover:

- Local Security Policy: password policy, account lockout policy, audit policy, and security options.
- Group Policy: Remote Desktop encryption level, idle session limit, and connection limit.
- Local Users and Groups: user lists, default accounts, Guest properties, and Administrators membership.
- Services: Remote Registry and other service settings.
- Event Viewer: System log and properties pages.
- Shared Folders, LSA registry values, and Installed Updates.
- Helper checks that display identity or administrator authentication details in a visible window.

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
