# Security Baseline Assessment Project

[中文](README.md)

This project supports security baseline assessments for common targets such as Windows, Linux/Unix, and database services. It works around security checklists, organizes assessment results, and produces reports and evidence materials.

## Overview

The project can help review account policies, access control, audit settings, services and ports, remote management, database configuration, and related baseline items.

It is useful for security baseline assessments, coursework, review and acceptance work, operations self-checks, and before/after hardening records.

## Supported Targets

- Windows Server virtual machines
- Linux/Unix hosts
- Database services running in Docker
- Common MySQL/MariaDB scenarios

## Outputs

Typical outputs include:

- Excel security assessment reports
- Screenshot evidence folders
- Command output or assessment records
- Result summaries for each target

Example output:

```text
centos安全检查报告.xlsx
centos安全检查证据\
```

Screenshot files use short names for easier review and archiving, for example:

```text
row05_身份鉴别.png
row06_口令策略.png
row22_服务端口.png
row24_超时锁定.png
```

## Highlights

- Windows GUI assessment support.
- Linux/Unix SSH assessment support.
- Basic database service security checks.
- Excel report generation.
- Screenshot evidence collection.
- Key-evidence summaries for long command output.

## Typical Use Cases

- Checking server configurations against a baseline checklist.
- Producing consistent reports for multiple virtual machines.
- Preparing materials for security courses, labs, or assignments.
- Comparing system state before and after hardening.
- Preparing screenshots for review, acceptance, or archiving.

## Project Layout

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
