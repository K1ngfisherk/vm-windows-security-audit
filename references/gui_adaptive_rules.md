# GUI Adaptive Rules

Use these rules when a checklist row does not match a known action.

## Context To Read

Use the row and nearby rows:

- Category.
- Check item.
- Expected result.
- Operation example.
- Existing remediation suggestion.
- Neighboring category rows that clarify the tool family.

Do not rely on the row number alone.

## Tool Inference

Prefer the tool explicitly named in the operation example. If none is named, infer:

- Password, lockout, audit, local policy, security options: `secpol.msc`.
- Group Policy, administrative templates, Remote Desktop Services, terminal services: `gpedit.msc`.
- Users, groups, Guest, Administrator, account rename: `lusrmgr.msc`.
- Services, startup type, logon account: `services.msc`.
- Event logs, system logs, log size/retention: `eventvwr.msc`.
- Shared folders, default shares: `fsmgmt.msc`.
- Registry paths or `HKEY`: `regedit`.
- Installed updates, programs/features: `control`.

## Execution Pattern

1. Open the inferred GUI tool from the guest interactive desktop.
2. Maximize the window.
3. Navigate using stable keyboard paths, tree expansion, search, or visible text.
4. Prefer opening a properties dialog when the list row truncates the value.
5. Capture screenshot with guest-side PyAutoGUI.
6. Verify the screenshot against the row requirement.

If navigation by exact coordinates is necessary, use coordinates relative to the active window and current screenshot, not host screen coordinates.

## Adaptive Result Writing

Use neutral, evidence-based language:

- State the observed configuration.
- State whether it meets the expected result.
- Avoid operation-step narration such as "opened gpedit and clicked...".
- Do not write `实际:` or `证据:` prefixes unless the user asks for that style.

## Failure Handling

Do not skip an unmatched row just because no pattern matched.

Mark `needs_manual_confirmation` only when:

- The GUI tool is absent.
- The account lacks permissions.
- The window cannot be displayed in the interactive session.
- The target item is not present after reasonable search.
- Screenshots cannot prove the requirement after resizing/opening details.

Record the exact blocker and keep any diagnostic screenshots separate from final evidence.
