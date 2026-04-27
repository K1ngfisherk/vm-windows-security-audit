#!/usr/bin/env python3
"""Guest-side PyAutoGUI runner for Windows security checklist evidence.

This runner is intentionally pragmatic: it contains stable actions learned from
Windows Server 2012 MMC workflows, plus an adaptive fallback for new rows. Run
it from the interactive guest desktop, preferably through a highest-privilege
scheduled task owned by the requested audit account.
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import time
from pathlib import Path
from typing import Any, Callable, Iterable

import pyautogui
import pygetwindow as gw
try:
    import pyperclip
except Exception:  # pragma: no cover - optional dependency guard for damaged guests.
    pyperclip = None


pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.12
CREATE_NEW_CONSOLE = 0x00000010


class Runner:
    def __init__(self, out_dir: Path, debug: bool = False):
        self.out_dir = out_dir
        self.out_dir.mkdir(parents=True, exist_ok=True)
        self.tmp_dir = self.out_dir / "tmp"
        self.tmp_dir.mkdir(parents=True, exist_ok=True)
        self.debug = debug
        self.log_path = self.tmp_dir / "guest_gui_runner.log"
        self.log_path.write_text("", encoding="utf-8")

    def log(self, message: str) -> None:
        with self.log_path.open("a", encoding="utf-8") as handle:
            handle.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")

    def wait(self, seconds: float = 1.0) -> None:
        time.sleep(seconds)

    def kill_ui(self) -> None:
        for name in [
            "mmc.exe",
            "regedit.exe",
            "control.exe",
            "OptionalFeatures.exe",
            "rundll32.exe",
            "eventvwr.exe",
            "cmd.exe",
        ]:
            subprocess.call(
                ["taskkill", "/f", "/im", name],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        self.wait(1.0)

    def windows(self, title_parts: str | Iterable[str]) -> list[Any]:
        if isinstance(title_parts, str):
            parts = [title_parts]
        else:
            parts = list(title_parts)
        return [win for win in gw.getAllWindows() if any(part in win.title for part in parts)]

    def wait_window(self, title_parts: str | Iterable[str], timeout: int = 25) -> Any:
        deadline = time.time() + timeout
        while time.time() < deadline:
            wins = self.windows(title_parts)
            if wins:
                return wins[0]
            self.wait(0.4)
        raise RuntimeError(f"window not found: {title_parts}")

    def active_window(self, title_parts: str | Iterable[str]) -> Any:
        wins = self.windows(title_parts)
        if not wins:
            raise RuntimeError(f"window disappeared: {title_parts}")
        return wins[0]

    def maximize(self, win: Any) -> Any:
        self.log(f"maximize title={win.title!r} left={win.left} top={win.top} size={win.width}x{win.height}")
        try:
            win.activate()
        except Exception as exc:
            self.log(f"activate failed: {exc!r}")
        self.wait(0.35)
        try:
            win.maximize()
        except Exception as exc:
            self.log(f"maximize failed: {exc!r}")
        self.wait(0.8)
        return self.active_window(win.title)

    def click(self, win: Any, x: int, y: int, delay: float = 0.4) -> None:
        pyautogui.click(win.left + x, win.top + y)
        self.wait(delay)

    def double_click(self, win: Any, x: int, y: int, delay: float = 0.8) -> None:
        pyautogui.doubleClick(win.left + x, win.top + y)
        self.wait(delay)

    def shot(self, filename: str, *, final: bool = True) -> str:
        root = self.out_dir if final else self.tmp_dir
        target = root / Path(filename).name
        target.parent.mkdir(parents=True, exist_ok=True)
        self.wait(0.7)
        image = pyautogui.screenshot()
        image.save(target)
        kind = "shot" if final else "tmp-shot"
        self.log(f"{kind} {target}")
        return str(target)

    def tmp_shot(self, filename: str) -> str:
        return self.shot(filename, final=False)

    def cleanup_tmp(self) -> None:
        cleanup_tmp_dir(self.out_dir)

    def copy_control_panel_location(self) -> str:
        if pyperclip is None:
            return ""
        try:
            pyautogui.hotkey("alt", "d")
            self.wait(0.2)
            pyautogui.hotkey("ctrl", "c")
            self.wait(0.2)
            location = pyperclip.paste() or ""
            pyautogui.press("esc")
            self.log(f"control-panel-location {location!r}")
            return location
        except Exception as exc:
            self.log(f"control-panel-location failed: {exc!r}")
            return ""

    def control_panel_location_has(self, win: Any, tokens: Iterable[str]) -> bool:
        title = getattr(win, "title", "")
        if any(token.lower() in title.lower() for token in tokens):
            return True
        location = self.copy_control_panel_location()
        return any(token.lower() in location.lower() for token in tokens)

    def open_mmc(self, msc_name: str, title_parts: str | Iterable[str], wait: float = 4.0) -> Any:
        self.kill_ui()
        subprocess.Popen([r"C:\Windows\System32\mmc.exe", msc_name])
        self.wait(wait)
        return self.maximize(self.wait_window(title_parts))

    def open_regedit_key(self, key: str) -> Any:
        self.kill_ui()
        subprocess.call(
            rf'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit" /v LastKey /t REG_SZ /d "{key}" /f >nul',
            shell=True,
        )
        subprocess.Popen([r"C:\Windows\regedit.exe"])
        self.wait(4.0)
        return self.maximize(self.wait_window(["注册表编辑器", "Registry Editor"]))

    def visible_cmd(self, row: int, name: str, commands: list[str], filename: str) -> list[str]:
        self.kill_ui()
        bat = self.tmp_dir / f"row{row:02d}_{name}_visible.cmd"
        lines = [
            "@echo off",
            "chcp 936 >nul",
            f"title ROW{row:02d}_{name}_VISIBLE",
            "color 0F",
            "mode con: cols=150 lines=42",
            f"echo === Row {row}: {name} ===",
            "echo.",
        ]
        for command in commands:
            lines += [f"echo ^> {command}", command, "echo."]
        lines += ["echo === Evidence window kept open for screenshot ==="]
        bat.write_text("\r\n".join(lines), encoding="mbcs")
        subprocess.Popen(["cmd.exe", "/k", str(bat)], creationflags=CREATE_NEW_CONSOLE)
        self.wait(4.0)
        try:
            win = self.wait_window(f"ROW{row:02d}_{name}_VISIBLE", 10)
            self.maximize(win)
        except Exception as exc:
            self.log(f"visible cmd window not found: {exc!r}")
        return [self.shot(filename)]

    def open_secpol_account_child(self, child: str) -> Any:
        win = self.open_mmc("secpol.msc", ["本地安全策略", "Local Security Policy"])
        # Coordinates target the maximized Server 2012 Local Security Policy tree.
        self.click(win, 25, 116, 0.5)  # expand Account Policies
        if child == "password":
            self.click(win, 69, 135, 1.2)
        elif child == "lockout":
            self.click(win, 85, 153, 1.2)
        else:
            raise ValueError(child)
        return self.active_window(["本地安全策略", "Local Security Policy"])

    def open_secpol_local_child(self, child: str) -> Any:
        win = self.open_mmc("secpol.msc", ["本地安全策略", "Local Security Policy"])
        self.click(win, 54, 86, 0.3)
        pyautogui.press("home")
        pyautogui.press("down", presses=2, interval=0.15)
        pyautogui.press("right")
        self.wait(0.3)
        if child == "audit":
            pyautogui.press("down")
        elif child == "security":
            pyautogui.press("down", presses=3, interval=0.15)
        elif child == "rights":
            pyautogui.press("down", presses=2, interval=0.15)
        else:
            raise ValueError(child)
        self.wait(1.2)
        return self.active_window(["本地安全策略", "Local Security Policy"])

    def open_gpedit_computer_admin_templates(self, debug_prefix: str | None = None) -> Any:
        win = self.open_mmc("gpedit.msc", ["本地组策略编辑器", "Local Group Policy Editor"], wait=5)
        self.click(win, 95, 92, 0.2)
        pyautogui.press("home")
        self.wait(0.2)
        pyautogui.press("right")
        self.wait(0.2)
        pyautogui.press("down")  # Computer Configuration
        self.wait(0.2)
        pyautogui.press("right")
        self.wait(0.2)
        pyautogui.press("down", presses=2, interval=0.12)  # Administrative Templates
        self.wait(0.2)
        pyautogui.press("right")
        self.wait(0.5)
        if debug_prefix:
            self.tmp_shot(f"{debug_prefix}_admin_templates.png")
        return self.active_window(["本地组策略编辑器", "Local Group Policy Editor"])

    def open_gpedit_computer_rdsh(self, folder_name: str) -> Any:
        win = self.open_gpedit_computer_admin_templates()
        # Use details pane; this was more stable than deep tree clicking on the
        # target Server 2012 image.
        self.click(win, 455, 150, 0.2)
        self.double_click(win, 430, 146, 0.8)  # Windows Components
        self.click(win, 455, 150, 0.2)
        pyautogui.press("end")
        self.wait(0.4)
        self.double_click(win, 430, 640, 0.8)  # Remote Desktop Services
        self.double_click(win, 455, 164, 0.8)  # Remote Desktop Session Host
        y_by_folder = {
            "security": 164,
            "session_time_limits": 204,
            "connections": 222,
        }
        self.double_click(win, 455, y_by_folder[folder_name], 0.8)
        return self.active_window(["本地组策略编辑器", "Local Group Policy Editor"])

    def adaptive_open_and_capture(self, item: dict[str, Any]) -> list[str]:
        tool = str(item.get("tool") or "control").split(";")[0]
        if tool in {"cmd", "powershell"}:
            return self.visible_cmd(item["row"], item.get("check_id") or "adaptive", ["whoami"], item["evidence"][0])
        if tool == "regedit":
            subprocess.Popen([r"C:\Windows\regedit.exe"])
            self.wait(3)
            try:
                self.maximize(self.wait_window(["注册表编辑器", "Registry Editor"], 10))
            except Exception:
                pass
        elif tool == "control":
            subprocess.Popen([r"C:\Windows\System32\control.exe"])
            self.wait(3)
        elif tool.endswith(".msc"):
            title = {
                "secpol.msc": ["本地安全策略", "Local Security Policy"],
                "gpedit.msc": ["本地组策略编辑器", "Local Group Policy Editor"],
                "lusrmgr.msc": ["本地用户和组", "Local Users and Groups"],
                "services.msc": ["服务", "Services"],
                "eventvwr.msc": ["事件查看器", "Event Viewer"],
                "fsmgmt.msc": ["共享文件夹", "Shared Folders"],
            }.get(tool, tool)
            self.open_mmc(tool, title)
        else:
            subprocess.Popen(tool, shell=True)
            self.wait(3)
        return [self.shot(item["evidence"][0])]


def first_evidence(item: dict[str, Any], default: str) -> str:
    evidence = item.get("evidence") or [default]
    return Path(evidence[0]).name


def all_evidence(item: dict[str, Any]) -> list[str]:
    return [Path(name).name for name in item.get("evidence", [])]


def action_identity_users(r: Runner, item: dict[str, Any]) -> list[str]:
    return r.visible_cmd(
        item["row"],
        "identity_users",
        [
            "whoami",
            "query user",
            "net user",
            "net accounts",
            r'reg query "HKLM\SYSTEM\CurrentControlSet\Control\Lsa" /v LimitBlankPasswordUse',
        ],
        first_evidence(item, f"row{item['row']:02d}_identity_users.png"),
    )


def action_secpol_password_policy(r: Runner, item: dict[str, Any]) -> list[str]:
    r.open_secpol_account_child("password")
    path = r.shot(first_evidence(item, f"row{item['row']:02d}_secpol_password_policy.png"))
    r.kill_ui()
    return [path]


def action_secpol_account_lockout_policy(r: Runner, item: dict[str, Any]) -> list[str]:
    r.open_secpol_account_child("lockout")
    path = r.shot(first_evidence(item, f"row{item['row']:02d}_secpol_account_lockout_policy.png"))
    r.kill_ui()
    return [path]


def action_gpedit_rdp_client_connection_encryption_level(r: Runner, item: dict[str, Any]) -> list[str]:
    win = r.open_gpedit_computer_rdsh("security")
    r.double_click(win, 505, 166, 1.0)
    path = r.shot(first_evidence(item, f"row{item['row']:02d}_gpedit_rdp_client_connection_encryption_level.png"))
    r.kill_ui()
    return [path]


def action_lusrmgr_users(r: Runner, item: dict[str, Any]) -> list[str]:
    win = r.open_mmc("lusrmgr.msc", ["本地用户和组", "Local Users and Groups"])
    r.click(win, 50, 104, 1.0)
    path = r.shot(first_evidence(item, f"row{item['row']:02d}_lusrmgr_users.png"))
    r.kill_ui()
    return [path]


def action_admin_auth_method(r: Runner, item: dict[str, Any]) -> list[str]:
    return r.visible_cmd(
        item["row"],
        "admin_auth_method",
        [
            "whoami",
            "query user",
            "net localgroup administrators",
            r'reg query "HKLM\SYSTEM\CurrentControlSet\Control\Lsa" /v LimitBlankPasswordUse',
        ],
        first_evidence(item, f"row{item['row']:02d}_admin_auth_method.png"),
    )


def action_fsmgmt_shares(r: Runner, item: dict[str, Any]) -> list[str]:
    win = r.open_mmc("fsmgmt.msc", ["共享文件夹", "Shared Folders"])
    r.click(win, 51, 104, 1.0)
    path = r.shot(first_evidence(item, f"row{item['row']:02d}_fsmgmt_shares.png"))
    r.kill_ui()
    return [path]


def action_regedit_lsa_restrictanonymous(r: Runner, item: dict[str, Any]) -> list[str]:
    win = r.open_regedit_key(r"计算机\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa")
    # Widen name column so restrictanonymous and restrictanonymoussam are readable.
    pyautogui.moveTo(win.left + 344, win.top + 54)
    pyautogui.dragTo(win.left + 500, win.top + 54, duration=0.3, button="left")
    r.wait(0.5)
    path = r.shot(first_evidence(item, f"row{item['row']:02d}_regedit_lsa_restrictanonymous.png"))
    r.kill_ui()
    return [path]


def action_privilege_separation(r: Runner, item: dict[str, Any]) -> list[str]:
    names = all_evidence(item)
    if not names:
        names = [
            f"row{item['row']:02d}_lusrmgr_administrators_members.png",
            f"row{item['row']:02d}_services_mysql_logon_account.png",
        ]
    captured: list[str] = []

    win = r.open_mmc("lusrmgr.msc", ["本地用户和组", "Local Users and Groups"])
    r.click(win, 43, 121, 0.6)
    r.double_click(win, 260, 124, 1.0)
    captured.append(r.shot(names[0]))
    r.kill_ui()

    if len(names) > 1:
        win = r.open_mmc("services.msc", ["服务", "Services"])
        r.click(win, 420, 146, 0.2)
        pyautogui.press("home")
        r.wait(0.2)
        pyautogui.write("MySQLa", interval=0.03)
        r.wait(0.8)
        pyautogui.press("enter")
        r.wait(1.0)
        pyautogui.hotkey("ctrl", "tab")
        r.wait(0.5)
        captured.append(r.shot(names[1]))
        r.kill_ui()
    return captured


def action_lusrmgr_default_accounts(r: Runner, item: dict[str, Any]) -> list[str]:
    names = all_evidence(item)
    if not names:
        names = [
            f"row{item['row']:02d}_lusrmgr_default_accounts.png",
            f"row{item['row']:02d}_lusrmgr_guest_properties_disabled.png",
        ]
    captured: list[str] = []
    win = r.open_mmc("lusrmgr.msc", ["本地用户和组", "Local Users and Groups"])
    r.click(win, 50, 104, 1.0)
    captured.append(r.shot(names[0]))
    if len(names) > 1:
        r.double_click(win, 235, 145, 1.0)
        captured.append(r.shot(names[1]))
    r.kill_ui()
    return captured


def action_lusrmgr_extra_accounts(r: Runner, item: dict[str, Any]) -> list[str]:
    return action_lusrmgr_users(r, item)


def action_secpol_audit_policy(r: Runner, item: dict[str, Any]) -> list[str]:
    r.open_secpol_local_child("audit")
    path = r.shot(first_evidence(item, f"row{item['row']:02d}_secpol_audit_policy.png"))
    r.kill_ui()
    return [path]


def action_eventvwr_system_log(r: Runner, item: dict[str, Any]) -> list[str]:
    r.kill_ui()
    subprocess.Popen([r"C:\Windows\System32\mmc.exe", "eventvwr.msc", "/c:System"])
    r.wait(6.0)
    r.maximize(r.wait_window(["事件查看器", "Event Viewer"], 25))
    path = r.shot(first_evidence(item, f"row{item['row']:02d}_eventvwr_system_log.png"))
    r.kill_ui()
    return [path]


def action_eventvwr_system_log_properties(r: Runner, item: dict[str, Any]) -> list[str]:
    r.kill_ui()
    subprocess.Popen([r"C:\Windows\System32\mmc.exe", "eventvwr.msc", "/c:System"])
    r.wait(6.0)
    win = r.maximize(r.wait_window(["事件查看器", "Event Viewer"], 25))
    r.click(win, 885, 257, 1.0)
    path = r.shot(first_evidence(item, f"row{item['row']:02d}_eventvwr_system_log_properties.png"))
    r.kill_ui()
    return [path]


def action_control_installed_updates(r: Runner, item: dict[str, Any]) -> list[str]:
    r.kill_ui()
    tokens = ["已安装更新", "Installed Updates"]
    titles = ["已安装更新", "Installed Updates", "程序和功能", "Programs and Features"]

    try:
        subprocess.Popen(
            [
                r"C:\Windows\System32\rundll32.exe",
                "shell32.dll,Control_RunDLL",
                "appwiz.cpl,,2",
            ]
        )
        r.wait(5.0)
        win = r.maximize(r.wait_window(titles, 10))
    except Exception as exc:
        r.log(f"direct installed-updates route failed: {exc!r}; falling back to appwiz link")
        r.kill_ui()
        subprocess.Popen([r"C:\Windows\System32\control.exe", "appwiz.cpl"])
        r.wait(5.0)
        win = r.maximize(r.wait_window(titles, 25))

    if not r.control_panel_location_has(win, tokens):
        # Server 2012 often opens Programs and Features first; use the left
        # navigation link with coordinates relative to the actual window.
        r.click(win, 86, 126, 4.0)
        win = r.active_window(titles)

    if not r.control_panel_location_has(win, tokens):
        raise RuntimeError("Programs and Features did not navigate to Installed Updates")

    path = r.shot(first_evidence(item, f"row{item['row']:02d}_control_installed_updates.png"))
    r.kill_ui()
    return [path]


def action_services_remote_registry(r: Runner, item: dict[str, Any]) -> list[str]:
    win = r.open_mmc("services.msc", ["服务", "Services"])
    r.click(win, 410, 145, 0.2)
    pyautogui.write("Remote Registry", interval=0.02)
    r.wait(1.0)
    r.click(win, 430, 278, 0.5)
    path = r.shot(first_evidence(item, f"row{item['row']:02d}_services_remote_registry.png"))
    r.kill_ui()
    return [path]


def action_gpedit_idle_session_limit(r: Runner, item: dict[str, Any]) -> list[str]:
    win = r.open_gpedit_computer_rdsh("session_time_limits")
    r.double_click(win, 505, 166, 1.0)
    path = r.shot(first_evidence(item, f"row{item['row']:02d}_gpedit_idle_session_limit.png"))
    r.kill_ui()
    return [path]


def action_gpedit_limit_number_of_connections(r: Runner, item: dict[str, Any]) -> list[str]:
    win = r.open_gpedit_computer_rdsh("connections")
    r.double_click(win, 505, 222, 1.0)
    path = r.shot(first_evidence(item, f"row{item['row']:02d}_gpedit_limit_number_of_connections.png"))
    r.kill_ui()
    return [path]


def action_secpol_security_options(r: Runner, item: dict[str, Any]) -> list[str]:
    r.open_secpol_local_child("security")
    path = r.shot(first_evidence(item, f"row{item['row']:02d}_secpol_security_options.png"))
    r.kill_ui()
    return [path]


def action_adaptive_gui(r: Runner, item: dict[str, Any]) -> list[str]:
    return r.adaptive_open_and_capture(item)


ACTIONS: dict[str, Callable[[Runner, dict[str, Any]], list[str]]] = {
    "identity_users": action_identity_users,
    "secpol_password_policy": action_secpol_password_policy,
    "secpol_account_lockout_policy": action_secpol_account_lockout_policy,
    "gpedit_rdp_client_connection_encryption_level": action_gpedit_rdp_client_connection_encryption_level,
    "lusrmgr_users": action_lusrmgr_users,
    "admin_auth_method": action_admin_auth_method,
    "fsmgmt_shares": action_fsmgmt_shares,
    "regedit_lsa_restrictanonymous": action_regedit_lsa_restrictanonymous,
    "privilege_separation": action_privilege_separation,
    "lusrmgr_default_accounts": action_lusrmgr_default_accounts,
    "lusrmgr_extra_accounts": action_lusrmgr_extra_accounts,
    "secpol_audit_policy": action_secpol_audit_policy,
    "eventvwr_system_log": action_eventvwr_system_log,
    "eventvwr_system_log_properties": action_eventvwr_system_log_properties,
    "control_installed_updates": action_control_installed_updates,
    "services_remote_registry": action_services_remote_registry,
    "gpedit_idle_session_limit": action_gpedit_idle_session_limit,
    "gpedit_limit_number_of_connections": action_gpedit_limit_number_of_connections,
    "secpol_security_options": action_secpol_security_options,
    "adaptive_gui": action_adaptive_gui,
}


def run_plan(plan_path: Path, out_dir: Path, only_row: int | None = None, debug: bool = False) -> dict[str, Any]:
    plan = json.loads(plan_path.read_text(encoding="utf-8"))
    runner = Runner(out_dir, debug=debug)
    results = []

    for item in plan["items"]:
        if only_row is not None and item["row"] != only_row:
            continue
        action_name = item.get("gui_action") or "adaptive_gui"
        action = ACTIONS.get(action_name, action_adaptive_gui)
        try:
            runner.log(f"start row={item['row']} action={action_name} lane={item.get('lane')}")
            evidence_paths = action(runner, item)
            results.append(
                {
                    "row": item["row"],
                    "status": "captured",
                    "action": action_name,
                    "lane": item.get("lane"),
                    "evidence": [Path(path).name for path in evidence_paths],
                }
            )
            runner.log(f"done row={item['row']} evidence={evidence_paths}")
        except Exception as exc:
            runner.log(f"ERROR row={item.get('row')} action={action_name}: {type(exc).__name__}: {exc}")
            try:
                error_name = f"row{int(item['row']):02d}_error_{action_name}.png"
                error_path = runner.tmp_shot(error_name)
                diagnostics = [Path(error_path).name]
            except Exception:
                diagnostics = []
            runner.kill_ui()
            results.append(
                {
                    "row": item.get("row"),
                    "status": "error",
                    "action": action_name,
                    "lane": item.get("lane"),
                    "error": f"{type(exc).__name__}: {exc}",
                    "evidence": [],
                    "diagnostics": diagnostics,
                }
            )
    return {"results": results, "log": str(runner.log_path), "tmp_dir": str(runner.tmp_dir)}


def cleanup_tmp_dir(out_dir: Path) -> None:
    tmp_dir = out_dir / "tmp"
    if tmp_dir.exists():
        shutil.rmtree(tmp_dir, ignore_errors=True)


def path_is_under(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root.resolve())
        return True
    except ValueError:
        return False


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--plan", required=True, type=Path)
    parser.add_argument("--out-dir", required=True, type=Path)
    parser.add_argument("--only-row", type=int)
    parser.add_argument("--result-json", type=Path)
    parser.add_argument("--debug", action="store_true")
    parser.add_argument("--keep-tmp", action="store_true", help="Keep out-dir/tmp diagnostics after the run.")
    args = parser.parse_args()

    if args.result_json:
        tmp_dir = args.out_dir / "tmp"
        if not path_is_under(args.result_json, tmp_dir):
            parser.error("--result-json must be written under --out-dir/tmp; use stdout for non-file results.")
        if not args.keep_tmp:
            parser.error("--result-json lives under --out-dir/tmp; use --keep-tmp and let the caller clean tmp after reading it.")

    result = run_plan(args.plan, args.out_dir, args.only_row, args.debug)
    if args.result_json:
        args.result_json.parent.mkdir(parents=True, exist_ok=True)
        args.result_json.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    if not args.keep_tmp:
        cleanup_tmp_dir(args.out_dir)
        result["tmp_removed"] = True
        if args.result_json and path_is_under(args.result_json, args.out_dir):
            args.result_json.unlink(missing_ok=True)
        elif args.result_json:
            args.result_json.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
