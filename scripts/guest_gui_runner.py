#!/usr/bin/env python3
"""Guest-side PyAutoGUI runner for Windows security checklist evidence.

This runner is intentionally pragmatic: it contains stable actions learned from
Windows Server 2012 MMC workflows, plus an adaptive fallback for new rows. Run
it from the interactive guest desktop, preferably through a highest-privilege
scheduled task owned by the requested audit account.
"""

from __future__ import annotations

import argparse
import ctypes
import ctypes.wintypes as wintypes
import json
import re
import shutil
import statistics
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
ADMIN_INTERVIEW_FINDING = "涉及访谈管理员，未进行操作"
ADMIN_INTERVIEW_RESULT = "未检查"

UIA_DUMP_SCRIPT = r"""
$ErrorActionPreference = "SilentlyContinue"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class NativeMethods {
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
}
"@
$handle = [NativeMethods]::GetForegroundWindow()
$root = [System.Windows.Automation.AutomationElement]::FromHandle($handle)
if ($null -eq $root) { $root = [System.Windows.Automation.AutomationElement]::RootElement }
$all = $root.FindAll([System.Windows.Automation.TreeScope]::Subtree, [System.Windows.Automation.Condition]::TrueCondition)
$result = New-Object System.Collections.Generic.List[object]
$max = [Math]::Min($all.Count, 1800)
for ($i = 0; $i -lt $max; $i++) {
  $e = $all.Item($i)
  $rect = $e.Current.BoundingRectangle
  if ($rect.Width -le 1 -or $rect.Height -le 1) { continue }
  $name = [string]$e.Current.Name
  $value = ""
  $pattern = $null
  try {
    if ($e.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$pattern)) {
      $value = [string]$pattern.Current.Value
    }
  } catch {}
  $legacyName = ""
  $legacyValue = ""
  $legacyDescription = ""
  $legacyPattern = $null
  try {
    if ($e.TryGetCurrentPattern([System.Windows.Automation.LegacyIAccessiblePattern]::Pattern, [ref]$legacyPattern)) {
      $legacyName = [string]$legacyPattern.Current.Name
      $legacyValue = [string]$legacyPattern.Current.Value
      $legacyDescription = [string]$legacyPattern.Current.Description
    }
  } catch {}
  if (
    [string]::IsNullOrWhiteSpace($name) -and
    [string]::IsNullOrWhiteSpace($value) -and
    [string]::IsNullOrWhiteSpace($legacyName) -and
    [string]::IsNullOrWhiteSpace($legacyValue) -and
    [string]::IsNullOrWhiteSpace($legacyDescription)
  ) { continue }
  $result.Add([pscustomobject]@{
    name = $name
    value = $value
    legacyName = $legacyName
    legacyValue = $legacyValue
    legacyDescription = $legacyDescription
    automationId = [string]$e.Current.AutomationId
    className = [string]$e.Current.ClassName
    controlType = [string]$e.Current.ControlType.ProgrammaticName
    x = [int]$rect.X
    y = [int]$rect.Y
    w = [int]$rect.Width
    h = [int]$rect.Height
  })
}
$result | ConvertTo-Json -Compress -Depth 4
"""


class EvidenceValidationError(RuntimeError):
    """Raised when a screenshot exists but does not prove the requested row."""


class UserConfirmationRequiredError(EvidenceValidationError):
    """Raised when the only available path would violate the evidence contract."""


COMMAND_WINDOW_TITLE_TOKENS = (
    "windows powershell",
    "powershell",
    "windows terminal",
    "command prompt",
    "cmd.exe",
    "命令提示符",
    "_visible",
)

BAD_FOREGROUND_TITLE_TOKENS = COMMAND_WINDOW_TITLE_TOKENS + (
    "optionalfeatures",
    "windows features",
    "windows 功能",
)

GA_ROOT = 2

SEMANTIC_RULES: list[tuple[str, dict[str, Any]]] = [
    ("secpol_password_policy", {
        "must_all_any": [["本地安全策略", "local security policy", "secpol"], ["密码策略", "password policy", "密码必须符合复杂性要求"]],
        "forbid_any": ["本地组策略", "local group policy", "gpedit", "本地用户和组", "local users and groups"],
    }),
    ("secpol_account_lockout_policy", {
        "must_all_any": [["本地安全策略", "local security policy", "secpol"], ["帐户锁定策略", "账户锁定策略", "account lockout policy", "锁定阈值"]],
        "forbid_any": ["本地组策略", "local group policy", "gpedit", "本地用户和组", "local users and groups"],
    }),
    ("secpol_audit_policy", {
        "must_all_any": [
            ["本地安全策略", "local security policy", "secpol"],
            ["审核策略", "audit policy"],
            ["审核帐户登录事件", "审核账户登录事件", "audit account logon events", "审核登录事件", "audit logon events"],
        ],
        "forbid_any": ["本地组策略", "local group policy", "gpedit", "本地用户和组", "local users and groups"],
    }),
    ("secpol_security_options", {
        "must_all_any": [
            ["本地安全策略", "local security policy", "secpol"],
            ["安全选项", "security options"],
            ["交互式登录", "interactive logon", "关机: 清除虚拟内存页面文件", "clear virtual memory pagefile", "microsoft 网络客户端"],
        ],
        "forbid_any": ["本地组策略", "local group policy", "gpedit", "本地用户和组", "local users and groups"],
    }),
    ("gpedit_rdp_client_connection_encryption_level", {
        "must_any": ["加密", "encryption", "rdp", "remote desktop"],
        "forbid_any": ["本地用户和组", "local users and groups", "lusrmgr"],
    }),
    ("gpedit_idle_session_limit", {
        "must_any": ["会话", "空闲", "session", "idle", "time limit"],
        "forbid_any": ["本地用户和组", "local users and groups", "lusrmgr"],
    }),
    ("gpedit_limit_number_of_connections", {
        "must_any": ["连接", "connection"],
        "forbid_any": ["本地用户和组", "local users and groups", "lusrmgr"],
    }),
    ("lusrmgr_guest_properties_disabled", {
        "must_all_any": [["guest", "来宾"], ["属性", "properties"], ["账户已禁用", "帐户已禁用", "account is disabled"]],
        "forbid_any": ["本地组策略", "local group policy", "gpedit"],
    }),
    ("lusrmgr_administrators_members", {
        "must_all_any": [["administrator", "administrators", "管理员"], ["属性", "properties", "成员", "members"]],
        "forbid_any": ["本地组策略", "local group policy", "gpedit"],
    }),
    ("lusrmgr_default_accounts", {
        "must_all_any": [
            ["本地用户和组", "local users and groups", "lusrmgr"],
            ["administrator"],
            ["guest"],
        ],
        "forbid_any": ["本地组策略", "local group policy", "gpedit"],
    }),
    ("lusrmgr_users", {
        "must_all_any": [
            ["本地用户和组", "local users and groups", "lusrmgr"],
            ["用户", "users"],
            ["administrator"],
            ["guest"],
        ],
        "forbid_any": ["本地组策略", "local group policy", "gpedit"],
    }),
    ("lusrmgr", {
        "must_any": ["本地用户和组", "local users and groups", "lusrmgr", "用户", "users"],
        "forbid_any": ["本地组策略", "local group policy", "gpedit"],
    }),
    ("secpol", {
        "must_any": ["本地安全策略", "local security policy", "secpol"],
        "forbid_any": ["本地组策略", "local group policy", "gpedit", "本地用户和组", "local users and groups"],
    }),
    ("eventvwr_system_log_properties", {
        "must_any": ["system", "系统", "属性", "properties", "event viewer", "事件查看器"],
        "forbid_any": ["本地组策略", "local group policy", "本地用户和组", "local users and groups"],
    }),
    ("eventvwr", {
        "must_any": ["event viewer", "事件查看器"],
        "forbid_any": ["本地组策略", "local group policy", "本地用户和组", "local users and groups"],
    }),
    ("services_remote_registry", {
        "must_all_any": [["services", "服务"], ["remote registry", "remote", "registry"]],
        "forbid_any": ["本地组策略", "local group policy", "本地用户和组", "local users and groups"],
    }),
    ("services", {
        "must_any": ["services", "服务"],
        "forbid_any": ["本地组策略", "local group policy", "本地用户和组", "local users and groups"],
    }),
    ("fsmgmt", {
        "must_any": ["shared folders", "共享文件夹", "共享"],
        "forbid_any": ["本地组策略", "local group policy", "本地用户和组", "local users and groups"],
    }),
    ("regedit_lsa_restrictanonymous", {
        "must_all_any": [
            ["registry editor", "注册表编辑器", "regedit"],
            ["restrictanonymous"],
            ["restrictanonymoussam"],
            ["reg_dword", "dword"],
        ],
        "forbid_any": ["本地组策略", "local group policy", "本地用户和组", "local users and groups"],
    }),
    ("regedit", {
        "must_any": ["registry editor", "注册表编辑器", "regedit"],
        "forbid_any": ["本地组策略", "local group policy", "本地用户和组", "local users and groups"],
    }),
    ("control_installed_updates", {
        "must_all_any": [
            ["installed updates", "已安装更新", "已安装的更新"],
            ["kb", "microsoft windows", "update for microsoft windows", "security update", "更新"],
        ],
        "forbid_any": [
            "本地组策略",
            "local group policy",
            "本地用户和组",
            "local users and groups",
            "windows features",
            "windows 功能",
            "optionalfeatures",
            "卸载或更改程序",
            "uninstall or change a program",
        ],
    }),
]


class Runner:
    def __init__(self, out_dir: Path, debug: bool = False):
        self.out_dir = out_dir
        self.out_dir.mkdir(parents=True, exist_ok=True)
        self.tmp_dir = self.out_dir / "tmp"
        self.tmp_dir.mkdir(parents=True, exist_ok=True)
        self.debug = debug
        self.log_path = self.tmp_dir / "guest_gui_runner.log"
        self.log_path.write_text("", encoding="utf-8")
        self.delivery_results: dict[int, dict[str, str]] = {}

    def log(self, message: str) -> None:
        with self.log_path.open("a", encoding="utf-8") as handle:
            handle.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")

    def wait(self, seconds: float = 1.0) -> None:
        time.sleep(seconds)

    def item_keywords(self, item: dict[str, Any], extra: Iterable[str] = ()) -> list[str]:
        tokens: list[str] = []
        tokens.extend(str(token) for token in item.get("keywords", []) if str(token).strip())
        source = item.get("source") or {}
        for key in ("item", "expected", "operation", "remediation"):
            text = str(source.get(key) or "")
            tokens.extend(re.findall(r"[\u4e00-\u9fffA-Za-z0-9_$\\.-]{2,}", text))
        tokens.extend(str(token) for token in extra if str(token).strip())

        deduped: list[str] = []
        seen: set[str] = set()
        for token in tokens:
            token = token.strip()
            if not token or "?" in token:
                continue
            if len(token) > 24 and not re.search(r"[A-Za-z0-9_$\\.-]", token):
                continue
            key = token.casefold()
            if key in seen:
                continue
            seen.add(key)
            deduped.append(token)
        return deduped[:24]

    def uia_elements(self) -> list[dict[str, Any]]:
        try:
            proc = subprocess.run(
                [
                    r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    UIA_DUMP_SCRIPT,
                ],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=12,
            )
        except Exception as exc:
            self.log(f"uia dump failed: {exc!r}")
            return []
        if proc.returncode != 0:
            self.log(f"uia dump returned {proc.returncode}: {proc.stderr.decode('utf-8', errors='ignore')}")
            return []
        text = proc.stdout.decode("utf-8", errors="ignore").strip()
        if not text:
            return []
        try:
            parsed = json.loads(text)
        except Exception as exc:
            self.log(f"uia json parse failed: {exc!r}; raw={text[:500]!r}")
            return []
        if isinstance(parsed, dict):
            return [parsed]
        if isinstance(parsed, list):
            return [entry for entry in parsed if isinstance(entry, dict)]
        return []

    def element_text(self, element: dict[str, Any]) -> str:
        return " ".join(
            [
                str(element.get("name") or ""),
                str(element.get("value") or ""),
                str(element.get("legacyName") or ""),
                str(element.get("legacyValue") or ""),
                str(element.get("legacyDescription") or ""),
            ]
        )

    def visible_text(self) -> str:
        snapshot = self.window_snapshot()
        titles = [snapshot.get("active_title", "")] + [w["title"] for w in snapshot["windows"]]
        element_text = []
        for element in self.uia_elements():
            element_text.append(self.element_text(element))
        list_text = [str(item.get("row_text") or item.get("text") or "") for item in self.listview_text_items()]
        return " ".join(titles + element_text + list_text)

    def require_visible_keywords(self, item: dict[str, Any], phase: str, extra: Iterable[str] = (), min_hits: int = 1) -> list[str]:
        keywords = self.item_keywords(item, extra)
        text = self.visible_text().casefold()
        hits = [token for token in keywords if token.casefold() in text]
        self.log(f"keyword-guard phase={phase} keywords={keywords} hits={hits}")
        if len(hits) < min_hits:
            diag = self.tmp_shot(f"row{int(item['row']):02d}_{phase}_keyword_guard.png")
            raise EvidenceValidationError(
                f"Keyword guard failed before PyAutoGUI click at {phase}. "
                f"Expected one of {keywords}; hits={hits}; diagnostic={Path(diag).name}"
            )
        return hits

    def require_visible_token_groups(
        self,
        item: dict[str, Any],
        phase: str,
        groups: Iterable[Iterable[str]],
    ) -> list[list[str]]:
        text = self.visible_text().casefold()
        hits: list[list[str]] = []
        missing: list[list[str]] = []
        normalized_groups = [[str(token).strip() for token in group if str(token).strip()] for group in groups]
        for group in normalized_groups:
            group_hits = [token for token in group if token.casefold() in text]
            hits.append(group_hits)
            if not group_hits:
                missing.append(group)
        self.log(f"token-group-guard phase={phase} groups={normalized_groups} hits={hits} missing={missing}")
        if missing:
            diag = self.tmp_shot(f"row{int(item['row']):02d}_{phase}_token_group_guard.png")
            raise EvidenceValidationError(
                f"Keyword group guard failed at {phase}. "
                f"Missing groups={missing}; diagnostic={Path(diag).name}"
            )
        return hits

    def is_clickable_text_element(self, element: dict[str, Any], *, allow_window: bool = False) -> bool:
        control_type = str(element.get("controlType") or "").casefold()
        width = int(element.get("w") or 0)
        height = int(element.get("h") or 0)
        if width <= 1 or height <= 1:
            return False
        if "window" in control_type and not allow_window:
            return False
        screen_w, screen_h = pyautogui.size()
        if width * height > screen_w * screen_h * 0.45 and not allow_window:
            return False
        return True

    def window_rect_for_hwnd(self, hwnd: int | None = None) -> tuple[int, int, int, int, int]:
        user32 = ctypes.windll.user32
        if hwnd is None:
            hwnd = int(user32.GetForegroundWindow())
        if not hwnd:
            raise EvidenceValidationError("No foreground window is available for GUI interaction")
        root = int(user32.GetAncestor(hwnd, GA_ROOT)) or int(hwnd)
        rect = wintypes.RECT()
        if not user32.GetWindowRect(root, ctypes.byref(rect)):
            raise EvidenceValidationError(f"Could not read target window rectangle for hwnd={root}")
        return root, int(rect.left), int(rect.top), int(rect.right), int(rect.bottom)

    def assert_point_in_window(self, x: int, y: int, *, hwnd: int | None = None, phase: str = "click") -> int:
        root, left, top, right, bottom = self.window_rect_for_hwnd(hwnd)
        if not (left <= x <= right and top <= y <= bottom):
            diag = self.tmp_shot(f"{phase}_out_of_window_click.png")
            raise EvidenceValidationError(
                f"Refusing window-outside click at {phase}: point={x},{y}; "
                f"target_rect={left},{top},{right},{bottom}; diagnostic={Path(diag).name}"
            )
        return root

    def foreground_title(self) -> str:
        try:
            active = gw.getActiveWindow()
            return (getattr(active, "title", "") or "").strip()
        except Exception:
            return ""

    def assert_foreground_native_gui(self, item: dict[str, Any] | None, phase: str, *, allow_command_window: bool = False) -> None:
        title = self.foreground_title()
        folded = title.casefold()
        hits = [token for token in BAD_FOREGROUND_TITLE_TOKENS if token in folded]
        command_hits = [token for token in COMMAND_WINDOW_TITLE_TOKENS if token in folded]
        if (command_hits and not allow_command_window) or any(token in hits for token in ("optionalfeatures", "windows features", "windows 功能")):
            row = int(item["row"]) if item and item.get("row") is not None else 0
            prefix = f"row{row:02d}_" if row else ""
            diag = self.tmp_shot(f"{prefix}{phase}_bad_foreground.png")
            raise EvidenceValidationError(
                f"Refusing GUI evidence at {phase}: foreground window is {title!r}; "
                f"hits={hits}; diagnostic={Path(diag).name}"
            )

    def activate_hwnd(self, hwnd: int) -> None:
        try:
            user32 = ctypes.windll.user32
            user32.ShowWindow(hwnd, 5)
            user32.SetForegroundWindow(hwnd)
            self.wait(0.15)
        except Exception as exc:
            self.log(f"activate-hwnd failed hwnd={hwnd}: {exc!r}")

    def verify_navigation_after_click(self, item: dict[str, Any], phase: str, extra: Iterable[str]) -> None:
        tokens_by_phase = {
            "secpol_password_policy": [["密码策略", "Password Policy"], ["密码必须符合复杂性要求", "Maximum password age", "强制密码历史"]],
            "secpol_account_lockout_policy": [["帐户锁定策略", "账户锁定策略", "Account Lockout Policy"], ["锁定阈值", "Account lockout threshold"]],
            "secpol_audit_policy": [["审核策略", "Audit Policy"], ["审核登录事件", "Audit logon events", "审核帐户登录事件"]],
            "secpol_security_options": [["安全选项", "Security Options"], ["交互式登录", "Interactive logon", "关机: 清除虚拟内存页面文件"]],
            "lusrmgr_users_node": [["用户", "Users"], ["Administrator", "Guest"]],
            "fsmgmt_shares_node": [["共享", "Shares"], ["共享路径", "Shared Path", "描述", "Description"]],
            "gpedit_administrative_templates": [["管理模板", "Administrative Templates"], ["Windows 组件", "Windows Components", "所有设置", "All Settings"]],
            "gpedit_windows_components": [["Windows 组件", "Windows Components"], ["远程桌面服务", "Remote Desktop Services", "终端服务"]],
            "gpedit_remote_desktop_services": [["远程桌面服务", "Remote Desktop Services", "终端服务"], ["远程桌面会话主机", "Remote Desktop Session Host"]],
            "gpedit_remote_desktop_session_host": [["远程桌面会话主机", "Remote Desktop Session Host"], ["安全", "Security", "连接", "Connections", "会话时间限制"]],
        }
        groups = tokens_by_phase.get(phase)
        if groups:
            self.require_visible_token_groups(item, f"{phase}_after_click", groups)
            return
        search = [str(token).strip() for token in extra if str(token).strip()]
        if search:
            text = self.visible_text().casefold()
            if not any(token.casefold() in text for token in search):
                diag = self.tmp_shot(f"row{int(item['row']):02d}_{phase}_post_click_missing.png")
                raise EvidenceValidationError(
                    f"Post-click verification failed at {phase}; expected one of {search}; diagnostic={Path(diag).name}"
                )

    def find_text_element(self, keywords: Iterable[str], *, allow_window: bool = False) -> tuple[dict[str, Any], str] | None:
        search = [token.strip() for token in keywords if token and token.strip() and "?" not in token]
        elements = [element for element in self.uia_elements() if self.is_clickable_text_element(element, allow_window=allow_window)]
        elements = sorted(
            elements,
            key=lambda e: (
                0 if any(kind in str(e.get("controlType") or "") for kind in ["TreeItem", "ListItem", "DataItem", "Button", "MenuItem"]) else 1,
                int(e.get("w") or 0) * int(e.get("h") or 0),
            ),
        )
        for token in search:
            folded = token.casefold()
            for element in elements:
                haystack = self.element_text(element).casefold()
                if folded in haystack:
                    return element, token
        return None

    def find_listview_windows(self) -> list[dict[str, int | str]]:
        """Return visible SysListView32 controls under the active MMC window."""
        user32 = ctypes.windll.user32
        hwnd_root = user32.GetForegroundWindow()
        if not hwnd_root:
            return []

        result: list[dict[str, int | str]] = []
        seen_hwnds: set[int] = set()
        enum_proc_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

        def enum_child(hwnd: int, _lparam: int) -> bool:
            hwnd_int = int(hwnd)
            if hwnd_int in seen_hwnds:
                return True
            seen_hwnds.add(hwnd_int)
            class_name = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(hwnd, class_name, len(class_name))
            rect = wintypes.RECT()
            user32.GetWindowRect(hwnd, ctypes.byref(rect))
            width = int(rect.right - rect.left)
            height = int(rect.bottom - rect.top)
            if class_name.value == "SysListView32" and width > 20 and height > 20:
                result.append(
                    {
                        "hwnd": hwnd_int,
                        "class": class_name.value,
                        "x": int(rect.left),
                        "y": int(rect.top),
                        "w": width,
                        "h": height,
                    }
                )
            return True

        callback = enum_proc_type(enum_child)
        user32.EnumChildWindows(hwnd_root, callback, 0)
        return result

    def read_listview_items(self, hwnd: int, *, max_items: int = 3000, max_subitems: int = 8) -> list[dict[str, Any]]:
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        LVM_FIRST = 0x1000
        LVM_GETITEMCOUNT = LVM_FIRST + 4
        LVM_GETITEMTEXTW = LVM_FIRST + 115
        PROCESS_VM_OPERATION = 0x0008
        PROCESS_VM_READ = 0x0010
        PROCESS_VM_WRITE = 0x0020
        PROCESS_QUERY_INFORMATION = 0x0400
        MEM_COMMIT = 0x1000
        MEM_RESERVE = 0x2000
        MEM_RELEASE = 0x8000
        PAGE_READWRITE = 0x04

        class LVITEMW(ctypes.Structure):
            _fields_ = [
                ("mask", wintypes.UINT),
                ("iItem", ctypes.c_int),
                ("iSubItem", ctypes.c_int),
                ("state", wintypes.UINT),
                ("stateMask", wintypes.UINT),
                ("pszText", ctypes.c_void_p),
                ("cchTextMax", ctypes.c_int),
                ("iImage", ctypes.c_int),
                ("lParam", ctypes.c_void_p),
                ("iIndent", ctypes.c_int),
                ("iGroupId", ctypes.c_int),
                ("cColumns", wintypes.UINT),
                ("puColumns", ctypes.c_void_p),
                ("piColFmt", ctypes.c_void_p),
                ("iGroup", ctypes.c_int),
            ]

        user32.SendMessageW.restype = ctypes.c_ssize_t
        user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
        kernel32.OpenProcess.restype = wintypes.HANDLE
        kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        kernel32.VirtualAllocEx.restype = ctypes.c_void_p
        kernel32.VirtualAllocEx.argtypes = [wintypes.HANDLE, ctypes.c_void_p, ctypes.c_size_t, wintypes.DWORD, wintypes.DWORD]
        kernel32.VirtualFreeEx.argtypes = [wintypes.HANDLE, ctypes.c_void_p, ctypes.c_size_t, wintypes.DWORD]
        kernel32.ReadProcessMemory.argtypes = [
            wintypes.HANDLE,
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_size_t,
            ctypes.POINTER(ctypes.c_size_t),
        ]
        kernel32.WriteProcessMemory.argtypes = [
            wintypes.HANDLE,
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_size_t,
            ctypes.POINTER(ctypes.c_size_t),
        ]

        count = int(user32.SendMessageW(hwnd, LVM_GETITEMCOUNT, 0, 0))
        if count <= 0:
            return []
        count = min(count, max_items)
        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        process = kernel32.OpenProcess(
            PROCESS_VM_OPERATION | PROCESS_VM_READ | PROCESS_VM_WRITE | PROCESS_QUERY_INFORMATION,
            False,
            pid.value,
        )
        if not process:
            return []

        text_chars = 512
        text_bytes = text_chars * 2
        item_size = ctypes.sizeof(LVITEMW)
        remote_item = kernel32.VirtualAllocEx(process, None, item_size, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE)
        remote_text = kernel32.VirtualAllocEx(process, None, text_bytes, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE)
        items: list[dict[str, Any]] = []
        try:
            if not remote_item or not remote_text:
                return []
            for index in range(count):
                columns: list[str] = []
                for subitem in range(max(1, max_subitems)):
                    lvitem = LVITEMW()
                    lvitem.iItem = index
                    lvitem.iSubItem = subitem
                    lvitem.pszText = remote_text
                    lvitem.cchTextMax = text_chars
                    written = ctypes.c_size_t()
                    kernel32.WriteProcessMemory(process, remote_text, ctypes.create_string_buffer(text_bytes), text_bytes, ctypes.byref(written))
                    kernel32.WriteProcessMemory(process, remote_item, ctypes.byref(lvitem), item_size, ctypes.byref(written))
                    user32.SendMessageW(hwnd, LVM_GETITEMTEXTW, index, remote_item)
                    buffer = ctypes.create_string_buffer(text_bytes)
                    read = ctypes.c_size_t()
                    kernel32.ReadProcessMemory(process, remote_text, buffer, text_bytes, ctypes.byref(read))
                    raw = bytes(buffer)
                    text = raw.decode("utf-16-le", errors="ignore").split("\x00", 1)[0].strip()
                    if subitem == 0 and not text:
                        break
                    columns.append(text)
                columns = [text for text in columns if text]
                if columns:
                    items.append({"hwnd": hwnd, "index": index, "text": columns[0], "columns": columns, "row_text": " ".join(columns)})
        finally:
            if remote_item:
                kernel32.VirtualFreeEx(process, remote_item, 0, MEM_RELEASE)
            if remote_text:
                kernel32.VirtualFreeEx(process, remote_text, 0, MEM_RELEASE)
            kernel32.CloseHandle(process)
        return items

    def listview_text_items(self) -> list[dict[str, Any]]:
        items: list[dict[str, Any]] = []
        for view in self.find_listview_windows():
            hwnd = int(view["hwnd"])
            try:
                for item in self.read_listview_items(hwnd):
                    item.update({"list_rect": view})
                    items.append(item)
            except Exception as exc:
                self.log(f"win32-list-read failed hwnd={hwnd}: {exc!r}")
        return items

    def set_listview_column_widths(self, phase: str, widths: Iterable[int | None]) -> None:
        user32 = ctypes.windll.user32
        LVM_FIRST = 0x1000
        LVM_GETCOLUMNWIDTH = LVM_FIRST + 29
        LVM_SETCOLUMNWIDTH = LVM_FIRST + 30
        applied: list[dict[str, Any]] = []
        failures: list[dict[str, Any]] = []
        for view in self.find_listview_windows():
            hwnd = int(view["hwnd"])
            view_applied: list[dict[str, int | None]] = []
            for column, width in enumerate(widths):
                if width is None:
                    view_applied.append({"column": column, "requested": None, "actual": None})
                    continue
                requested_width = max(40, int(width))
                user32.SendMessageW(hwnd, LVM_SETCOLUMNWIDTH, column, requested_width)
                actual_width = int(user32.SendMessageW(hwnd, LVM_GETCOLUMNWIDTH, column, 0))
                view_applied.append({"column": column, "requested": requested_width, "actual": actual_width})
                if actual_width < requested_width - 8:
                    failures.append({"hwnd": hwnd, "column": column, "requested": requested_width, "actual": actual_width})
                self.wait(0.05)
            applied.append({"hwnd": hwnd, "rect": view, "widths": view_applied})
        self.log(f"listview-column-widths phase={phase} applied={applied}")
        self.wait(0.4)
        if not applied or failures:
            diag = self.tmp_shot(f"{phase}_column_width_failed.png")
            raise EvidenceValidationError(
                f"Could not expand required list view columns at {phase}. "
                f"applied={applied}; failures={failures}; diagnostic={Path(diag).name}"
            )

    def find_listview_item(self, keywords: Iterable[str]) -> tuple[dict[str, Any], str] | None:
        search = [token.strip() for token in keywords if token and token.strip() and "?" not in token]
        if not search:
            return None
        items = self.listview_text_items()
        for token in search:
            folded = token.casefold()
            for item in items:
                if str(item.get("row_text") or item.get("text") or "").casefold() == folded:
                    return item, token
        for token in search:
            folded = token.casefold()
            for item in items:
                if folded in str(item.get("row_text") or item.get("text") or "").casefold():
                    return item, token
        return None

    def click_list_text(
        self,
        item: dict[str, Any],
        phase: str,
        extra: Iterable[str],
        *,
        double: bool = False,
        delay: float = 0.8,
    ) -> None:
        row_keywords = self.item_keywords(item)
        keywords: list[str] = []
        for token in list(extra):
            token = str(token).strip()
            if token and token.casefold() not in {seen.casefold() for seen in keywords}:
                keywords.append(token)
        found = self.find_listview_item(keywords)
        self.log(f"win32-list-only phase={phase} row_keywords={row_keywords} search_keywords={keywords} found={found}")
        if not found:
            sample = [str(entry.get("text") or "") for entry in self.listview_text_items()[:80]]
            self.log(f"win32-list-only-miss phase={phase} sample={sample}")
            diag = self.tmp_shot(f"row{int(item['row']):02d}_{phase}_list_text_not_found.png")
            raise EvidenceValidationError(
                f"Could not find list view text before click at {phase}. "
                f"Keywords extracted from row: {row_keywords}; search_keywords={keywords}; diagnostic={Path(diag).name}"
            )
        list_item, token = found
        self.click_listview_item(list_item, token, double=double, delay=delay)
        if double:
            self.verify_navigation_after_click(item, phase, extra)

    def click_listview_item(
        self,
        item: dict[str, Any],
        token: str,
        *,
        double: bool = False,
        expand: bool = False,
        delay: float = 0.8,
    ) -> None:
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        LVM_FIRST = 0x1000
        LVM_ENSUREVISIBLE = LVM_FIRST + 19
        LVM_GETITEMRECT = LVM_FIRST + 14
        LVIR_BOUNDS = 0
        PROCESS_VM_OPERATION = 0x0008
        PROCESS_VM_READ = 0x0010
        PROCESS_VM_WRITE = 0x0020
        PROCESS_QUERY_INFORMATION = 0x0400
        MEM_COMMIT = 0x1000
        MEM_RESERVE = 0x2000
        MEM_RELEASE = 0x8000
        PAGE_READWRITE = 0x04

        hwnd = int(item["hwnd"])
        index = int(item["index"])
        user32.SendMessageW(hwnd, LVM_ENSUREVISIBLE, index, 0)
        self.wait(0.2)

        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        process = kernel32.OpenProcess(
            PROCESS_VM_OPERATION | PROCESS_VM_READ | PROCESS_VM_WRITE | PROCESS_QUERY_INFORMATION,
            False,
            pid.value,
        )
        if not process:
            raise EvidenceValidationError(f"Could not open list view process for token {token!r}")

        remote_rect = kernel32.VirtualAllocEx(
            process,
            None,
            ctypes.sizeof(wintypes.RECT),
            MEM_COMMIT | MEM_RESERVE,
            PAGE_READWRITE,
        )
        try:
            rect = wintypes.RECT()
            rect.left = LVIR_BOUNDS
            written = ctypes.c_size_t()
            kernel32.WriteProcessMemory(process, remote_rect, ctypes.byref(rect), ctypes.sizeof(rect), ctypes.byref(written))
            ok = user32.SendMessageW(hwnd, LVM_GETITEMRECT, index, remote_rect)
            if not ok:
                raise EvidenceValidationError(f"Could not read list view row rectangle for token {token!r}")
            read = ctypes.c_size_t()
            kernel32.ReadProcessMemory(process, remote_rect, ctypes.byref(rect), ctypes.sizeof(rect), ctypes.byref(read))
        finally:
            if remote_rect:
                kernel32.VirtualFreeEx(process, remote_rect, 0, MEM_RELEASE)
            kernel32.CloseHandle(process)

        point = wintypes.POINT()
        point.x = int(rect.left + max(4, (rect.right - rect.left) // 2))
        point.y = int(rect.top + max(4, (rect.bottom - rect.top) // 2))
        user32.ClientToScreen(hwnd, ctypes.byref(point))
        self.assert_foreground_native_gui(None, f"list_{token}")
        root = self.assert_point_in_window(point.x, point.y, hwnd=hwnd, phase=f"list_{token}")
        self.activate_hwnd(root)
        pyautogui.click(point.x, point.y)
        self.wait(0.15)
        if double or expand:
            pyautogui.press("enter")
        self.log(f"win32-list-click token={token!r} text={item.get('text')!r} hwnd={hwnd} index={index} at={point.x},{point.y}")
        self.wait(delay)

    def click_text(
        self,
        item: dict[str, Any],
        phase: str,
        extra: Iterable[str] = (),
        double: bool = False,
        expand: bool = False,
        delay: float = 0.8,
        include_item_keywords: bool = False,
        scroll_attempts: int = 0,
        scroll_key: str = "wheel",
    ) -> None:
        row_keywords = self.item_keywords(item)
        keywords = []
        candidates = list(extra)
        if include_item_keywords:
            candidates += row_keywords
        for token in candidates:
            token = str(token).strip()
            if token and token.casefold() not in {seen.casefold() for seen in keywords}:
                keywords.append(token)
        found = None
        for attempt in range(scroll_attempts + 1):
            found = self.find_text_element(keywords)
            self.log(f"text-click phase={phase} attempt={attempt} row_keywords={row_keywords} search_keywords={keywords} found={found}")
            if found:
                break
            list_found = self.find_listview_item(keywords)
            if list_found:
                list_item, token = list_found
                self.log(f"win32-list-found phase={phase} attempt={attempt} row_keywords={row_keywords} search_keywords={keywords} found={list_item}")
                self.click_listview_item(list_item, token, double=double, expand=expand, delay=delay)
                if double or expand:
                    self.verify_navigation_after_click(item, phase, extra)
                return
            if attempt < scroll_attempts:
                if scroll_key == "wheel":
                    pyautogui.scroll(-5)
                else:
                    pyautogui.press(scroll_key)
                self.wait(0.35)
        if not found:
            diag = self.tmp_shot(f"row{int(item['row']):02d}_{phase}_text_not_found.png")
            raise EvidenceValidationError(
                f"Could not find visible UI text before click at {phase}. "
                f"Keywords extracted from row: {row_keywords}; search_keywords={keywords}; diagnostic={Path(diag).name}"
            )
        element, token = found
        x = int(element["x"] + max(2, element["w"] // 2))
        y = int(element["y"] + max(2, element["h"] // 2))
        self.assert_foreground_native_gui(item, phase)
        hwnd = self.assert_point_in_window(x, y, phase=phase)
        self.activate_hwnd(hwnd)
        pyautogui.click(x, y)
        self.wait(0.15)
        if expand:
            pyautogui.press("right")
        elif double:
            pyautogui.press("enter")
        self.log(f"text-clicked phase={phase} token={token!r} at={x},{y}")
        self.wait(delay)
        if double or expand:
            self.verify_navigation_after_click(item, phase, extra)

    def set_english_input(self, phase: str) -> None:
        """Ask the foreground window to use the US keyboard layout before ASCII list search."""
        try:
            user32 = ctypes.windll.user32
            hkl = user32.LoadKeyboardLayoutW("00000409", 1)
            hwnd = user32.GetForegroundWindow()
            user32.PostMessageW(hwnd, 0x0050, 0, hkl)  # WM_INPUTLANGCHANGEREQUEST
            user32.ActivateKeyboardLayout(hkl, 0)
            self.wait(0.2)
            pyautogui.press("esc")
            self.wait(0.1)
            self.log(f"english-input phase={phase} hwnd={hwnd} hkl={hkl}")
        except Exception as exc:
            self.log(f"english-input failed phase={phase}: {exc!r}")

    def focus_details_pane(self, phase: str, tabs: int = 1) -> None:
        for _ in range(max(0, tabs)):
            pyautogui.press("tab")
            self.wait(0.15)
        self.log(f"focus-details phase={phase} tabs={tabs}")

    def keyboard_search_list_item(
        self,
        item: dict[str, Any],
        phase: str,
        text: str,
        *,
        enter: bool = False,
        tabs: int = 1,
        delay: float = 0.8,
    ) -> None:
        row_keywords = self.item_keywords(item)
        if not text or text.casefold() not in [token.casefold() for token in row_keywords]:
            self.log(f"keyboard-list phase={phase} text={text!r} row_keywords={row_keywords}")
        self.focus_details_pane(phase, tabs=tabs)
        self.set_english_input(phase)
        pyautogui.press("home")
        self.wait(0.2)
        pyautogui.write(text, interval=0.02)
        self.wait(delay)
        if enter:
            pyautogui.press("enter")
            self.wait(delay)
        self.log(f"keyboard-list phase={phase} text={text!r} enter={enter} row_keywords={row_keywords}")

    def keyboard_select_visible(self, phase: str, text: str, *, enter: bool = False, delay: float = 0.8) -> None:
        if not text:
            raise ValueError("text is required")
        self.set_english_input(phase)
        pyautogui.write(text, interval=0.02)
        self.wait(delay)
        if enter:
            pyautogui.press("enter")
            self.wait(delay)
        self.log(f"keyboard-select phase={phase} text={text!r} enter={enter}")

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
        px = int(win.left + x)
        py = int(win.top + y)
        hwnd = int(getattr(win, "_hWnd", 0) or getattr(win, "hWnd", 0) or 0) or None
        root = self.assert_point_in_window(px, py, hwnd=hwnd, phase="coordinate_click")
        self.activate_hwnd(root)
        pyautogui.click(px, py)
        self.wait(delay)

    def double_click(self, win: Any, x: int, y: int, delay: float = 0.8) -> None:
        px = int(win.left + x)
        py = int(win.top + y)
        hwnd = int(getattr(win, "_hWnd", 0) or getattr(win, "hWnd", 0) or 0) or None
        root = self.assert_point_in_window(px, py, hwnd=hwnd, phase="coordinate_double_click")
        self.activate_hwnd(root)
        pyautogui.doubleClick(px, py)
        self.wait(delay)

    def image_stats(self, image: Any) -> dict[str, Any]:
        sample = image.convert("L").resize((64, 64))
        pixels = list(sample.getdata())
        extrema = sample.getextrema()
        mean = statistics.fmean(pixels)
        stdev = statistics.pstdev(pixels)
        usable = (extrema[1] - extrema[0]) >= 12 and stdev >= 4.0 and not (mean <= 3.0 or mean >= 252.0)
        return {
            "luma_extrema": list(extrema),
            "luma_mean": round(mean, 3),
            "luma_stdev": round(stdev, 3),
            "image_usable": usable,
        }

    def window_snapshot(self) -> dict[str, Any]:
        windows = []
        for win in gw.getAllWindows():
            title = (getattr(win, "title", "") or "").strip()
            if not title:
                continue
            windows.append(
                {
                    "title": title,
                    "left": getattr(win, "left", None),
                    "top": getattr(win, "top", None),
                    "width": getattr(win, "width", None),
                    "height": getattr(win, "height", None),
                }
            )
        try:
            active = gw.getActiveWindow()
            active_title = (getattr(active, "title", "") or "").strip() if active else ""
        except Exception:
            active_title = ""
        return {"active_title": active_title, "windows": windows}

    def semantic_rule(self, filename: str, semantic_hint: str | None = None) -> dict[str, Any] | None:
        key = f"{Path(filename).stem.lower()} {str(semantic_hint or '').lower()}"
        for marker, rule in SEMANTIC_RULES:
            if marker in key:
                return rule
        return None

    def validate_candidate(
        self,
        filename: str,
        image: Any,
        *,
        allow_command_window: bool = False,
        semantic_hint: str | None = None,
    ) -> dict[str, Any]:
        stats = self.image_stats(image)
        snapshot = self.window_snapshot()
        ui_elements = self.uia_elements()
        list_items = self.listview_text_items()
        ui_text = " ".join(
            [snapshot.get("active_title", "")] + [w["title"] for w in snapshot["windows"]]
            + [self.element_text(e) for e in ui_elements]
            + [str(item.get("row_text") or item.get("text") or "") for item in list_items]
        ).casefold()
        rule = self.semantic_rule(filename, semantic_hint)
        failures: list[str] = []
        if not stats["image_usable"]:
            failures.append(f"blank_or_low_information_image stats={stats}")
        active_title = snapshot.get("active_title", "").casefold()
        command_window_hits = [token for token in COMMAND_WINDOW_TITLE_TOKENS if token in active_title]
        if command_window_hits and not allow_command_window:
            failures.append(
                "non_native_gui_command_window_active="
                f"{snapshot.get('active_title', '')!r}; ask_user_before_using_command_window_evidence"
            )
        if rule:
            must_any = [token.casefold() for token in rule.get("must_any", [])]
            must_all_any = [[token.casefold() for token in group] for group in rule.get("must_all_any", [])]
            forbid_any = [token.casefold() for token in rule.get("forbid_any", [])]
            if must_any and not any(token in ui_text for token in must_any):
                failures.append(f"missing_expected_window_tokens={rule.get('must_any', [])}")
            for group in must_all_any:
                if group and not any(token in ui_text for token in group):
                    failures.append(f"missing_expected_window_token_group={group}")
            forbidden_hits = [token for token in forbid_any if token in ui_text]
            if forbidden_hits:
                failures.append(f"forbidden_window_tokens={forbidden_hits}")
        return {
            "accepted": not failures,
            "failures": failures,
            "rule": rule or {},
            "image": stats,
            "active_title": snapshot["active_title"],
            "window_titles": [w["title"] for w in snapshot["windows"]],
            "command_window_evidence_allowed": allow_command_window,
            "semantic_hint": semantic_hint or "",
            "matched_ui_text_sample": ui_text[:500],
        }

    def shot(
        self,
        filename: str,
        *,
        final: bool = True,
        allow_command_window: bool = False,
        semantic_hint: str | None = None,
    ) -> str:
        root = self.out_dir if final else self.tmp_dir
        target = root / Path(filename).name
        target.parent.mkdir(parents=True, exist_ok=True)
        self.wait(0.7)
        if final:
            self.assert_foreground_native_gui(None, f"before_screenshot_{Path(filename).stem}", allow_command_window=allow_command_window)
        image = pyautogui.screenshot()
        if final:
            candidate = self.tmp_dir / f"candidate_{target.name}"
            image.save(candidate)
            validation = self.validate_candidate(
                target.name,
                image,
                allow_command_window=allow_command_window,
                semantic_hint=semantic_hint,
            )
            validation_path = self.tmp_dir / f"{target.stem}.validation.json"
            validation_path.write_text(json.dumps(validation, ensure_ascii=False, indent=2), encoding="utf-8")
            self.log(f"validation {target.name} {json.dumps(validation, ensure_ascii=False)}")
            if not validation["accepted"]:
                raise EvidenceValidationError(f"Screenshot validation failed for {target.name}: {'; '.join(validation['failures'])}")
        image.save(target)
        kind = "shot" if final else "tmp-shot"
        self.log(f"{kind} {target}")
        return str(target)

    def tmp_shot(self, filename: str) -> str:
        return self.shot(filename, final=False)

    def evidence_shot(
        self,
        item: dict[str, Any],
        filename: str,
        *,
        semantic_hint: str | None = None,
        allow_command_window: bool = False,
    ) -> str:
        hint_parts = [
            semantic_hint or "",
            str(item.get("gui_action") or ""),
            str(item.get("check_id") or ""),
            str(item.get("evidence_slug") or ""),
        ]
        hint = " ".join(part for part in hint_parts if part)
        return self.shot(filename, allow_command_window=allow_command_window, semantic_hint=hint)

    def delivery_result_for_item(self, item: dict[str, Any]) -> dict[str, str]:
        action = str(item.get("gui_action") or item.get("check_id") or "")
        text = self.visible_text()
        folded = text.casefold()
        if action == "gpedit_rdp_client_connection_encryption_level":
            if "未配置" in text or "not configured" in folded:
                return {"finding": "远程桌面客户端连接加密级别策略为未配置。", "result": "不符合"}
            if "已启用" in text or "enabled" in folded:
                return {"finding": "远程桌面客户端连接加密级别策略已启用。", "result": "符合"}
        if action == "services_remote_registry":
            return {"finding": "Remote Registry 服务状态已在服务管理器中显示。"}
        if action == "control_installed_updates":
            kbs = sorted(set(re.findall(r"\bKB\d{6,8}\b", text, flags=re.I)))
            if kbs:
                return {"finding": f"已安装更新页面可见补丁：{', '.join(kbs[:8])}。"}
            return {"finding": "已安装更新页面已打开，未从可见区域识别到 KB 编号。"}
        return {}

    def record_delivery_result(self, item: dict[str, Any]) -> None:
        result = self.delivery_result_for_item(item)
        if result:
            self.delivery_results[int(item["row"])] = result

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

    def command_window_evidence_allowed(self, item: dict[str, Any]) -> bool:
        value = item.get("allow_command_window_evidence") or item.get("explicit_command_window_evidence")
        return str(value).casefold() in {"1", "true", "yes", "y"}

    def visible_cmd(self, item: dict[str, Any], name: str, commands: list[str], filename: str) -> list[str]:
        row = int(item["row"])
        if not self.command_window_evidence_allowed(item):
            raise UserConfirmationRequiredError(
                f"Row {row} would require command-window evidence ({name}), but Windows screenshot evidence must be a native GUI page. "
                "Stop and ask the user whether to authorize command-window evidence or provide a manual/native-GUI path."
            )
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
        return [self.evidence_shot(item, filename, semantic_hint=name, allow_command_window=True)]

    def open_secpol_account_child(self, child: str, item: dict[str, Any]) -> Any:
        win = self.open_mmc("secpol.msc", ["本地安全策略", "Local Security Policy"])
        self.click_text(item, "secpol_account_policies", ["帐户策略", "账户策略", "Account Policies"], expand=True, delay=0.6)
        if child == "password":
            self.click_text(item, "secpol_password_policy", ["密码策略", "Password Policy"], double=True, delay=1.2)
        elif child == "lockout":
            self.click_text(item, "secpol_account_lockout_policy", ["帐户锁定策略", "账户锁定策略", "Account Lockout Policy"], double=True, delay=1.2)
        else:
            raise ValueError(child)
        return self.active_window(["本地安全策略", "Local Security Policy"])

    def open_secpol_local_child(self, child: str, item: dict[str, Any]) -> Any:
        win = self.open_mmc("secpol.msc", ["本地安全策略", "Local Security Policy"])
        self.click_text(item, "secpol_local_policies", ["本地策略", "Local Policies"], expand=True, delay=0.6)
        if child == "audit":
            self.click_text(item, "secpol_audit_policy", ["审核策略", "Audit Policy"], double=True, delay=1.2)
        elif child == "security":
            self.click_text(item, "secpol_security_options", ["安全选项", "Security Options"], double=True, delay=1.2)
        elif child == "rights":
            self.click_text(item, "secpol_user_rights", ["用户权限分配", "User Rights Assignment"], double=True, delay=1.2)
        else:
            raise ValueError(child)
        return self.active_window(["本地安全策略", "Local Security Policy"])

    def open_gpedit_computer_admin_templates(self, item: dict[str, Any], debug_prefix: str | None = None) -> Any:
        win = self.open_mmc("gpedit.msc", ["本地组策略编辑器", "Local Group Policy Editor"], wait=5)
        self.click_text(item, "gpedit_computer_configuration", ["计算机配置", "Computer Configuration"], double=True, delay=0.8)
        self.click_list_text(item, "gpedit_administrative_templates", ["管理模板", "Administrative Templates"], double=True, delay=0.8)
        if debug_prefix:
            self.tmp_shot(f"{debug_prefix}_admin_templates.png")
        return self.active_window(["本地组策略编辑器", "Local Group Policy Editor"])

    def open_gpedit_computer_all_settings(self, item: dict[str, Any]) -> Any:
        self.open_gpedit_computer_admin_templates(item)
        self.click_list_text(item, "gpedit_all_settings", ["所有设置", "All Settings"], delay=1.0)
        return self.active_window(["本地组策略编辑器", "Local Group Policy Editor"])

    def open_gpedit_computer_rdsh(self, folder_name: str, item: dict[str, Any]) -> Any:
        self.open_gpedit_computer_admin_templates(item)
        self.click_list_text(item, "gpedit_windows_components", ["Windows 组件", "Windows Components"], double=True, delay=0.8)
        self.focus_details_pane("gpedit_windows_components_details")
        self.click_list_text(
            item,
            "gpedit_remote_desktop_services",
            ["远程桌面服务", "Remote Desktop Services", "终端服务"],
            double=True,
            delay=0.8,
        )
        self.focus_details_pane("gpedit_rd_services_details")
        self.click_list_text(
            item,
            "gpedit_remote_desktop_session_host",
            ["远程桌面会话主机", "Remote Desktop Session Host"],
            double=True,
            delay=0.8,
        )
        folder_tokens = {
            "security": ["安全", "Security", "加密", "encryption"],
            "session_time_limits": ["会话时间限制", "Session Time Limits", "空闲", "idle"],
            "connections": ["连接", "Connections", "连接数量"],
        }
        self.focus_details_pane(f"gpedit_{folder_name}_details")
        self.click_list_text(
            item,
            f"gpedit_{folder_name}",
            folder_tokens[folder_name],
            double=True,
            delay=0.8,
        )
        return self.active_window(["本地组策略编辑器", "Local Group Policy Editor"])

    def adaptive_open_and_capture(self, item: dict[str, Any]) -> list[str]:
        tool = str(item.get("tool") or "control").split(";")[0]
        if tool in {"cmd", "powershell"}:
            return self.visible_cmd(item, item.get("check_id") or "adaptive", ["whoami"], item["evidence"][0])
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
        return [self.evidence_shot(item, item["evidence"][0], semantic_hint=str(item.get("gui_action") or item.get("check_id") or ""))]


def first_evidence(item: dict[str, Any], default: str) -> str:
    evidence = item.get("evidence") or [default]
    return Path(evidence[0]).name


def all_evidence(item: dict[str, Any]) -> list[str]:
    return [Path(name).name for name in item.get("evidence", [])]


def action_identity_users(r: Runner, item: dict[str, Any]) -> list[str]:
    return r.visible_cmd(
        item,
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
    r.open_secpol_account_child("password", item)
    path = r.evidence_shot(item, first_evidence(item, f"row{item['row']:02d}_secpol_password_policy.png"), semantic_hint="secpol_password_policy")
    r.kill_ui()
    return [path]


def action_secpol_account_lockout_policy(r: Runner, item: dict[str, Any]) -> list[str]:
    r.open_secpol_account_child("lockout", item)
    path = r.evidence_shot(item, first_evidence(item, f"row{item['row']:02d}_secpol_account_lockout_policy.png"), semantic_hint="secpol_account_lockout_policy")
    r.kill_ui()
    return [path]


def action_gpedit_rdp_client_connection_encryption_level(r: Runner, item: dict[str, Any]) -> list[str]:
    try:
        r.open_gpedit_computer_rdsh("security", item)
    except EvidenceValidationError as exc:
        r.log(f"gpedit security path failed; falling back to all settings: {exc}")
        r.open_gpedit_computer_all_settings(item)
    r.click_text(item, "gpedit_rdp_encryption_policy", ["设置客户端连接加密级别", "Set client connection encryption level", "encryption level", "加密级别", "要求使用特定的安全层"], double=True, delay=1.0, scroll_attempts=8)
    path = r.evidence_shot(item, first_evidence(item, f"row{item['row']:02d}_gpedit_rdp_client_connection_encryption_level.png"), semantic_hint="gpedit_rdp_client_connection_encryption_level")
    r.record_delivery_result(item)
    r.kill_ui()
    return [path]


def action_lusrmgr_users(r: Runner, item: dict[str, Any]) -> list[str]:
    win = r.open_mmc("lusrmgr.msc", ["本地用户和组", "Local Users and Groups"])
    r.click_text(item, "lusrmgr_users_node", ["用户", "Users"], double=True, delay=1.0)
    path = r.evidence_shot(item, first_evidence(item, f"row{item['row']:02d}_lusrmgr_users.png"), semantic_hint="lusrmgr_users")
    r.kill_ui()
    return [path]


def action_admin_auth_method(r: Runner, item: dict[str, Any]) -> list[str]:
    return r.visible_cmd(
        item,
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
    r.click_text(item, "fsmgmt_shares_node", ["共享", "Shares"], double=True, delay=1.0)
    path = r.evidence_shot(item, first_evidence(item, f"row{item['row']:02d}_fsmgmt_shares.png"), semantic_hint="fsmgmt")
    r.kill_ui()
    return [path]


def action_regedit_lsa_restrictanonymous(r: Runner, item: dict[str, Any]) -> list[str]:
    win = r.open_regedit_key(r"计算机\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa")
    r.set_listview_column_widths("regedit_lsa_restrictanonymous_columns", [260, 110, 260])
    r.require_visible_token_groups(
        item,
        "regedit_lsa_restrictanonymous_visible",
        [["restrictanonymous"], ["restrictanonymoussam"], ["REG_DWORD", "DWORD"]],
    )
    r.wait(0.5)
    path = r.evidence_shot(item, first_evidence(item, f"row{item['row']:02d}_regedit_lsa_restrictanonymous.png"), semantic_hint="regedit_lsa_restrictanonymous")
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
    r.click_text(item, "lusrmgr_groups_node", ["组", "Groups"], double=True, delay=0.6)
    r.keyboard_search_list_item(item, "lusrmgr_administrators_group", "Administrators", enter=True, tabs=1, delay=1.0)
    captured.append(r.evidence_shot(item, names[0], semantic_hint="lusrmgr_administrators_members"))
    r.kill_ui()

    if len(names) > 1:
        win = r.open_mmc("services.msc", ["服务", "Services"])
        pyautogui.press("home")
        r.wait(0.2)
        pyautogui.write("MySQLa", interval=0.03)
        r.wait(0.8)
        r.require_visible_keywords(item, "services_mysql_before_enter", ["MySQL", "服务", "Services"], min_hits=1)
        pyautogui.press("enter")
        r.wait(1.0)
        pyautogui.hotkey("ctrl", "tab")
        r.wait(0.5)
        captured.append(r.evidence_shot(item, names[1], semantic_hint="services"))
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
    r.click_text(item, "lusrmgr_users_node", ["用户", "Users"], double=True, delay=1.0)
    r.set_listview_column_widths("lusrmgr_default_accounts_columns", [240, 160, 320])
    row_text = " ".join(r.item_keywords(item)).casefold()
    required_accounts = [["Administrator"], ["Guest"]]
    if "defaultaccount" in row_text:
        required_accounts.append(["DefaultAccount"])
    if "wdagutilityaccount" in row_text:
        required_accounts.append(["WDAGUtilityAccount"])
    r.require_visible_token_groups(
        item,
        "lusrmgr_default_accounts_visible",
        required_accounts,
    )
    captured.append(r.evidence_shot(item, names[0], semantic_hint="lusrmgr_default_accounts"))
    if len(names) > 1:
        r.keyboard_search_list_item(item, "lusrmgr_guest_account", "Guest", enter=True, tabs=1, delay=1.0)
        r.require_visible_token_groups(
            item,
            "lusrmgr_guest_properties_disabled_visible",
            [["Guest"], ["属性", "Properties"], ["账户已禁用", "帐户已禁用", "Account is disabled"]],
        )
        captured.append(r.evidence_shot(item, names[1], semantic_hint="lusrmgr_guest_properties_disabled"))
    r.kill_ui()
    return captured


def action_lusrmgr_extra_accounts(r: Runner, item: dict[str, Any]) -> list[str]:
    return action_lusrmgr_users(r, item)


def action_secpol_audit_policy(r: Runner, item: dict[str, Any]) -> list[str]:
    r.open_secpol_local_child("audit", item)
    path = r.evidence_shot(item, first_evidence(item, f"row{item['row']:02d}_secpol_audit_policy.png"), semantic_hint="secpol_audit_policy")
    r.kill_ui()
    return [path]


def action_eventvwr_system_log(r: Runner, item: dict[str, Any]) -> list[str]:
    r.kill_ui()
    subprocess.Popen([r"C:\Windows\System32\mmc.exe", "eventvwr.msc", "/c:System"])
    r.wait(6.0)
    r.maximize(r.wait_window(["事件查看器", "Event Viewer"], 25))
    path = r.evidence_shot(item, first_evidence(item, f"row{item['row']:02d}_eventvwr_system_log.png"), semantic_hint="eventvwr")
    r.kill_ui()
    return [path]


def action_eventvwr_system_log_properties(r: Runner, item: dict[str, Any]) -> list[str]:
    r.kill_ui()
    subprocess.Popen([r"C:\Windows\System32\mmc.exe", "eventvwr.msc", "/c:System"])
    r.wait(6.0)
    win = r.maximize(r.wait_window(["事件查看器", "Event Viewer"], 25))
    r.click_text(item, "eventvwr_system_properties", ["属性", "Properties"], delay=1.0)
    path = r.evidence_shot(item, first_evidence(item, f"row{item['row']:02d}_eventvwr_system_log_properties.png"), semantic_hint="eventvwr_system_log_properties")
    r.kill_ui()
    return [path]


def action_control_installed_updates(r: Runner, item: dict[str, Any]) -> list[str]:
    r.kill_ui()
    tokens = ["已安装更新", "Installed Updates"]
    titles = ["已安装更新", "Installed Updates", "程序和功能", "Programs and Features"]

    # Do not use appwiz.cpl,,2 here: on Server 2012 it can route to the
    # Windows Features page and trigger a missing OptionalFeatures.exe dialog.
    subprocess.Popen([r"C:\Windows\System32\control.exe", "appwiz.cpl"])
    r.wait(5.0)
    win = r.maximize(r.wait_window(titles, 25))

    if not r.control_panel_location_has(win, tokens):
        # Server 2012 often opens Programs and Features first; use the left
        # navigation link, but require visible UI text before clicking.
        r.click_text(item, "control_installed_updates_link", ["查看已安装的更新", "View installed updates", "已安装更新", "Installed Updates"], delay=4.0)
        win = r.active_window(titles)

    if not r.control_panel_location_has(win, tokens):
        raise RuntimeError("Programs and Features did not navigate to Installed Updates")
    r.require_visible_token_groups(
        item,
        "control_installed_updates_visible",
        [
            ["已安装更新", "已安装的更新", "Installed Updates"],
            ["KB", "Microsoft Windows", "Update for Microsoft Windows", "Security Update", "更新"],
        ],
    )

    path = r.evidence_shot(item, first_evidence(item, f"row{item['row']:02d}_control_installed_updates.png"), semantic_hint="control_installed_updates")
    r.record_delivery_result(item)
    r.kill_ui()
    return [path]


def action_services_remote_registry(r: Runner, item: dict[str, Any]) -> list[str]:
    win = r.open_mmc("services.msc", ["服务", "Services"])
    r.click_text(item, "services_remote_registry", ["Remote Registry"], delay=1.0)
    path = r.evidence_shot(item, first_evidence(item, f"row{item['row']:02d}_services_remote_registry.png"), semantic_hint="services_remote_registry")
    r.record_delivery_result(item)
    r.kill_ui()
    return [path]


def action_gpedit_idle_session_limit(r: Runner, item: dict[str, Any]) -> list[str]:
    r.open_gpedit_computer_rdsh("session_time_limits", item)
    r.click_text(item, "gpedit_idle_session_policy", ["设置活动但空闲的远程桌面服务会话的时间限制", "活动但空闲", "空闲", "Set time limit for active but idle", "idle session"], double=True, delay=1.0, scroll_attempts=8)
    path = r.evidence_shot(item, first_evidence(item, f"row{item['row']:02d}_gpedit_idle_session_limit.png"), semantic_hint="gpedit_idle_session_limit")
    r.kill_ui()
    return [path]


def action_gpedit_limit_number_of_connections(r: Runner, item: dict[str, Any]) -> list[str]:
    try:
        r.open_gpedit_computer_rdsh("connections", item)
        r.click_text(item, "gpedit_limit_connections_policy", ["限制连接的数量", "Limit number of connections", "MaxInstanceCount"], double=True, delay=1.0, scroll_attempts=8)
    except EvidenceValidationError as exc:
        r.log(f"gpedit connections path failed; falling back to all settings: {exc}")
        r.open_gpedit_computer_all_settings(item)
        r.click_text(item, "gpedit_limit_connections_policy_all_settings", ["限制连接的数量", "Limit number of connections", "MaxInstanceCount"], double=True, delay=1.0, scroll_attempts=8)
    path = r.evidence_shot(item, first_evidence(item, f"row{item['row']:02d}_gpedit_limit_number_of_connections.png"), semantic_hint="gpedit_limit_number_of_connections")
    r.kill_ui()
    return [path]


def action_secpol_security_options(r: Runner, item: dict[str, Any]) -> list[str]:
    r.open_secpol_local_child("security", item)
    path = r.evidence_shot(item, first_evidence(item, f"row{item['row']:02d}_secpol_security_options.png"), semantic_hint="secpol_security_options")
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
        if action_name == "skip_admin_interview" or item.get("check_id") == "admin_interview":
            runner.log(f"skip row={item['row']} reason=administrator_interview")
            results.append(
                {
                    "row": item["row"],
                    "status": "skipped",
                    "skip_reason": "administrator_interview",
                    "action": action_name,
                    "lane": item.get("lane"),
                    "finding": item.get("finding") or ADMIN_INTERVIEW_FINDING,
                    "result": item.get("result") or ADMIN_INTERVIEW_RESULT,
                    "evidence": [],
                }
            )
            continue
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
                    **runner.delivery_results.get(int(item["row"]), {}),
                }
            )
            runner.log(f"done row={item['row']} evidence={evidence_paths}")
        except Exception as exc:
            runner.log(f"ERROR row={item.get('row')} action={action_name}: {type(exc).__name__}: {exc}")
            if isinstance(exc, UserConfirmationRequiredError):
                status = "needs_user_confirmation"
            elif isinstance(exc, EvidenceValidationError):
                status = "validation_failed"
            else:
                status = "error"
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
                    "status": status,
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
