#!/usr/bin/env python3
"""Run Linux/Unix audit commands over SSH and write text evidence."""

from __future__ import annotations

import argparse
import json
import logging
import re
import shlex
import socket
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

try:
    import paramiko
except ImportError as exc:  # pragma: no cover
    raise SystemExit("Password SSH requires Python package 'paramiko'. Use key/agent SSH or install paramiko.") from exc

logging.getLogger("paramiko").setLevel(logging.CRITICAL)


def safe_file_part(text: str) -> str:
    if not text.strip():
        return "command"
    value = re.sub(r'[\\/:*?"<>|\s]+', "_", text)
    value = re.sub(r"[^A-Za-z0-9_.-]", "_", value)
    return value.strip("_") or "command"


def load_commands(path: Path) -> list[dict[str, Any]]:
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    if isinstance(data, dict):
        data = data.get("items") or data.get("commands") or [data]
    if not isinstance(data, list):
        raise SystemExit("CommandsJson must be an array or an object containing commands/items")
    return [dict(item) for item in data]


def assert_tcp_port(host: str, port: int, timeout: float = 3.0) -> None:
    with socket.create_connection((host, port), timeout=timeout):
        return


def command_text(item: dict[str, Any], command_id: str) -> str:
    command = str(item.get("command") or "")
    if not command.strip():
        raise SystemExit(f"Command item {command_id!r} has no command text.")
    return command


def connect_ssh(args: argparse.Namespace) -> paramiko.SSHClient:
    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    connect_args: dict[str, Any] = {
        "hostname": args.host_name,
        "port": args.port,
        "username": args.user,
        "timeout": 10,
        "banner_timeout": 10,
        "auth_timeout": 10,
        "look_for_keys": False,
        "allow_agent": False,
    }
    if args.key_file:
        connect_args["key_filename"] = args.key_file
        connect_args["look_for_keys"] = True
        connect_args["allow_agent"] = True
    else:
        connect_args["password"] = args.password
    client.connect(**connect_args)
    return client


def run_commands(args: argparse.Namespace) -> dict[str, Any]:
    commands = load_commands(args.commands_json)
    if args.validate_only:
        assert_tcp_port(args.host_name, args.port)
        return {
            "validated": True,
            "commands": len(commands),
            "transport": "ssh-paramiko",
            "target": f"{args.user}@{args.host_name}:{args.port}",
        }

    assert_tcp_port(args.host_name, args.port)
    label = safe_file_part(args.task_label) or safe_file_part(args.host_name)
    output_dir = args.output_root / f"{label}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    output_dir.mkdir(parents=True, exist_ok=True)

    manifest: list[dict[str, Any]] = []
    client = connect_ssh(args)
    try:
        for index, item in enumerate(commands, start=1):
            raw_id = str(item.get("id") or f"cmd{index:02d}")
            command_id = safe_file_part(raw_id)
            name = str(item.get("name") or command_id)
            command = command_text(item, command_id)
            file_name = f"{index:02d}_{command_id}.txt"
            output_path = output_dir / file_name
            wrapped = "printf '%s\\n' '$ command'; " + command
            remote = "sh -lc " + shlex.quote(wrapped)
            transport = client.get_transport()
            if transport is None:
                raise RuntimeError("SSH transport is not available")
            channel = transport.open_session()
            channel.set_combine_stderr(True)
            channel.exec_command(remote)
            stdin = channel.makefile_stdin("wb", -1)
            stdout = channel.makefile("rb", -1)
            stdin.close()
            out = stdout.read().decode("utf-8", errors="replace")
            exit_code = channel.recv_exit_status()
            output_text = out
            output_path.write_text(output_text, encoding="utf-8")
            manifest.append(
                {
                    "id": command_id,
                    "name": name,
                    "command": command,
                    "transport": "ssh-paramiko",
                    "output": str(output_path),
                    "exitCode": exit_code,
                    "screenshot": None,
                }
            )
    finally:
        client.close()

    manifest_path = output_dir / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"output": str(output_dir), "manifest": str(manifest_path), "commands": len(commands)}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--host-name", required=True)
    parser.add_argument("--user", required=True)
    parser.add_argument("--password", default="")
    parser.add_argument("--key-file", default="")
    parser.add_argument("--port", type=int, default=22)
    parser.add_argument("--commands-json", required=True, type=Path)
    parser.add_argument("--output-root", required=True, type=Path)
    parser.add_argument("--task-label", default="")
    parser.add_argument("--validate-only", action="store_true")
    args = parser.parse_args()

    if not args.password and not args.key_file:
        raise SystemExit("Password SSH requires --password, or use key/agent SSH through run-ssh-commands.ps1.")

    try:
        result = run_commands(args)
    except Exception as exc:
        raise SystemExit(f"SSH command path failed: {type(exc).__name__}: {exc}") from exc
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
