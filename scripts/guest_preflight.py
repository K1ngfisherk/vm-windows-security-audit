#!/usr/bin/env python3
"""Guest-side GUI runtime preflight."""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path


def main() -> int:
    result = {
        "python": sys.executable,
        "cwd": str(Path.cwd()),
        "userprofile": os.environ.get("USERPROFILE"),
        "sessionname": os.environ.get("SESSIONNAME"),
        "pyautogui": False,
        "screenshot": False,
        "screen_size": None,
        "error": None,
    }

    try:
        import pyautogui

        result["pyautogui"] = True
        result["screen_size"] = list(pyautogui.size())
        image = pyautogui.screenshot()
        result["screenshot"] = True
        result["screenshot_size"] = list(image.size)
        # Sample a few pixels to catch fully black screenshots.
        points = [(0, 0), (image.size[0] // 2, image.size[1] // 2), (image.size[0] - 1, image.size[1] - 1)]
        result["sample_pixels"] = [image.getpixel(point) for point in points]
    except Exception as exc:  # pragma: no cover
        result["error"] = repr(exc)

    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["pyautogui"] and result["screenshot"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
