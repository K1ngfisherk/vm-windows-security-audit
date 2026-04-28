#!/usr/bin/env python3
"""Guest-side GUI runtime preflight.

This must run inside the real interactive desktop. A non-interactive VMware
guest operation can import PyAutoGUI but still fail to capture the visible
desktop, or capture a blank service desktop. Treat those as hard failures.
"""

from __future__ import annotations

import json
import os
import statistics
import sys
from pathlib import Path
from typing import Any


def image_statistics(image: Any) -> dict[str, Any]:
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


def main() -> int:
    result: dict[str, Any] = {
        "python": sys.executable,
        "cwd": str(Path.cwd()),
        "username": os.environ.get("USERNAME"),
        "userprofile": os.environ.get("USERPROFILE"),
        "sessionname": os.environ.get("SESSIONNAME"),
        "pyautogui": False,
        "screenshot": False,
        "screen_size": None,
        "image_usable": False,
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
        result.update(image_statistics(image))
    except Exception as exc:  # pragma: no cover
        result["error"] = repr(exc)

    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["pyautogui"] and result["screenshot"] and result["image_usable"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
