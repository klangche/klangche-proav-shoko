"""
Display analyzer - handles information about connected displays
"""

import importlib
import math
from typing import Any, Dict, List


class DisplayAnalyzer:
    """Analyzes connected displays."""

    def _is_internal_display(self, monitor_name: str, width_mm: int, height_mm: int) -> bool:
        """Heuristic to determine if a display is internal (laptop panel)."""
        # Method 1: Check EDID manufacturer via EnumDisplayDevices
        try:
            import ctypes
            from ctypes import wintypes

            user32 = ctypes.windll.user32

            class DISPLAY_DEVICE(ctypes.Structure):
                _fields_ = [
                    ('cb', wintypes.DWORD),
                    ('DeviceName', wintypes.WCHAR * 32),
                    ('DeviceString', wintypes.WCHAR * 128),
                    ('StateFlags', wintypes.DWORD),
                    ('DeviceID', wintypes.WCHAR * 128),
                    ('DeviceKey', wintypes.WCHAR * 128),
                ]

            i = 0
            while True:
                dd = DISPLAY_DEVICE()
                dd.cb = ctypes.sizeof(dd)
                if not user32.EnumDisplayDevicesW(None, i, ctypes.byref(dd), 0):
                    break

                if dd.DeviceName == monitor_name:
                    j = 0
                    while True:
                        dm = DISPLAY_DEVICE()
                        dm.cb = ctypes.sizeof(dm)
                        if not user32.EnumDisplayDevicesW(dd.DeviceName, j, ctypes.byref(dm), 0):
                            break
                        device_id = dm.DeviceID.lower()
                        laptop_mfrs = ['auo', 'lgd', 'boe', 'cso', 'cmn', 'ivo', 'sdc', 'chi', 'ktf', 'inn']
                        if any(mfr in device_id for mfr in laptop_mfrs):
                            return True
                        j += 1
                i += 1
        except Exception:
            pass

        # Method 2: Physical size heuristic (laptop panels are usually < 18")
        if width_mm and height_mm:
            diag_inches = math.sqrt(width_mm ** 2 + height_mm ** 2) / 25.4
            if diag_inches < 18:
                return True

        return False

    def get_display_info(self) -> List[Dict[str, Any]]:
        """
        Gets information about all connected displays.

        Returns:
            List of display information.
        """
        displays = []
        try:
            screeninfo = importlib.import_module("screeninfo")
            monitors = screeninfo.get_monitors()
            for i, monitor in enumerate(monitors):
                is_int = self._is_internal_display(monitor.name or '', monitor.width_mm, monitor.height_mm)
                display_info = {
                    'index': i,
                    'name': monitor.name or f'Display {i+1}',
                    'width': monitor.width,
                    'height': monitor.height,
                    'resolution': f"{monitor.width}x{monitor.height}",
                    'is_primary': i == 0,
                    'is_internal': is_int,
                    'x': monitor.x,
                    'y': monitor.y
                }
                displays.append(display_info)
        except ModuleNotFoundError:
            print("screeninfo is not installed. Run: pip install screeninfo")
        except Exception as e:
            print(f"Could not read display information: {e}")

        return displays
