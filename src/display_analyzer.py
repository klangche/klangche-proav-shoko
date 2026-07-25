"""
Display analyzer - handles information about connected displays
"""

import importlib
from typing import Any, Dict, List


class DisplayAnalyzer:
    """Analyzes connected displays."""

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
                display_info = {
                    'index': i,
                    'name': monitor.name or f'Display {i+1}',
                    'width': monitor.width,
                    'height': monitor.height,
                    'resolution': f"{monitor.width}x{monitor.height}",
                    'is_primary': i == 0,
                    'x': monitor.x,
                    'y': monitor.y
                }
                displays.append(display_info)
        except ModuleNotFoundError:
            print("[!] screeninfo is not installed. Run: pip install screeninfo")
        except Exception as e:
            print(f"[!] Could not read display information: {e}")

        return displays
