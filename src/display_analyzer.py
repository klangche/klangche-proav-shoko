"""
Skärmanalysator - hanterar information om anslutna skärmar
"""

import sys
from typing import List, Dict, Any

try:
    from screeninfo import get_monitors
except ImportError:
    print("❌ screeninfo är inte installerat. Kör: pip install screeninfo")
    sys.exit(1)


class DisplayAnalyzer:
    """Analyserar anslutna skärmar."""

    def get_display_info(self) -> List[Dict[str, Any]]:
        """
        Hämtar information om alla anslutna skärmar.

        Returns:
            Lista med skärminformation.
        """
        displays = []
        try:
            monitors = get_monitors()
            for i, monitor in enumerate(monitors):
                display_info = {
                    'index': i,
                    'name': monitor.name or f'Skärm {i+1}',
                    'width': monitor.width,
                    'height': monitor.height,
                    'resolution': f"{monitor.width}x{monitor.height}",
                    'is_primary': i == 0,
                    'x': monitor.x,
                    'y': monitor.y
                }
                displays.append(display_info)
        except Exception as e:
            print(f"⚠️  Kunde inte läsa skärminformation: {e}")

        return displays
