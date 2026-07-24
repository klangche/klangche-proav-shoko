"""
Plattformsspecifika funktioner
"""

import sys
import os
import platform
from typing import Dict, Any


class PlatformUtils:
    """Verktyg för plattformsinformation."""

    @staticmethod
    def get_platform_info() -> Dict[str, Any]:
        """
        Samlar in plattformsinformation.

        Returns:
            Dictionary med plattformsinformation.
        """
        system = platform.system()
        version = platform.version()
        arch = platform.machine()

        is_apple_silicon = False
        if system == 'Darwin' and arch == 'arm64':
            is_apple_silicon = True

        return {
            'os': system,
            'version': version,
            'architecture': arch,
            'is_apple_silicon': is_apple_silicon,
            'python_version': sys.version
        }

    @staticmethod
    def is_admin() -> bool:
        """Kontrollera om programmet körs med administratörsrättigheter."""
        try:
            if sys.platform == 'win32':
                import ctypes
                return ctypes.windll.shell32.IsUserAnAdmin() != 0
            else:
                return os.geteuid() == 0
        except:
            return False
