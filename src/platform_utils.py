"""
Platform-specific utility functions
"""

import sys
import os
import platform
from typing import Dict, Any


class PlatformUtils:
    """Utilities for platform information."""

    @staticmethod
    def get_platform_info() -> Dict[str, Any]:
        """
        Collects platform information.

        Returns:
            Dictionary with platform information.
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
        """Check whether the program is running with administrator/root privileges."""
        try:
            if sys.platform == 'win32':
                import ctypes
                return ctypes.windll.shell32.IsUserAnAdmin() != 0
            else:
                return os.geteuid() == 0
        except:
            return False
