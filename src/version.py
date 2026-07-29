"""ProAV Shoko version information.

Stage: "alpha", "beta", "rc", or "" for stable.
Prerelease auto-increments on each build via scripts/build.py.
"""

MAJOR = 0
MINOR = 4
PATCH = 0
STAGE = "beta"
PRERELEASE = 11


def get_version():
    base = f"{MAJOR}.{MINOR}.{PATCH}"
    if STAGE:
        return f"{base}-{STAGE}.{PRERELEASE}"
    return base


def get_exe_name():
    return f"ProAV Shoko {get_version()}"
