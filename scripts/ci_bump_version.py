"""CI script: bump prerelease number in src/version.py and sync other files.
Prints: version=<full-version>  tag=v<full-version>"""
import json, re
from datetime import date
from pathlib import Path

def read_text(path):
    return path.read_text(encoding="utf-8")

def write_text(path, text):
    path.write_text(text, encoding="utf-8")

version_py = Path("src/version.py")
ns = {}
exec(read_text(version_py), ns)

new_pr = ns["PRERELEASE"] + 1
new_version = f"{ns['MAJOR']}.{ns['MINOR']}.{ns['PATCH']}-{ns['STAGE']}.{new_pr}"

# Update src/version.py
lines = read_text(version_py).splitlines()
for i, line in enumerate(lines):
    if line.startswith("PRERELEASE = "):
        lines[i] = f"PRERELEASE = {new_pr}"
        break
write_text(version_py, "\n".join(lines) + "\n")

# Update proav-shoko.json
json_path = Path("proav-shoko.json")
data = json.loads(read_text(json_path))
data["version"] = new_version
data["lastUpdated"] = date.today().isoformat()
write_text(json_path, json.dumps(data, indent=2, ensure_ascii=False) + "\n")

# Update pyproject.toml -- only [project] version line
toml_path = Path("pyproject.toml")
content = read_text(toml_path)
content = re.sub(r'^version = ".*?"', f'version = "{new_version}"', content, count=1, flags=re.MULTILINE)
write_text(toml_path, content)

print(f"version={new_version}")
print(f"tag=v{new_version}")
