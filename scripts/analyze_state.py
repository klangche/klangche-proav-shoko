#!/usr/bin/env python3
"""
Simple script to analyze the current state and fix the GUI.
"""

import os
import re

print("=" * 70)
print("ANALYZING PROAV SHOKO - FIX GUI AND ROOT DIRECTORY")
print("=" * 70)

# 1. Analyze current gui.py
print("\n1. GUI ANALYSIS:")
if os.path.exists('src/gui.py'):
    with open('src/gui.py', 'r') as f:
        gui_content = f.read()
    
    # Check key features
    has_update_display = 'def _update_tree_display' in gui_content
    has_verdict = 'VERDICT' in gui_content
    has_full_tree = 'Full USB & Display Tree' in gui_content
    has_tagg = 'Tagg' in gui_content
    
    print(f"   Has _update_tree_display: {has_update_display}")
    print(f"   Has VERDICT sections: {has_verdict}")
    print(f"   Has 'Full USB & Display Tree': {has_full_tree}")
    print(f"   Has Tagg patterns: {has_tagg}")
    
    # Find Tagg patterns
    tagg_matches = re.findall(r'Tagg \w+\.\w+', gui_content)
    print(f"   Found Tagg patterns: {len(tagg_matches)}")
    if tagg_matches:
        for m in tagg_matches[:3]:
            print(f"     - {m}")
else:
    print("   ✗ src/gui.py NOT FOUND")

# 2. Analyze current directory structure
print("\n2. ROOT DIRECTORY ANALYSIS:")
items = os.listdir('.')
py_files = [f for f in items if f.endswith('.py') and os.path.isfile(f)]

print(f"   Python files: {len(py_files)}")
for f in sorted(py_files):
    size = os.path.getsize(f)
    print(f"     - {f:40} ({size:>8} bytes)")

# 3. Check scripts/ directory
print("\n3. SCRIPTS/ DIRECTORY:")
if os.path.exists('scripts'):
    script_items = os.listdir('scripts')
    print(f"   Total items: {len(script_items)}")
    
    # Count by type
    py_files = [f for f in script_items if f.endswith('.py')]
    dirs = [f for f in script_items if os.path.isdir(os.path.join('scripts', f))]
    
    print(f"   Python files: {len(py_files)}")
    print(f"   Directories: {len(dirs)}")
    
    print("\n   Files:")
    for f in sorted(py_files):
        size = os.path.getsize(os.path.join('scripts', f))
        print(f"     - {f:30} ({size:>8} bytes)")
else:
    print("   ✗ scripts/ directory NOT FOUND")

# 4. Check requirements from CLI output
print("\n4. TAGGING REQUIREMENTS (from CLI output):")
print("   CLI shows these TAGGING lines:")
print("     - Tagg xxx.tree")
print("     - Tagg xxx.score")
print("     - Tag xxx.warnings")

# 5. Analysis
print("\n" + "=" * 70)
print("ANALYSIS SUMMARY:")
print("=" * 70)

issues = []

# Issue 1: TAGGING support
if os.path.exists('src/gui.py'):
    gui_content = open('src/gui.py').read()
    if 'Tagg' not in gui_content:
        issues.append("GUI MISSING TAGGING support (needs Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings)")
    else:
        # Check if it has the right patterns
        tagg_matches = re.findall(r'Tagg \w+\.\w+', gui_content)
        if len(tagg_matches) < 2:  # Should have at least Tagg xxx.tree and Tagg xxx.score
            issues.append(f"GUI Tagg support incomplete - only {len(tagg_matches)} patterns found")
        else:
            print(f"   ✓ GUI has Tagg support ({len(tagg_matches)} patterns)")
else:
    issues.append("GUI (src/gui.py) NOT FOUND")

# Issue 2: Root directory cleanup
if len(py_files) > 10:
    issues.append("ROOT DIRECTORY has too many files (should move build/config files to scripts/)")
else:
    print(f"   ✓ Root directory has reasonable number of files ({len(py_files)})")

# Issue 3: Scripts directory
if os.path.exists('scripts'):
    script_py_files = len([f for f in os.listdir('scripts') if f.endswith('.py')])
    print(f"   ✓ scripts/ directory exists with {script_py_files} Python files")
else:
    issues.append("scripts/ directory NOT FOUND")

print(f"\nTotal issues found: {len(issues)}")
for i, issue in enumerate(issues, 1):
    print(f"  {i}. {issue}")

print("\n" + "=" * 70)
print("RECOMMENDED ACTIONS:")
print("=" * 70)

if issues:
    print("1. FIX GUI TAGGING SUPPORT:")
    print("   - Update src/gui.py _update_tree_display method")
    print("   - Add Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings sections")
    print("   - Ensure it matches CLI output format")
    
    print("\n2. CLEAN UP ROOT DIRECTORY:")
    print("   - Move build/config files to scripts/ directory")
    print("   - Keep only essential files: run.py, pyproject.toml, README.md")
    
    print("\n3. ADD PROGRESS BAR/TIMER:")
    print("   - Install tqdm: pip install tqdm")
    print("   - Add progress indicator to USB scanning")
    print("   - Add progress display in GUI")
else:
    print("✓ All requirements appear to be met")

print("\n" + "=" * 70)
print("NEXT STEPS:")
print("=" * 70)
print("1. Check if analysis_method.py has been run")
print("2. Review scripts/analyze_method.py for specific fixes needed")
print("3. Implement Tagg/Tagging support in GUI")
print("4. Clean up root directory structure")
