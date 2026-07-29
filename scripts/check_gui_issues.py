#!/usr/bin/env python3
"""
Check git state and verify gui.py fixes
"""
import subprocess
import os
import sys

def run_cmd(cmd):
    result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    return result.stdout.strip(), result.stderr.strip(), result.returncode

def main():
    print("=== Git Status ===")
    stdout, stderr, code = run_cmd("git status --porcelain")
    
    if stdout:
        print("Modified files:")
        for line in stdout.split('\n'):
            if line:
                print(f"  {line}")
    else:
        print("No modified files")
    
    # Check if gui.py exists in git
    print("\n=== File Info ===")
    
    # Check git diff for gui.py
    print("\n=== Git diff for src/gui.py ===")
    stdout, stderr, code = run_cmd("git diff src/gui.py")
    
    if stdout:
        lines = stdout.split('\n')
        
        # Count changes
        added = sum(1 for line in lines if line.startswith('+') and not line.startswith('+++'))
        removed = sum(1 for line in lines if line.startswith('-') and not line.startswith('---'))
        
        print(f"Added lines: {added}")
        print(f"Removed lines: {removed}")
        
        # Show context around key changes
        print("\n=== Context around _update_tree_display method ===")
        
        # Find _update_tree_display in the diff
        in_method = False
        method_lines = []
        indent_level = None
        
        for i, line in enumerate(lines):
            if 'def _update_tree_display' in line:
                in_method = True
                indent_level = len(line) - len(line.lstrip())
                method_lines.append(line)
            elif in_method:
                if line.strip() and len(line) - len(line.lstrip()) <= indent_level and line.strip() and 'def ' in line:
                    # New method found
                    break
                method_lines.append(line)
        
        # Show method with context
        for line in method_lines[:100]:  # First 100 lines
            if line.strip():
                print(line)
            else:
                print()
                
        if len(method_lines) > 100:
            print("... (method continues)")
    
    else:
        print("No changes in src/gui.py")
    
    # Check specific issues
    print("\n=== Issue Checks ===")
    
    # Check for text_color
    stdout, stderr, code = run_cmd("grep -n 'text_color' src/gui.py || echo 'NO_TEXT_COLOR_FOUND'")
    if "NO_TEXT_COLOR_FOUND" in stdout:
        print("✓ PASS: No text_color found")
    else:
        print("✗ FAIL: text_color found!")
        print(stdout)
    
    # Check for temp_log.txt
    stdout, stderr, code = run_cmd("grep -n 'temp_log.txt' src/gui.py || echo 'NO_TEMP_LOG_FOUND'")
    if "NO_TEMP_LOG_FOUND" in stdout:
        print("✓ PASS: No temp_log.txt file saving")
    else:
        print("✗ FAIL: temp_log.txt file saving found")
        print(stdout)
    
    # Check if _update_tree_display has VERDICT
    stdout, stderr, code = run_cmd("grep -n 'VERDICT' src/gui.py | grep '_update_tree_display'")
    if stdout:
        print("✓ PASS: VERDICT found in _update_tree_display method")
        
        # Show where VERDICT appears
        lines = stdout.split('\n')
        print(f"  VERDICT appears at line numbers: {', '.join(lines)}")
    else:
        print("✗ FAIL: VERDICT not found in _update_tree_display")
    
    # Show first 20 lines of _update_tree_display for reference
    print("\n=== First 20 lines of _update_tree_display method ===")
    stdout, stderr, code = run_cmd("sed -n '/def _update_tree_display/,/^[[:space:]]*def /p' src/gui.py | head -50")
    print(stdout)

if __name__ == "__main__":
    main()
