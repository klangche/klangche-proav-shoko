#!/usr/bin/env python3

import re
import os

def check_gui_issues():
    print("=== Checking src/gui.py for issues ===\n")
    
    # Read the file
    with open('src/gui.py', 'r') as f:
        content = f.read()
    
    # Check for text_color
    if 'text_color' in content:
        print("FAIL: text_color found!")
        
        # Find all matches
        lines = content.split('\n')
        for i, line in enumerate(lines, 1):
            if 'text_color' in line:
                print(f"  Line {i}: {line.strip()}")
    else:
        print("PASS: No text_color found")
    
    # Check for temp_log.txt saving (dangerous operation)
    if 'temp_log.txt' in content:
        print("FAIL: temp_log.txt file saving found!")
        
        # Find all matches
        lines = content.split('\n')
        for i, line in enumerate(lines, 1):
            if 'temp_log.txt' in line:
                print(f"  Line {i}: {line.strip()}")
    else:
        print("PASS: No temp_log.txt file saving")
    
    # Check if _update_tree_display has VERDICT
    lines = content.split('\n')
    in_method = False
    method_lines = []
    
    for line in lines:
        if 'def _update_tree_display' in line:
            in_method = True
            method_lines.append(line)
        elif in_method:
            method_lines.append(line)
            if line.strip() and 'def ' in line and not line.strip().startswith('        '):
                # End of the method
                break
    
    # Check for VERDICT in the method
    method_text = '\n'.join(method_lines)
    if 'VERDICT' in method_text:
        print("PASS: VERDICT found in _update_tree_display method")
        
        # Count VERDICT occurrences
        verdict_count = method_text.count('VERDICT')
        print(f"  VERDICT appears {verdict_count} times")
    else:
        print("FAIL: VERDICT not found in _update_tree_display")
    
    # Show context around _update_tree_display
    print("\n=== Context around _update_tree_display ===")
    in_method = False
    indent_level = None
    
    for line in lines:
        if 'def _update_tree_display' in line:
            in_method = True
            indent_level = len(line) - len(line.lstrip())
            print(line)
        elif in_method:
            if line.strip() and len(line) - len(line.lstrip()) <= indent_level and 'def ' in line and not line.strip().startswith('        '):
                break
            print(line)
    
    # Create a backup or version with CRLF line endings
    print("\n=== Creating CRLF version ===")
    with open('src/gui_crlf.py', 'w', newline='\r\n') as f:
        f.write(content)
    print(f"Created src/gui_crlf.py with CRLF line endings")
    
    print("\n=== Summary ===")
    if 'text_color' not in content and 'temp_log.txt' not in content and 'VERDICT' in method_text:
        print("✓ GUI is safe to run - no permission errors or text_color issues")
        print("✓ Tree display updates with VERDICT sections")
        return True
    else:
        print("✗ GUI has issues that need fixing")
        return False

if __name__ == "__main__":
    check_gui_issues()
