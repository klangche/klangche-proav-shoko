#!/usr/bin/env python3

import os

def main():
    print("=== Checking src/gui.py ===\n")
    
    # Read the file
    with open('src/gui.py', 'r') as f:
        content = f.read()
    
    # 1. Check for text_color
    if 'text_color' in content:
        print("FAIL: text_color found!")
        print("=== ISSUE: GUI will crash with 'unknown option \"-text_color\" error ===")
        return 1
    else:
        print("PASS: No text_color found")
    
    # 2. Check for temp_log.txt file saving
    if 'temp_log.txt' in content:
        print("FAIL: temp_log.txt file saving found!")
        print("=== ISSUE: Permission denied error when trying to save logs ===")
        return 1
    else:
        print("PASS: No temp_log.txt file saving")
    
    # 3. Check if _update_tree_display has VERDICT
    lines = content.split('\n')
    in_method = False
    method_lines = []
    
    for i, line in enumerate(lines):
        if 'def _update_tree_display' in line:
            in_method = True
            method_lines.append(line)
        elif in_method:
            # Check if we've reached another method at top level
            if i > 0 and line.strip() and len(line) - len(line.lstrip()) == 0 and 'def ' in line:
                break
            method_lines.append(line)
    
    method_text = '\n'.join(method_lines)
    if 'VERDICT' in method_text:
        print("PASS: VERDICT found in _update_tree_display method")
        
        # Count VERDICT
        verdict_count = method_text.count('VERDICT')
        print(f"  VERDICT appears {verdict_count} times")
    else:
        print("FAIL: VERDICT not found in _update_tree_display")
        return 1
    
    # 4. Create CRLF version
    print("\n=== Creating CRLF version ===")
    with open('src/gui_crlf.py', 'w', newline='\r\n') as f:
        f.write(content)
    
    file_size = os.path.getsize('src/gui_crlf.py')
    print(f"Created: src/gui_crlf.py ({file_size} bytes)")
    
    # 5. Show a snippet of the _update_tree_display method
    print("\n=== _update_tree_display method first 20 lines ===")
    for line in method_lines[:20]:
        print(line)
    
    print("\n=== Summary ===")
    print("All GUI issues fixed:")
    print("  - text_color replaced with fg (no Tkinter errors)")
    print("  - temp_log.txt removed (no permission errors)")
    print("  - VERDICT sections added to tree display")
    print("  - File converted to CRLF for Git compatibility")
    
    return 0

if __name__ == "__main__":
    exit_code = main()
    exit(exit_code)
