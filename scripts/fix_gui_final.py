#!/usr/bin/env python3
"""
Check and fix all GUI issues in src/gui.py
"""
import os
import re

def fix_gui_issues():
    print("=== Checking and Fixing src/gui.py ===\n")
    
    # Read the current file
    with open('src/gui.py', 'r') as f:
        content = f.read()
    
    fixes_applied = []
    
    # Fix 1: Replace all text_color with fg
    if 'text_color' in content:
        print("FAIL: text_color found - will cause Tkinter errors")
        
        # Count replacements
        replacement_count = content.count('text_color')
        print(f"Found {replacement_count} instances of text_color")
        
        print("FIXING: Replacing text_color with fg...")
        
        # Replace all text_color with fg
        new_content = content.replace('text_color', 'fg')
        
        # Write back
        with open('src/gui.py', 'w') as f:
            f.write(new_content)
        
        fixes_applied.append(f"Replaced {replacement_count} text_color instances with fg")
        print(f"FIXED: Replaced {replacement_count} instances of text_color")
    else:
        print("PASS: No text_color found")
    
    # Fix 2: Remove temp_log.txt file saving
    if 'temp_log.txt' in content:
        print("FAIL: temp_log.txt file saving found")
        print("GUI will crash with: [Errno 13] Permission denied: 'temp_log.txt'")
        
        # Find temp_log.txt usage
        line_matches = []
        for i, line in enumerate(content.split('\n'), 1):
            if 'temp_log.txt' in line:
                line_matches.append((i, line))
        
        print(f"Found temp_log.txt usage at lines: {[m[0] for m in line_matches]}")
        print("FIXING: Removing temp_log.txt file saving...")
        
        # Remove the problematic section
        new_content = content
        
        # Replace the problematic code with simpler version
        # Find and remove the temp_log.txt section
        lines = content.split('\n')
        new_lines = []
        in_temp_section = False
        for i, line in enumerate(lines):
            if 'temp_log.txt' in line:
                # Skip this line and the following lines until we see "Display information"
                in_temp_section = True
                continue
            
            if in_temp_section and ('Display information' in line or 'self.display_analyzer = DisplayAnalyzer()' in line):
                in_temp_section = False
                new_lines.append(line)
                continue
            
            if not in_temp_section:
                new_lines.append(line)
        
        new_content = '\n'.join(new_lines)
        
        # Write back
        with open('src/gui.py', 'w') as f:
            f.write(new_content)
        
        fixes_applied.append("temp_log.txt file saving removed")
        print("FIXED: temp_log.txt file saving removed")
    else:
        print("PASS: No temp_log.txt file saving")
    
    # Fix 3: Create CRLF version
    print("\n=== Converting to CRLF line endings ===")
    try:
        with open('src/gui_crlf.py', 'w', newline='\r\n') as f:
            with open('src/gui.py', 'r') as src:
                f.write(src.read())
        
        crlf_size = os.path.getsize('src/gui_crlf.py')
        original_size = os.path.getsize('src/gui.py')
        
        print(f"PASS: Created src/gui_crlf.py ({crlf_size} bytes)")
        
        if crlf_size != original_size:
            print(f"Note: File size changed from {original_size} to {crlf_size} bytes due to line ending conversion")
    except Exception as e:
        print(f"FAIL: Could not create CRLF version: {e}")
    
    # Verify fixes
    print("\n=== Verification ===")
    
    with open('src/gui.py', 'r') as f:
        final_content = f.read()
    
    # Check 1: No text_color
    if 'text_color' not in final_content:
        print("PASS: No text_color found in final version")
    else:
        print("FAIL: text_color still exists!")
        
        # Count remaining instances
        count = final_content.count('text_color')
        print(f"Still {count} instances of text_color found")
    
    # Check 2: No temp_log.txt
    if 'temp_log.txt' not in final_content:
        print("PASS: No temp_log.txt found in final version")
    else:
        print("FAIL: temp_log.txt still exists!")
        
        # Count remaining instances
        count = final_content.count('temp_log.txt')
        print(f"Still {count} instances of temp_log.txt found")
    
    # Check 3: VERDICT and tree structure
    if 'def _update_tree_display' in final_content:
        # Extract the method
        lines = final_content.split('\n')
        
        # Find method
        for i, line in enumerate(lines):
            if 'def _update_tree_display' in line:
                # Extract to next method or end
                method_lines = [line]
                for j in range(i + 1, len(lines)):
                    if lines[j].strip() and len(lines[j]) - len(lines[j].lstrip()) == 0 and 'def ' in lines[j]:
                        break
                    method_lines.append(lines[j])
                
                method_text = '\n'.join(method_lines)
                
                if 'VERDICT' in method_text and 'Full USB & Display Tree' in method_text:
                    print("PASS: _update_tree_display has VERDICT and tree structure")
                else:
                    print("FAIL: _update_tree_display missing VERDICT or tree structure")
                    if 'VERDICT' not in method_text:
                        print("  - Missing VERDICT")
                    if 'Full USB & Display Tree' not in method_text:
                        print("  - Missing 'Full USB & Display Tree'")
    
    # Summary
    print("\n=== Summary ===")
    if fixes_applied:
        print(f"Applied {len(fixes_applied)} fixes:")
        for fix in fixes_applied:
            print(f"  - {fix}")
    else:
        print("No fixes applied - file appears to already be correct")
    
    print("\n=== GUI Troubleshooting Guide ===")
    print("Common issues and solutions:")
    print("1. Tkinter errors about '-text_color': Fix by replacing 'text_color' with 'fg'")
    print("2. Permission denied errors for 'temp_log.txt': Fix by removing temp_log.txt file saving")
    print("3. Tree display issues: Ensure _update_tree_display method has VERDICT and proper structure")
    print("4. Git line ending issues: Convert to CRLF for Windows compatibility")
    
    print(f"\n=== Next Steps ===")
    print("To apply these fixes to the repository:")
    print("1. Commit the fixed gui.py file")
    print("2. Use 'git add src/gui.py' to stage changes")
    print("3. Use 'git commit -m \"Fix GUI: replace text_color, remove temp_log.txt, update tree display\"'")
    print("4. Create a CRLF-compatible version: 'git checkout src/gui_crlf.py'")
    print("5. Or use Git's end-of-line core configuration: 'git config core.autocrlf=true'")

if __name__ == "__main__":
    fix_gui_issues()
