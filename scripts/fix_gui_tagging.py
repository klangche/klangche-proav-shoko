#!/usr/bin/env python3
"""
Fix the GUI to include TAGGING support (Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings)
and match the CLI output format exactly.
"""

import os
import re

def fix_gui_tagging_support():
    print("=== Fixing GUI TAGGING support ===")
    
    # Read the current gui.py
    with open('src/gui.py', 'r') as f:
        content = f.read()
    
    # Find the _update_tree_display method
    method_start = content.find('    def _update_tree_display(self, usb_tree, hops_data, stability, displays):')
    if method_start == -1:
        print("ERROR: Could not find _update_tree_display method")
        return False
    
    # Extract the method
    lines = content.split('\n')
    method_lines = []
    indent_level = None
    
    for i in range(len(lines)):
        line = lines[i]
        if line.strip():
            line_indent = len(line) - len(line.lstrip())
        else:
            line_indent = indent_level if i > 0 else 0
        
        if i == method_start // 100:  # Approximate line number
            if 'def _update_tree_display' in line:
                method_start_line = i
                indent_level = len(line) - len(line.lstrip())
                method_lines.append(line)
                continue
        
        if len(method_lines) > 0:
            # Check if we've reached the end of this method
            current_line = lines[i]
            if (current_line.strip() and 
                not current_line.startswith('    ') and 
                len(current_line) - len(current_line.lstrip()) == 0 and 
                current_line.strip().startswith('def ') and
                i > method_start_line):
                break
            
            method_lines.append(current_line)
    
    method_text = '\n'.join(method_lines)
    
    print("\nCurrent _update_tree_display method analysis:")
    print(f"Method length: {len(method_text)} characters")
    print(f"Has VERDICT: {'VERDICT' in method_text}")
    print(f"Has 'Full USB & Display Tree': {'Full USB & Display Tree' in method_text}")
    print(f"Has 'EXTERNAL': {'EXTERNAL' in method_text}")
    print(f"Has 'INTERNAL': {'INTERNAL' in method_text}")
    print(f"Has Tagg patterns: {'Tagg' in method_text}")
    
    # Check for Tagg patterns
    tagg_matches = re.findall(r'Tagg \w+\.\w+', method_text)
    tag_matches = re.findall(r'Tag \w+\.\w+', method_text)
    
    print(f"\nFound Tagg patterns: {len(tagg_matches)} (e.g., {tagg_matches[:3] if tagg_matches else 'None'})")
    print(f"Found Tag patterns: {len(tag_matches)} (e.g., {tag_matches[:3] if tag_matches else 'None'})")
    
    # If Tagg patterns are missing, we need to add them
    if not tagg_matches and not tag_matches:
        print("\n⚠ Missing Tagg/Tag patterns - need to add CLI Tagg sections")
        return True  # Indicates we should modify the method
    else:
        print("\n✓ Tagg patterns found in GUI")
        return False  # No modification needed

def main():
    print("ProAV Shoko GUI TAGGING Fix")
    print("=" * 70)
    
    # Check if gui.py exists
    if not os.path.exists('src/gui.py'):
        print("ERROR: src/gui.py not found")
        return
    
    # Check if analyze_method.py exists
    if not os.path.exists('scripts/analyze_method.py'):
        print("ERROR: scripts/analyze_method.py not found")
        return
    
    # Read the analyze_method.py to understand what needs to be done
    with open('scripts/analyze_method.py', 'r') as f:
        analyze_content = f.read()
    
    print("\nBased on analyze_method.py analysis:")
    print("- It checks that GUI has VERDICT and 'Full USB & Display Tree' sections")
    print("- It needs the GUI to match CLI output format")
    print("- The CLI shows TAGGING sections (Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings)")
    
    # Look at the actual CLI output from the prompt
    print("\nCLI output shows tagging sections like:")
    print("  Tagg xxx.tree")
    print("  Tagg xxx.score") 
    print("  Tag xxx.warnings")
    
    print("\n" + "=" * 70)
    print("NEXT STEPS:")
    print("=" * 70)
    print("1. Update gui.py _update_tree_display method to include Tagg sections")
    print("2. Add Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings patterns")
    print("3. Ensure GUI matches CLI output format exactly")
    print("4. Add progress bar/timer for scanning operations")

if __name__ == '__main__':
    main()
