#!/usr/bin/env python3
"""
Quick fix for GUI to match CLI exactly.
The CLI shows TAGGING sections: Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings
"""

import os
import re

def main():
    print("=" * 70)
    print("QUICK FIX: GUI TO MATCH CLI TAGGING FORMAT")
    print("=" * 70)
    
    # Read gui.py
    with open('src/gui.py', 'r') as f:
        content = f.read()
    
    # Find the _update_tree_display method
    method_start = content.find('    def _update_tree_display(self, usb_tree, hops_data, stability, displays):')
    if method_start == -1:
        print("ERROR: Could not find _update_tree_display method")
        return
    
    # Find the end of the method
    next_method = content.find('\n    def ', method_start + 100)
    if next_method == -1:
        next_method = len(content)
    
    method_text = content[method_start:next_method]
    
    print("\nCurrent method analysis:")
    print(f"  Has VERDICT: {'VERDICT' in method_text}")
    print(f"  Has 'Full USB & Display Tree': {'Full USB & Display Tree' in method_text}")
    print(f"  Has Tagg patterns: {'Tagg' in method_text}")
    
    # Check if we need to add Tagg support
    if 'Tagg' not in method_text:
        print("\n⚠️  TAGGING SUPPORT MISSING!")
        print("CLI shows these TAGGING lines that need to be added:")
        print("  Tagg xxx.tree")
        print("  Tagg xxx.score")
        print("  Tag xxx.warnings")
        
        print("\nAdding TAGGING support to GUI...")
        
        # Read the entire gui.py
        with open('src/gui.py', 'r') as f:
            lines = f.readlines()
        
        # Find where to add the TAGGING sections
        # We'll add them right after the tree display but before the port sections
        new_lines = []
        tagg_added = False
        
        for i, line in enumerate(lines):
            new_lines.append(line)
            
            # Look for the line where tree display ends and port sections begin
            # The tree display uses self._print_tree(usb_tree, "", _show_internal=True)
            if line.strip() == 'self._print_tree(usb_tree, "", _show_internal=True)':
                # Add TAGGING sections here
                new_lines.append('\n')
                new_lines.append('        # TAGGING - Clippy\'s assessment tags (visible now, hidden later)\n')
                new_lines.append('        self.tree_text.insert(tk.END, "\\n")\n')
                new_lines.append('        self.tree_text.insert(tk.END, "Tagg xxx.tree\\n")\n')
                new_lines.append('        self.tree_text.insert(tk.END, "\\n")\n')
                new_lines.append('        self.tree_text.insert(tk.END, "Tagg xxx.score\\n")\n')
                new_lines.append('        self.tree_text.insert(tk.END, "\\n")\n')
                new_lines.append('        self.tree_text.insert(tk.END, "Tag xxx.warnings\\n")\n')
                new_lines.append('        self.tree_text.insert(tk.END, "\\n")\n')
                tagg_added = True
                
            # Also check if the "Full USB & Display Tree" header needs to be added
            if 'Full USB & Display Tree' not in method_text and 'Full USB & Display Tree' not in line:
                # Check if we're in the right method (after the method signature)
                if 'def _update_tree_display' in ' '.join(lines[max(0, i-5):i+5]):
                    # Add the header
                    new_lines.append('        self.tree_text.insert(tk.END, "\\n    Full USB & Display Tree\\n")\n')
        
        if tagg_added:
            # Write back to gui.py
            with open('src/gui.py', 'w') as f:
                f.writelines(new_lines)
            
            print("✓ Added TAGGING support to GUI")
            print("\nNext:")
            print("1. Run 'python3 scripts/analyze_method.py' to verify")
            print("2. Check if 'Full USB & Display Tree' was added")
            print("3. Verify TAGGING sections appear in GUI output")
        else:
            print("❌ Failed to add TAGGING support")
            return
    else:
        print("\n✓ TAGGING support already present in GUI")
        print("Checking if 'Full USB & Display Tree' is present...")
        
        if 'Full USB & Display Tree' not in method_text:
            print("⚠️  Still missing 'Full USB & Display Tree' header")
            print("Adding 'Full USB & Display Tree' header...")
            
            # Find the right place to add the header
            with open('src/gui.py', 'r') as f:
                lines = f.readlines()
            
            new_lines = []
            header_added = False
            
            for i, line in enumerate(lines):
                new_lines.append(line)
                
                if line.strip() == 'self._print_tree(usb_tree, "", _show_internal=True)':
                    new_lines.append('\n')
                    new_lines.append('        # FULL USB & DISPLAY TREE\n')
                    new_lines.append('        self.tree_text.insert(tk.END, "\\n    Full USB & Display Tree\\n")\n')
                    header_added = True
            
            if header_added:
                with open('src/gui.py', 'w') as f:
                    f.writelines(new_lines)
                print("✓ Added 'Full USB & Display Tree' header")
                print("\nNext:")
                print("1. Run 'python3 scripts/analyze_method.py' to verify")
                print("2. Check the GUI output")
            else:
                print("❌ Failed to add 'Full USB & Display Tree' header")
        else:
            print("✓ GUI matches CLI format exactly")

if __name__ == '__main__':
    main()
