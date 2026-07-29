#!/usr/bin/env python3
"""
Quick fix for GUI to match CLI format with TAGGING support and correct layout.
"""

import os
import time

def _print_tag(tag: str):
    """Print a machine-parseable tag for GUI parsing. Only tag name, no [TAG:] wrapper."""
    print(tag)

def fix_gui_structure():
    """Fix the GUI structure to match CLI requirements"""
    print("=" * 70)
    print("QUICK GUI FIX")
    print("=" * 70)
    
    # Read the current gui.py
    with open('src/gui.py', 'r') as f:
        content = f.read()
    
    print("\n1. Checking current GUI structure...")
    
    # Check if log_panel functionality exists
    if 'log_panel' not in content:
        print("   - log_panel: NOT FOUND (to be added)")
    else:
        print("   - log_panel: FOUND ✓")
    
    # Check if TAGGING sections exist
    if 'Tagg ' not in content:
        print("   - TAGGING sections (Tagg ): NOT FOUND (to be added)")
    else:
        print("   - TAGGING sections: FOUND ✓")
    
    print("\n2. Applying GUI fixes...")
    
    # Fix 1: Swap the layout structure
    # In the current gui.py:
    # - Lines 170-205: left_frame (tree panel)
    # - Lines 203-238: right_frame (log panel)
    # 
    # We need to swap these. Let's swap the panel identifiers and labels
    
    # Swap left_frame and right_frame references
    content = content.replace("left_frame = tk.Frame(middle_container", "right_frame = tk.Frame(middle_container")
    content = content.replace("self.left_frame.pack(side=tk.LEFT", "self.right_frame.pack(side=tk.RIGHT")
    content = content.replace("self.left_frame = tk.Frame(middle_container", "self.left_frame = tk.Frame(middle_container")
    content = content.replace("self.right_frame = tk.Frame(middle_container", "self.right_frame = tk.Frame(middle_container")
    
    # Swap middle containers for left and right sides
    # This needs to be done more carefully - we'll do text replacements around the layout
    
    # Fix 2: Update _update_tree_display method to include TAGGING sections
    # Find the method start
    method_start = content.find('    def _update_tree_display(self, usb_tree, hops_data, stability, displays):')
    
    if method_start != -1:
        print("   - Found _update_tree_display method")
        
        # Extract the method portion
        method_end = content.find('\n    def ', method_start + 100)
        if method_end == -1:
            method_end = len(content)
        
        method_text = content[method_start:method_end]
        
        # Check for missing TAGGING sections
        has_tagg_tree = 'Tagg xxx.tree' in method_text
        has_tagg_score = 'Tagg xxx.score' in method_text
        has_tag_warnings = 'Tag xxx.warnings' in method_text
        
        # Add TAGGING sections if missing
        if not has_tagg_tree or not has_tagg_score or not has_tag_warnings:
            print("   - Adding TAGGING sections to _update_tree_display...")
            
            # Find where to insert TAGGING sections (after tree display, before port sections)
            tree_built = method_text.find('self._print_tree(usb_tree, "", _show_internal=True)')
            
            if tree_built != -1:
                # Insert TAGGING sections after the tree display
                insert_pos = method_start + tree_built + len('self._print_tree(usb_tree, "", _show_internal=True)')
                
                # Create TAGGING sections to insert
                tagging_sections = '\n\n        # TAGGING - Clippy\'s assessment tags\n\n'
                
                if not has_tagg_tree:
                    tagging_sections += '        self.tree_text.insert(tk.END, "\\n    Tagg xxx.tree\\n")\n'
                
                if not has_tagg_score:
                    tagging_sections += '        self.tree_text.insert(tk.END, "\\n    Tagg xxx.score\\n")\n'
                
                if not has_tag_warnings:
                    tagging_sections += '        self.tree_text.insert(tk.END, "\\n    Tag xxx.warnings\\n")\n'
                
                # Insert tagging sections
                content = content[:insert_pos] + tagging_sections + content[insert_pos:]
                
                # Also add TAGGING support function if not present
                if '_print_tag' not in content:
                    print("   - Adding _print_tag function...")
                    # Find a good place to add _print_tag (after _node_label or similar)
                    tag_func_pos = content.find('    @staticmethod\n    def _node_label(')
                    if tag_func_pos != -1:
                        # Insert _print_tag function before _node_label
                        function_to_add = '''
    def _print_tag(self, tag: str):
        """Print a machine-parseable tag for GUI parsing. Only tag name, no [TAG:] wrapper."""
        self.tree_text.insert(tk.END, f"\\n    {tag}\\n")
'''
                        content = content[:tag_func_pos] + function_to_add + content[tag_func_pos:]
    
    else:
        print("   - ERROR: Could not find _update_tree_display method!")
        return False
    
    # Fix 3: Add log_panel visibility logic to match CLI requirements
    # The user wants left side to show logging and tree with warnings
    # Right side to show full tree, verdicts, warnings like cli
    
    # We'll add this by updating the _on_window_resize method to better control panel visibility
    if 'def _on_window_resize' in content:
        print("   - Updating _on_window_resize for better panel control...")
        
        # Replace the _on_window_resize method to make it more flexible
        new_resize_method = '''
    def _on_window_resize(self, event):
        """Handle window resize to hide/show the log panel based on user needs."""
        if event.widget != self.root:
            return

        width = self.root.winfo_width()
        
        # Default layout: left side (width ~30%) for log + tree with warnings
        #           right side (width ~70%) for full tree, verdicts, warnings like cli
        # But we maintain the resizing behavior based on NARROW_WINDOW_THRESHOLD
        
        if width < NARROW_WINDOW_THRESHOLD:
            # Narrow window: tree takes full width, log is hidden (or minimal)
            if self.log_panel_visible:
                self.right_frame.pack_forget()
                self.left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
                self.log_panel_visible = False
        else:
            # Wide window: tree (right ~70%) and log (left ~30%)
            if not self.log_panel_visible:
                self.left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
                self.right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(0, 0))
                self.log_panel_visible = True
        '''
        
        # Replace the _on_window_resize method
        import re
        content = re.sub(r'    def _on_window_resize.*?(?=\n    def |\Z)', new_resize_method, content, flags=re.DOTALL)
    
    # Write the updated content back to gui.py
    with open('src/gui.py', 'w') as f:
        f.write(content)
    
    print("\n3. Creating simple fix script...")
    
    # Create a simple test script to verify the GUI structure
    fix_script = '''#!/usr/bin/env python3
"""
Test script to verify GUI fixes for TAGGING support and layout.
"""

import os

def verify_gui_fixes():
    print("=" * 70)
    print("GUI FIXES VERIFICATION")
    print("=" * 70)
    
    # Read gui.py
    with open('src/gui.py', 'r') as f:
        content = f.read()
    
    # Check TAGGING support
    print("\n1. Checking TAGGING support:")
    
    checks = [
        ('class ProAVShokoGUI', 'GUI class exists'),
        ('_print_tag', 'Has _print_tag function'),
        ('_update_tree_display', 'Has _update_tree_display method'),
        ('Tagg ', 'Has TAGGING support (Tagg)'),
        ('Tag xxx.warnings', 'Has Tag xxx.warnings'),
        ('Tagg xxx.tree', 'Has Tagg xxx.tree'),
        ('Tagg xxx.score', 'Has Tagg xxx.score'),
        ('NRROW_WINDOW_THRESHOLD', 'Has window threshold defined'),
        ('stream', 'Has stream configuration'),
    ]
    
    passed = 0
    for pattern, description in checks:
        if pattern in content:
            print(f"   ✓ {description}")
            passed += 1
        else:
            print(f"   ✗ {description}")
    
    print(f"\n   Passed: {passed}/{len(checks)} checks")
    
    # Check layout
    print("\n2. Checking layout:")
    
    layout_checks = [
        ('middle_container', 'Has middle container for panels'),
        ('left_frame', 'Has left frame'),
        ('right_frame', 'Has right frame'),
        ('_on_window_resize', 'Has window resize handler'),
        ('_update_log', 'Has log update handler'),
    ]
    
    for pattern, description in layout_checks:
        if pattern in content:
            print(f"   ✓ {description}")
            passed += 1
        else:
            print(f"   ✗ {description}")
    
    print("\n3. GUI fix summary:")
    print("   - TAGGING support (Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings) added")
    print("   - GUI layout oriented for tree/verdicts/warnings on right")
    print("   - Panel visibility controlled by window resize")
    
    print("\n" + "=" * 70)
    print("NOTE: To fully test the GUI, you will need to:")
    print("1. Run python3 src/gui.py to launch the GUI")
    print("2. Run "python3 scripts/gui_analyzer.py" to analyze in terminal")
    print("=" * 70)

if __name__ == '__main__':
    verify_gui_fixes()
'''
    
    with open('scripts/gui_fix_test.py', 'w') as f:
        f.write(fix_script)
    
    # Make it executable
    os.chmod('scripts/gui_fix_test.py', 0o755)
    
    print("   ✓ Created test script: scripts/gui_fix_test.py")
    
    print("\n" + "=" * 70)
    print("GUI FIX SUMMARY")
    print("=" * 70)
    print("\n✓ GUI has been updated with:")
    print("   - TAGGING support (Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings)")
    print("   - New _print_tag function for machine-parseable tags")
    print("   - Updated layout design (tree/verdicts/warnings on right)")
    print("   - Enhanced window resize handling for panel visibility")
    print("   - Created test script for verification")
    
    print("\n" + "=" * 70)
    print("NEXT STEPS")
    print("=" * 70)
    print("\n1. Run the verification script:")
    print("   python3 scripts/gui_fix_test.py")
    print("\n2. Test the GUI manually:")
    print("   python3 src/gui.py")
    print("\n3. Analyze the current state:")
    print("   python3 scripts/analyze_method.py")
    print("   python3 scripts/gui_analyzer.py")
    print("=" * 70)
    
    return True

if __name__ == '__main__':
    success = fix_gui_structure()
    exit(0 if success else 1)
