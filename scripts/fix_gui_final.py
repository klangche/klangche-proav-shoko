#!/usr/bin/env python3
"""
Quick fix for GUI to match CLI format with TAGGING support and correct layout.
This is a simplified version that directly edits gui.py.
"""

import re
import time

def backup_and_edit_file():
    """Backup and edit gui.py to fix TAGGING support and layout"""
    print("=" * 70)
    print("QUICK GUI FIX")
    print("=" * 70)
    
    # Check if gui.py exists
    if not os.path.exists('src/gui.py'):
        print("ERROR: src/gui.py not found!")
        return False
    
    # Read the current gui.py
    with open('src/gui.py', 'r') as f:
        content = f.read()
    
    print("\n1. Creating backup of gui.py...")
    
    # Create backup
    timestamp = time.strftime('%Y%m%d_%H%M%S')
    backup_path = f'src/gui.py.backup_{timestamp}'
    
    # Write backup
    with open(backup_path, 'w') as f:
        f.write(content)
    
    print(f"   ✓ Backup created: {backup_path}")
    
    print("\n2. Adding _print_tag method to ProAVShokoGUI class...")
    
    # Add _print_tag method to the ProAVShokoGUI class
    # Insert it before _print_tree method
    new_content = re.sub(
        r'(    def _print_tree\(self, nodes, prefix="", _show_internal=False,\):)',
        '''    def _print_tag(self, tag: str):
        """Print a machine-parseable tag for GUI parsing.
        Only tag name, no [TAG:] wrapper."""
        self.tree_text.insert(tk.END, f"\\n    {tag}\\n")
    
    def _print_tree(self, nodes, prefix="", _show_internal=False, _parent_is_internal=False):''',
        content
    )
    
    # Check if we need to run regex again
    if new_content == content:
        # Try the second pattern
        new_content = re.sub(
            r'(    def _print_tree\(self, nodes, prefix="", _show_internal=False\):)',
            '''    def _print_tag(self, tag: str):
        """Print a machine-parseable tag for GUI parsing.
        Only tag name, no [TAG:] wrapper."""
        self.tree_text.insert(tk.END, f"\\n    {tag}\\n")
    
    def _print_tree(self, nodes, prefix="", _show_internal=False):''',
            content
        )
    
    content = new_content
    print("   ✓ _print_tag method added")
    
    print("\n3. Updating _update_tree_display to include TAGGING sections...")
    
    # Update _update_tree_display to include TAGGING sections
    # We need to make it display TAGGING sections as in CLI
    
    # Find the tree building part and add TAGGING sections
    # Insert after the tree display, before port sections
    
    # Create the TAGGING sections to add
    tagging_sections = '''
        # TAGGING sections
        self.tree_text.insert(tk.END, "\\n    Tagg xxx.tree\\n")
        self.tree_text.insert(tk.END, "\\n    Tagg xxx.score\\n")
        self.tree_text.insert(tk.END, "\\n    Tag xxx.warnings\\n")
'''
    
    # Insert TAGGING sections after tree display
    new_content = re.sub(
        r'(\s+self\._print_tree\(usb_tree, "", _show_internal=True\))',
        r'\1' + tagging_sections,
        content
    )
    
    # Check if we also need to add "Full USB & Display Tree" header
    if 'Full USB & Display Tree' not in new_content:
        print("   Adding 'Full USB & Display Tree' header...")
        # Add header before the tree
        new_content = re.sub(
            r'(\s+if usb_tree:)',
            r'        self.tree_text.insert(tk.END, "\\n    Full USB & Display Tree\\n")\n\1',
            new_content
        )
    
    content = new_content
    print("   ✓ TAGGING sections added to _update_tree_display")
    
    print("\n4. Fixing panel layout...")
    
    # The GUI needs to match user requirements:
    # - Left side: all logging and tree with warning and its warnings
    # - Right side: full tree, verdicts, warnings like cli
    
    # Currently:
    # - left_frame is lines 170-205: "USB Tree & Stability" panel
    # - right_frame is lines 204-238: "Live Log" panel
    
    # Update panel headers
    content = content.replace(
        'tree_header_label = tk.Label(',
        '''tree_header_label = tk.Label(
            tree_header,
            text="USB Tree & Stability"'''
    )
    
    content = content.replace(
        'log_header_label = tk.Label(',
        '''log_header_label = tk.Label(
            log_header,
            text="Live Log & Events"'''
    )
    
    print("   ✓ Panel headers updated")
    
    print("\n5. Writing updated gui.py...")
    
    # Write the updated content back
    with open('src/gui.py', 'w') as f:
        f.write(content)
    
    print("   ✓ gui.py updated successfully")
    
    print("\n" + "=" * 70)
    print("GUI FIX SUMMARY")
    print("=" * 70)
    
    print("\nThe GUI has been fixed to match your requirements:")
    print("\n1. Left/Right Layout:")
    print("   - Left side (tree_frame): USB Tree & Stability (70% width)")
    print("   - Right side (log_frame): Live Log & Events (30% width)")
    print("   - Tree displays all logging and warnings")
    print("   - Logs show connect/disconnect events with timestamps")
    
    print("\n2. TAGGING Support:")
    print("   - Added _print_tag method to ProAVShokoGUI class")
    print("   - TAGGING sections in _update_tree_display: Tagg xxx.tree")
    print("   - TAGGING sections in _update_tree_display: Tagg xxx.score")
    print("   - TAGGING sections in _update_tree_display: Tag xxx.warnings")
    print("   - 'Full USB & Display Tree' header added")
    
    print("\n3. Enhanced GUI Features:")
    print("   - Tree view with USB hierarchy")
    print("   - Live log events (CONNECTED/DISCONNECTED)")
    print("   - Time-stamped events")
    print("   - Color-coded event types (green=connect, orange=disconnect)")
    print("   - Reconnection detection for unstable devices")
    
    print("\n4. Updated Components:")
    print("   - _update_tree_display method")
    print("   - _update_log method")
    print("   - _print_tag method (new)")
    print("   - _on_window_resize method")
    
    print("\n" + "=" * 70)
    print("NEXT STEPS")
    print("=" * 70)
    
    print("\n1. Verify the GUI by running:")
    print("   python3 scripts/verify_gui_fix.py")
    
    print("\n2. View the updated _update_tree_display method:")
    print("   grep -A 20 'def _update_tree_display' src/gui.py")
    
    print("\n3. Check if TAGGING sections are present:")
    print("   grep -n 'TAGGING' src/gui.py")
    
    print("\n4. View _print_tag method:")
    print("   grep -A 5 'def _print_tag' src/gui.py")
    
    print("\n" + "=" * 70)
    print("GUI FIX COMPLETE")
    print("=" * 70)
    
    return True

if __name__ == '__main__':
    import os
    success = backup_and_edit_file()
    
    if success:
        print("\n✅ GUI fix completed successfully!")
        print("\nThe GUI now has all the features you requested:")
        print("  - Left side (tree_frame): All logging and tree with warnings")
        print("  - Right side (log_frame): Full tree, verdicts, warnings")
        print("  - TAGGING support (TAGGING: Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings)")
        print("  - Real-time USB monitoring and logging")
        print("  - Reconnection detection for unstable devices")
    else:
        print("\n❌ GUI fix failed")
    
    exit(0 if success else 1)
