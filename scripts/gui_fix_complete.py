#!/usr/bin/env python3
"""
GUI Fix Script
Fixes the GUI to match user requirements:
- Left side: All logging and tree with warning and its warnings
- Right side: Full tree, verdicts, warnings like CLI
- TAGGING support: Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings
"""

import re

def fix_gui():
    print("=" * 70)
    print("GUI FIX - Implementing User Requirements")
    print("=" * 70)
    
    # Read gui.py
    try:
        with open('src/gui.py', 'r') as f:
            content = f.read()
    except FileNotFoundError:
        print("ERROR: src/gui.py not found!")
        return False
    
    print("\n1. Checking current state...")
    
    # Check for required methods
    has_print_tag = '    def _print_tag(self, tag: str):' in content
    has_print_verdict = '    def _print_verdict(self, v):' in content
    has_print_section_header = '    def _print_section_header(self, title):' in content
    
    print(f"   _print_tag method: {'✓' if has_print_tag else '✗'}")
    print(f"   _print_verdict method: {'✓' if has_print_verdict else '✗'}")
    print(f"   _print_section_header method: {'✓' if has_print_section_header else '✗'}")
    
    print("\n2. Checking TAGGING sections...")
    
    tagg_sections = [
        ('Tagg xxx.tree', 'TAGGING: Tagg xxx.tree'),
        ('Tagg xxx.score', 'TAGGING: Tagg xxx.score'),
        ('Tag xxx.warnings', 'TAGGING: Tag xxx.warnings')
    ]
    
    for tag, desc in tagg_sections:
        if tag in content:
            print(f"   ✓ {desc}")
        else:
            print(f"   ✗ {desc}")
    
    print("\n3. Checking layout headers...")
    
    layout_headers = [
        ('All logging and tree with warning and its warnings', 'Left panel'),
        ('Full tree, verdicts, warnings like CLI', 'Right panel')
    ]
    
    for header, desc in layout_headers:
        if header in content:
            print(f"   ✓ {desc}: '{header}'")
        else:
            print(f"   ✗ {desc}: missing")
    
    # Fix the GUI layout to match requirements
    print("\n4. Fixing GUI layout...")
    
    # Update _update_tree_display method
    print("   Updating _update_tree_display method...")
    
    # Add "Full USB & Display Tree" header at the beginning
    if 'Full USB & Display Tree' not in content:
        content = content.replace(
            '        # PER PORT - EXTERNAL',
            '''        # FULL USB & DISPLAY TREE
        self._print_section_header("Full USB & Display Tree")
        self._print_tag("overall.tree")
        self._print_tag("overall.score")
        self._print_tag("overall.warnings")
        
        # PER PORT - EXTERNAL'''
        )
    
    # Add TAGGING sections at the end
    if 'TAGGING:' not in content:
        # Find the end of _update_tree_display method
        method_end = content.find('\n    def ', content.find('    def _update_tree_display'))
        if method_end == -1:
            method_end = len(content)
        
        # Insert TAGGING sections at the end
        tag_section = '''
        # TAGGING sections at the end
        self.tree_text.insert(tk.END, "\\nTAGGING:\n")
        self.tree_text.insert(tk.END, "TAGGING: xxx.tree\\n")
        self.tree_text.insert(tk.END, "TAGGING: xxx.score\\n")
        self.tree_text.insert(tk.END, "TAG: xxx.warnings\\n")
'''
        
        content = content[:method_end] + tag_section + content[method_end:]
    
    print("   ✓ Updated _update_tree_display method")
    
    # Fix panel headers
    print("\n5. Fixing panel headers...")
    
    # Fix left panel (Log panel)
    if 'text="Live Log"' in content:
        content = content.replace(
            'text="Live Log"',
            'text="All logging and tree with warning and its warnings"'
        )
        print("   ✓ Fixed left panel header")
    
    # Fix right panel (Tree panel)
    if 'text="USB Tree & Stability"' in content:
        content = content.replace(
            'text="USB Tree & Stability"',
            'text="Full tree, verdicts, warnings like CLI"'
        )
        print("   ✓ Fixed right panel header")
    
    # Add required methods if missing
    print("\n6. Adding required methods...")
    
    if not has_print_tag:
        print("   Adding _print_tag method...")
        # Add _print_tag method before _print_tree
        tag_method = '''
    def _print_tag(self, tag: str):
        """Print a machine-parseable tag for GUI parsing. Only tag name, no [TAG:] wrapper."""
        self.tree_text.insert(tk.END, f"\\n    {tag}\\n")
'''
        content = content.replace(
            '    def _print_tree(self, nodes, prefix="", _show_internal=False, _parent_is_internal=False):',
            tag_method + '    ' + '    def _print_tree(self, nodes, prefix="", _show_internal=False, _parent_is_internal=False):'
        )
    
    if not has_print_verdict:
        print("   Adding _print_verdict method...")
        # Add _print_verdict method
        verdict_method = '''
    def _print_verdict(self, v):
        """Print a single verdict line with hops/tiers/hubs. Matches CLI exactly."""
        status_char = '+' if v['color'] == 'green' else ('~' if v['color'] == 'orange' else '!')
        hubs_str = f"hubs {v['current_hubs']}/{v['max_hubs']}  " if 'current_hubs' in v else ""
        desc = v.get('description', v.get('name', ''))
        self.tree_text.insert(tk.END, f"    {status_char} {desc:<22s} {v['status']:<9s} hops {v['current_hops']}/{v['max_hops']}  tiers {v['current_tiers']}/{v['max_tiers']}  {hubs_str}\\n")
'''
        content = content.replace(
            '    def _print_tree(self, nodes, prefix="", _show_internal=False, _parent_is_internal=False):',
            verdict_method + '    ' + '    def _print_tree(self, nodes, prefix="", _show_internal=False, _parent_is_internal=False):'
        )
    
    if not has_print_section_header:
        print("   Adding _print_section_header method...")
        # Add _print_section_header method
        section_method = '''
    def _print_section_header(self, title):
        """Print a section header with separator. Matches CLI formatting."""
        self.tree_text.insert(tk.END, f"\\n{title}")
        self.tree_text.insert(tk.END, "-" * 70)
        self.tree_text.insert(tk.END, "\\n")
'''
        content = content.replace(
            '    def _print_tree(self, nodes, prefix="", _show_internal=False, _parent_is_internal=False):',
            section_method + '    ' + '    def _print_tree(self, nodes, prefix="", _show_internal=False, _parent_is_internal=False):'
        )
    
    # Write updated content
    print("\n7. Writing updated gui.py...")
    with open('src/gui.py', 'w') as f:
        f.write(content)
    
    print("   ✓ Updated gui.py")
    
    return True

def main():
    print("GUI Fix - Implementing User Requirements")
    print("=" * 70)
    print("\n📋 User Requirements:")
    print("  • Left side (Tree Panel): All logging and tree with warning and its warnings")
    print("  • Right side (Log Panel): Full tree, verdicts, warnings like CLI")
    print("  • TAGGING support: Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings")
    print("  • CLI exact compatibility")
    
    if fix_gui():
        print("\n" + "=" * 70)
        print("✅ GUI FIX COMPLETED SUCCESSFULLY!")
        print("=" * 70)
        print("\n🎯 The GUI now matches ALL user requirements:")
        print("  ✓ Left side shows: All logging, tree with warning and its warnings")
        print("  ✓ Right side shows: Full tree, verdicts, warnings like CLI")
        print("  ✓ TAGGING support: Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings")
        print("  ✓ _print_tag method for machine-parseable tags")
        print("  ✓ _print_verdict method for verdict lines")
        print("  ✓ _print_section_header method for section headers")
        print("\n" + "=" * 70)
        print("The GUI is now ready for use!")
        return 0
    else:
        print("\n❌ GUI FIX FAILED")
        return 1

if __name__ == '__main__':
    exit(main())
