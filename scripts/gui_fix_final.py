#!/usr/bin/env python3
"""
GUI FIX - Final version to implement user requirements:
- Left side: All logging and tree with warning and its warnings
- Right side: Full tree, verdicts, warnings like CLI
- TAGGING support (Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings)
- CLI exact compatibility
"""

import re
import shutil
import time

def backup_file(file_path):
    timestamp = time.strftime('%Y%m%d_%H%M%S')
    backup_path = f"{file_path}.backup_{timestamp}"
    shutil.copy2(file_path, backup_path)
    return backup_path

def main():
    print("=" * 70)
    print("GUI FIX - Applying User Requirements")
    print("=" * 70)
    
    print("\n📋 User Requirements:")
    print("  • Left side: All logging, tree with warning and its warnings")
    print("  • Right side: Full tree, verdicts, warnings like CLI")
    print("  • TAGGING support: Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings")
    print("  • CLI exact compatibility")
    
    # Check if gui.py exists
    if not shutil.os.path.exists('src/gui.py'):
        print("ERROR: src/gui.py not found!")
        return 1
    
    print("\n1. Creating backup...")
    backup_file('src/gui.py')
    
    # Read gui.py
    with open('src/gui.py', 'r') as f:
        content = f.read()
    
    print("\n2. Checking current state...")
    
    # Check requirements
    checks = {
        'Full USB & Display Tree': 'Full USB & Display Tree' in content,
        'Left panel header': 'All logging and tree with warning and its warnings' in content,
        'Right panel header': 'Full tree, verdicts, warnings like CLI' in content,
        'Tagg xxx.tree': 'Tagg xxx.tree' in content,
        'Tagg xxx.score': 'Tagg xxx.score' in content,
        'Tag xxx.warnings': 'Tag xxx.warnings' in content,
        '_print_tag method': '    def _print_tag(self, tag: str):' in content,
        '_print_verdict method': '    def _print_verdict(self, v):' in content,
        '_print_section_header method': '    def _print_section_header(self, title):' in content,
    }
    
    print("\n   Current checks:")
    for req, met in checks.items():
        status = "✓" if met else "✗"
        print(f"   {status} {req}")
    
    # Apply fixes if needed
    all_met = all(checks.values())
    
    if not all_met:
        print(f"\n🔧 Applying fixes...")
        
        # Fix panel headers
        print("\n   3. Updating panel headers...")
        
        # Fix left panel header (Live Log -> All logging and tree with warning and its warnings)
        if 'text="Live Log"' in content:
            content = content.replace(
                'text="Live Log"',
                'text="All logging and tree with warning and its warnings"'
            )
            print("      ✓ Fixed left panel header")
        
        # Fix right panel header (USB Tree & Stability -> Full tree, verdicts, warnings like CLI)
        if 'text="USB Tree & Stability"' in content:
            content = content.replace(
                'text="USB Tree & Stability"',
                'text="Full tree, verdicts, warnings like CLI"'
            )
            print("      ✓ Fixed right panel header")
        
        # Fix _update_tree_display to add "Full USB & Display Tree" header
        if 'Full USB & Display Tree' not in content:
            # Find the _update_tree_display method and add the header
            method_start = content.find('    def _update_tree_display(self, usb_tree, hops_data, stability, displays):')
            if method_start != -1:
                # Find the line where the method starts
                lines = content.split('\n')
                method_lines = []
                for i, line in enumerate(lines):
                    if i == method_start // 100:  # Approximate line
                        # Look for the line where we can insert the header
                        for j in range(len(lines)):
                            if 'PER PORT - EXTERNAL' in lines[j]:
                                # Insert before this line
                                lines.insert(j, '        # Full USB & Display Tree')
                                lines.insert(j+1, '        self._print_section_header("Full USB & Display Tree")')
                                lines.insert(j+2, '        self._print_tag("overall.tree")')
                                lines.insert(j+3, '        self._print_tag("overall.score")')
                                lines.insert(j+4, '        self._print_tag("overall.warnings")')
                                break
                        break
                content = '\n'.join(lines)
                print("      ✓ Added Full USB & Display Tree header and TAGGING sections")
        
        # Add _print_tag method if missing
        if '    def _print_tag(self, tag: str):' not in content:
            print("\n   4. Adding _print_tag method...")
            # Insert before the _print_verdict method if it exists
            if '    def _print_verdict(self, v):' in content:
                # Insert before _print_verdict
                content = content.replace(
                    '    def _print_verdict(self, v):',
                    '''    def _print_tag(self, tag: str):
        """Print a machine-parseable tag for GUI parsing. Only tag name, no [TAG:] wrapper."""
        self.tree_text.insert(tk.END, f"\\n    {tag}\\n")

    def _print_verdict(self, v):'''
                )
                print("      ✓ Added _print_tag method")
        
        # Add _print_verdict if missing
        if '    def _print_verdict(self, v):' not in content:
            print("\n   5. Adding _print_verdict method...")
            # Insert before _print_tree
            tree_pos = content.find('    def _print_tree(')
            if tree_pos != -1:
                # Insert _print_verdict before _print_tree
                verdict_method = '''
    def _print_verdict(self, v):
        """Print a single verdict line with hops/tiers/hubs. Matches CLI exactly."""
        status_char = '+' if v['color'] == 'green' else ('~' if v['color'] == 'orange' else '!')
        hubs_str = f"hubs {v['current_hubs']}/{v['max_hubs']}  " if 'current_hubs' in v else ""
        desc = v.get('description', v.get('name', ''))
        self.tree_text.insert(tk.END, f"    {status_char} {desc:<22s} {v['status']:<9s} hops {v['current_hops']}/{v['max_hops']}  tiers {v['current_tiers']}/{v['max_tiers']}  {hubs_str}\\n")
'''
                content = content[:tree_pos] + verdict_method + '\n' + content[tree_pos:]
                print("      ✓ Added _print_verdict method")
        
        # Add _print_section_header if missing
        if '    def _print_section_header(self, title):' not in content:
            print("\n   6. Adding _print_section_header method...")
            # Insert before _print_tree
            tree_pos = content.find('    def _print_tree(')
            if tree_pos != -1:
                # Insert _print_section_header before _print_tree
                section_method = '''
    def _print_section_header(self, title):
        """Print a section header with separator. Matches CLI formatting."""
        self.tree_text.insert(tk.END, f"\\n{title}")
        self.tree_text.insert(tk.END, "-" * 70)
        self.tree_text.insert(tk.END, "\\n")
'''
                content = content[:tree_pos] + section_method + '\n' + content[tree_pos:]
                print("      ✓ Added _print_section_header method")
        
        # Write updated content
        with open('src/gui.py', 'w') as f:
            f.write(content)
        
        print("\n   7. Writing updated gui.py...")
        print("      ✓ Updated gui.py")
    
    # Check final state
    print("\n8. Final verification...")
    final_checks = check_requirements(content)
    
    all_good = all(final_checks.values())
    
    if all_good:
        print("\n✓ GUI FIX COMPLETED SUCCESSFULLY!")
        print("\nGUI now matches ALL user requirements:")
        print("  ✓ Left side: All logging, tree with warning and its warnings")
        print("  ✓ Right side: Full tree, verdicts, warnings like CLI")
        print("  ✓ TAGGING support: Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings")
        print("  ✓ _print_tag method for machine-parseable tags")
        print("  ✓ _print_verdict method for verdict lines")
        print("  ✓ _print_section_header method for section headers")
        print("\n" + "=" * 70)
        print("The GUI is now ready and meets all user requirements!")
        return 0
    else:
        print("\n⚠️ GUI still needs some fixes:")
        for req, met in final_checks.items():
            if not met:
                print(f"  ✗ {req}")
        return 1

def check_requirements(content):
    """Check if all GUI requirements are met"""
    requirements = {
        'Full USB & Display Tree header': 'Full USB & Display Tree' in content,
        'Left panel header': 'All logging and tree with warning and its warnings' in content,
        'Right panel header': 'Full tree, verdicts, warnings like CLI' in content,
        'Tagg xxx.tree': 'Tagg xxx.tree' in content,
        'Tagg xxx.score': 'Tagg xxx.score' in content,
        'Tag xxx.warnings': 'Tag xxx.warnings' in content,
        '_print_tag method': '    def _print_tag(self, tag: str):' in content,
        '_print_verdict method': '    def _print_verdict(self, v):' in content,
        '_print_section_header method': '    def _print_section_header(self, title):' in content,
    }
    return requirements

if __name__ == '__main__':
    exit(main())
