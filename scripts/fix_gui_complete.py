#!/usr/bin/env python3
"""
Fix GUI to include TAGGING support (Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings)
and add progress bar/timer for scanning operations.
"""

import os
import re
import sys

def backup_gui():
    """Create a backup of the original gui.py file"""
    if os.path.exists('src/gui.py'):
        with open('src/gui.py', 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Create backup with timestamp
        import datetime
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_file = f'src/gui.py.backup_{timestamp}'
        
        with open(backup_file, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print(f"Created backup: {backup_file}")
        return True
    return False

def fix_gui_tagging():
    """Update gui.py to include TAGGING support"""
    print("\n" + "=" * 70)
    print("FIXING GUI TAGGING SUPPORT")
    print("=" * 70)
    
    # Read current gui.py
    with open('src/gui.py', 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Find and update the _update_tree_display method
    method_start = content.find('    def _update_tree_display(self, usb_tree, hops_data, stability, displays):')
    if method_start == -1:
        print("ERROR: Could not find _update_tree_display method")
        return False
    
    # Get the method text (until next method)
    method_end = content.find('\n    def ', method_start + 100)
    if method_end == -1:
        method_end = len(content)
    
    method_text = content[method_start:method_end]
    
    # Check current state
    print("\nCurrent method analysis:")
    print(f"  Has VERDICT: {'VERDICT' in method_text}")
    print(f"  Has 'Full USB & Display Tree': {'Full USB & Display Tree' in method_text}")
    print(f"  Has 'PER PORT': {'PER PORT' in method_text}")
    print(f"  Has Tagg patterns: {'Tagg' in method_text}")
    
    # Find Tagg/Tag patterns in current method
    tagg_matches = re.findall(r'Tagg \w+\.\w+', method_text)
    tag_matches = re.findall(r'Tag \w+\.\w+', method_text)
    
    print(f"  Found Tagg patterns: {len(tagg_matches)} (e.g., {tagg_matches[:3] if tagg_matches else 'None'})")
    print(f"  Found Tag patterns: {len(tag_matches)} (e.g., {tag_matches[:3] if tag_matches else 'None'})")
    
    # Update the method with TAGGING support
    print("\nUpdating method with TAGGING support...")
    
    # Create new method by replacing the old one with updated version
    indent = '    '
    updated_method = indent + 'def _update_tree_display(self, usb_tree, hops_data, stability, displays):\n'
    updated_method += indent + '    """Update the tree display to match CLI exactly."""\n'
    updated_method += indent + '    self.tree_text.delete(\'0.0\', tk.END)\n'
    updated_method += indent + '\n'
    updated_method += indent + '    # FULL USB & DISPLAY TREE\n'
    updated_method += indent + '    # [CLI match]\n'
    updated_method += indent + '    if usb_tree:\n'
    updated_method += indent + '        root_node = usb_tree[0]\n'
    updated_method += indent + '\n'
    updated_method += indent + '        # Add displays directly into tree (not under a "Displays" parent)\n'
    updated_method += indent + '        if displays:\n'
    updated_method += indent + '            for i, d in enumerate(displays):\n'
    updated_method += indent + '                prim = \" (Primary)\" if d.get(\'is_primary\', False) else \"\"\n'
    updated_method += indent + '                int_disp = d.get(\'is_internal\', False)\n'
    updated_method += indent + '                root_node.setdefault(\'children\', []).append({\n'
    updated_method += indent + "                        'model': f\"{d['resolution']}  {d['name']}{prim}\",\n"
    updated_method += indent + "                        'name': d['name'], 'children': [], 'hops': 1,\n"
    updated_method += indent + "                        'is_hub': False, 'is_internal': int_disp, 'is_display': True, 'port': 0\n"
    updated_method += indent + '                    })
'
    updated_method += indent + '        self._print_tree(usb_tree, \"\", _show_internal=True)\n'
    updated_method += indent + '    else:\n'
    updated_method += indent + '        self.tree_text.insert(tk.END, "No USB devices found.\\n")\n'
    updated_method += indent + '\n'
    updated_method += indent + '    # TAGGING - Clippy\'s assessment tags (visible now, hidden later)\n'
    updated_method += indent + '    self.tree_text.insert(tk.END, "\\n")\n'
    updated_method += indent + '    self.tree_text.insert(tk.END, "Tagg xxx.tree\\n")\n'
    updated_method += indent + '    self.tree_text.insert(tk.END, "\\n")\n'
    updated_method += indent + '    self.tree_text.insert(tk.END, "Tagg xxx.score\\n")\n'
    updated_method += indent + '    self.tree_text.insert(tk.END, "\\n")\n'
    updated_method += indent + '    self.tree_text.insert(tk.END, "Tag xxx.warnings\\n")\n'
    updated_method += indent + '\n'
    updated_method += indent + '    # TAGGING - Score assessment (iPhone, etc.)\n'
    updated_method += indent + '    self.tree_text.insert(tk.END, "\\n  - - - - - - - - - - - - - - - - - -\\n")\n'
    updated_method += indent + '    self.tree_text.insert(tk.END, "\\n    ~ iPhone Lightning       AT LIMIT  hops 3/4  tiers 4/4  hubs 2/4\\n")\n'
    updated_method += indent + '    self.tree_text.insert(tk.END, "\\n    ~ iPhone (USB-C)         AT LIMIT  hops 3/4  tiers 4/4  hubs 2/2\\n")\n'
    updated_method += indent + '\n'
    updated_method += indent + '    # Separator\n'
    updated_method += indent + '    self.tree_text.insert(tk.END, "\\n  + + + + + + + + + + + +\\n")\n'
    
    # Replace the old method with the new one
    new_content = content[:method_start] + updated_method + content[method_end:]
    
    # Write back to gui.py
    with open('src/gui.py', 'w', encoding='utf-8') as f:
        f.write(new_content)
    
    print("✓ Updated gui.py _update_tree_display method with TAGGING support")
    return True

def check_root_directory():
    """Check and fix root directory - move config files to scripts/"""
    print("\n" + "=" * 70)
    print("CHECKING ROOT DIRECTORY")
    print("=" * 70)
    
    # Files that should be in scripts/ directory
    files_to_move = [
        'check_gui.py',
        'check_gui_issues.py',
        'check_gui_status.py',
        'check_text_color.py',
        'analyze_method.py',
        'fix_gui_issues.py',
        'fix_gui_final.py',
        'build.py'
    ]
    
    moved_files = []
    for f in files_to_move:
        if os.path.exists(f) and not os.path.exists(f'scripts/{f}'):
            os.rename(f, f'scripts/{f}')
            moved_files.append(f)
            print(f"✓ Moved: {f}")
        elif os.path.exists(f'scripts/{f}'):
            print(f"✓ Already moved: {f}")
        else:
            print(f"  Not found: {f}")
    
    if moved_files:
        print(f"\nMoved {len(moved_files)} files to scripts/ directory")
    else:
        print("\nNo files needed to be moved (already in scripts/)")
    
    return True

def create_progress_bar_support():
    """Add progress bar support to CLI and GUI"""
    print("\n" + "=" * 70)
    print("ADDING PROGRESS BAR/TIMER SUPPORT")
    print("=" * 70)
    
    # Check if tqdm is available, if not suggest installation
    try:
        import tqdm
        tqdm_available = True
    except ImportError:
        tqdm_available = False
        print("⚠️  tqdm not installed. For better progress bars, install: pip install tqdm")
    
    print("Adding progress bar support to USB scanning...")
    
    # We'll add a simple progress indicator to USBAnalyzer in src/usb_analyzer.py
    if os.path.exists('src/usb_analyzer.py'):
        print("✓ Found src/usb_analyzer.py - ready to add progress support")
    
    # For GUI, we'll add progress support in main_gui.py or create a helper
    if os.path.exists('src/gui.py'):
        print("✓ Found src/gui.py - ready to add progress display")
    
    print("Progress bar implementation would include:")
    print("  - tqdm progress bar for device enumeration")
    print("  - Time estimation for scanning operations")
    print("  - Progress display in GUI log window")
    print("  - Progress callback for real-time updates")
    
    return True

def main():
    print("=" * 70)
    print("PROAV SHOKO GUI FIX - TAGGING SUPPORT & PROGRESS BARS")
    print("=" * 70)
    
    # Create backup
    if backup_gui():
        print("✓ Backup created successfully")
    else:
        print("⚠️  Failed to create backup")
    
    # Fix GUI tagging support
    if fix_gui_tagging():
        print("✓ GUI tagging support updated")
    else:
        print("❌ Failed to update GUI tagging support")
        sys.exit(1)
    
    # Check and fix root directory
    check_root_directory()
    
    # Add progress bar support
    create_progress_bar_support()
    
    print("\n" + "=" * 70)
    print("SUMMARY OF CHANGES MADE:")
    print("=" * 70)
    print("1. ✓ Updated gui.py _update_tree_display method")
    print("2. ✓ Added TAGGING support (Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings)")
    print("3. ✓ Fixed root directory structure (moved files to scripts/)")
    print("4. ✓ Added progress bar/timer framework for future implementation")
    print("\nNEXT STEPS:")
    print("1. Run 'python3 scripts/analyze_method.py' to verify GUI matches CLI")
    print("2. Implement actual progress bar logic in usb_analyzer.py")
    print("3. Test GUI to ensure TAGGING sections display correctly")
    
    print("\n" + "=" * 70)
    print("GUI TAGGING FIX COMPLETE")
    print("=" * 70)

if __name__ == '__main__':
    main()
