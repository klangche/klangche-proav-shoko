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
        lines = f.readlines()
        content = ''.join(lines)
    
    fixes_applied = []
    
    # Fix 1: Replace all text_color with fg
    if 'text_color' in content:
        print("❌ FAIL: text_color found - will cause Tkinter errors")
        
        # Find all text_color occurrences
        text_color_matches = []
        for i, line in enumerate(lines):
            if 'text_color' in line:
                text_color_matches.append((i, line))
        
        # Show the problematic lines
        print(f"   Found {len(text_color_matches)} instances of text_color:")
        for line_num, line in text_color_matches:
            print(f"   Line {line_num + 1}: {line.strip()}")
        
        print("   🔧 FIXING: Replacing text_color with fg...")
        
        # Replace all text_color with fg
        new_lines = []
        replaced_count = 0
        for line in lines:
            new_line = line.replace('text_color', 'fg')
            if new_line != line:
                replaced_count += 1
            new_lines.append(new_line)
        
        lines = new_lines
        fixes_applied.append("text_color → fg replacement")
        
        # Write back
        with open('src/gui.py', 'w') as f:
            f.writelines(lines)
        
        print(f"   ✓ FIXED: Replaced {replaced_count} instances of text_color")
    else:
        print("✓ PASS: No text_color found")
    
    # Fix 2: Remove temp_log.txt file saving
    if 'temp_log.txt' in content:
        print("❌ FAIL: temp_log.txt file saving found")
        print("   GUI will crash with: [Errno 13] Permission denied: 'temp_log.txt'")
        
        # Find temp_log.txt usage
        temp_log_matches = []
        for i, line in enumerate(lines):
            if 'temp_log.txt' in line:
                temp_log_matches.append((i, line))
        
        print(f"   Found temp_log.txt usage at:")
        for line_num, line in temp_log_matches:
            print(f"   Line {line_num + 1}: {line.strip()}")
        
        print("   🔧 FIXING: Removing temp_log.txt file saving...")
        
        # Remove the problematic section
        new_lines = []
        in_temp_section = False
        for line in lines:
            if 'temp_log.txt' in line:
                in_temp_section = True
            
            if in_temp_section:
                # Skip lines with temp_log.txt and the following few lines
                # Continue skipping until we leave the problematic block
                if '            # 3.' in line or '            print("\\n[+] Scanning displays..."' in line:
                    in_temp_section = False
                    new_lines.append(line)
                continue
            else:
                new_lines.append(line)
        
        lines = new_lines
        fixes_applied.append("temp_log.txt file saving removed")
        
        # Write back
        with open('src/gui.py', 'w') as f:
            f.writelines(lines)
        
        print("   ✓ FIXED: temp_log.txt file saving removed")
    else:
        print("✓ PASS: No temp_log.txt file saving")
    
    # Fix 3: Update _update_tree_display method
    # Look for the method
    in_method = False
    method_start = None
    method_end = None
    indent_level = None
    
    for i, line in enumerate(lines):
        if 'def _update_tree_display' in line:
            method_start = i
            indent_level = len(line) - len(line.lstrip())
            in_method = True
        elif in_method:
            if line.strip() and len(line) - len(line.lstrip()) <= indent_level and 'def ' in line:
                method_end = i
                break
    
    if method_start is not None:
        # Extract the method
        method_lines = lines[method_start:method_end if method_end else len(lines)]
        method_text = ''.join(method_lines)
        
        # Check if the method needs updating
        has_verdict = 'VERDICT' in method_text
        has_full_tree = 'Full USB & Display Tree' in method_text
        has_external_section = 'EXTERNAL' in method_text
        has_internal_section = 'INTERNAL' in method_text
        
        if not (has_verdict and has_full_tree):
            print("❌ FAIL: _update_tree_display method needs updating")
            print("   Current method structure:")
            print(f"   - Has VERDICT: {has_verdict}")
            print(f"   - Has Full USB & Display Tree: {has_full_tree}")
            
            print("   🔧 FIXING: Rewriting _update_tree_display method...")
            
            # Read the current working copy
            os.chdir(os.path.dirname(os.path.abspath(__file__)))
            
            # Create a new method
            new_method = '''    def _update_tree_display(self, usb_tree, hops_data, stability, displays):
        """Update the tree display to match CLI exactly."""
        self.tree_text.delete('0.0', tk.END)

        # FULL USB & Display Tree
        self.tree_text.insert(tk.END, "Full USB & Display Tree\\n")
        if usb_tree:
            root_node = usb_tree[0]

            # Add displays directly into tree (not under a "Displays" parent)
            if displays:
                for i, d in enumerate(displays):
                    prim = " (Primary)" if d.get('is_primary', False) else ""
                    int_disp = d.get('is_internal', False)
                    root_node.setdefault('children', []).append({
                        'model': f"{d['resolution']}  {d['name']}{prim}",
                        'name': d['name'], 'children': [], 'hops': 1,
                        'is_hub': False, 'is_internal': int_disp, 'is_display': True, 'port': 0
                    })

            self._print_tree(usb_tree, "", _show_internal=True)
        else:
            self.tree_text.insert(tk.END, "No USB devices found.\\n")
        self.tree_text.insert(tk.END, "\\n")

        # PER PORT - EXTERNAL
        root_orig = usb_tree[0] if usb_tree else {}
        orig_children = list(root_orig.get('children', []))
        
        # Print main port structure (first level)
        for idx, child in enumerate(orig_children):
            if child.get('is_display'):
                continue
            if not child.get('is_internal', False):
                
                # Print port header
                port_info = next((p for p in stability.get('ports', []) if p.get('id') == idx + 1), None)
                if port_info:
                    label = f"  {port_info['label']} ({len(port_info['devices'])} endpoint{'s' if len(port_info['devices']) != 1 else ''}, hops={port_info['max_hops']}, tiers={port_info['max_tiers']}, hubs={port_info.get('external_hubs', 0)})"
                else:
                    label = f"  {child.get('model', 'Port')}"
                self.tree_text.insert(tk.END, f"{label}\\n")
                
                # Print port tree
                self._print_port_tree(child)
                self.tree_text.insert(tk.END, "\\n")
        
        # VERDICT section for EXTERNAL
        self.tree_text.insert(tk.END, "VERDICT\\n")
        for v in stability.get('verdicts', []):
            status_char = '+' if v['color'] == 'green' else ('~' if v['color'] == 'orange' else '!')
            hops_str = f"hops {v['current_hops']}/{v['max_hops']}  "
            tiers_str = f"tiers {v['current_tiers']}/{v['max_tiers']}  "
            if 'current_hubs' in v:
                hubs_str = f"hubs {v['current_hubs']}/{v['max_hubs']}"
                self.tree_text.insert(tk.END, f"  {status_char} {v['description']:<25s} {v['status']:<8s} {hops_str}{tiers_str}{hubs_str}\\n")
            else:
                self.tree_text.insert(tk.END, f"  {status_char} {v['description']:<25s} {v['status']:<8s} {hops_str}{tiers_str}\\n")
        
        # Print separator
        self.tree_text.insert(tk.END, "\\n  " + "- " * 35 + "\\n")

        # PER PORT - INTERNAL
        self.tree_text.insert(tk.END, "=" * 31 + "INTERNAL" + "=" * 31)
        self.tree_text.insert(tk.END, "\\n")
        
        # Print internal ports
        for idx, child in enumerate(orig_children):
            if child.get('is_display'):
                continue
            if child.get('is_internal', False):
                
                # Print port header
                port_info = next((p for p in stability.get('ports', []) if p.get('id') == idx + 1), None)
                if port_info:
                    label = f"  {port_info['label']} ({len(port_info['devices'])} endpoint{'s' if len(port_info['devices']) != 1 else ''}, hops={port_info['max_hops']}, tiers={port_info['max_tiers']}, hubs={port_info.get('external_hubs', 0)})    (internal)"
                else:
                    label = f"  {child.get('model', 'Port')}    (internal)"
                self.tree_text.insert(tk.END, f"{label}\\n")
                
                # Print port tree
                self._print_port_tree(child)
                self.tree_text.insert(tk.END, "\\n")
        
        # VERDICT section for INTERNAL
        internal_verdicts = [v for v in stability.get('verdicts', []) if v['current_hops'] < v['max_hops'] or v['current_tiers'] < v['max_tiers'] or v['current_hubs'] < v['max_hubs']]
        if internal_verdicts:
            self.tree_text.insert(tk.END, "VERDICT\\n")
            for v in internal_verdicts:
                status_char = '+' if v['color'] == 'green' else ('~' if v['color'] == 'orange' else '!')
                hops_str = f"hops {v['current_hops']}/{v['max_hops']}  "
                tiers_str = f"tiers {v['current_tiers']}/{v['max_tiers']}  "
                if 'current_hubs' in v:
                    hubs_str = f"hubs {v['current_hubs']}/{v['max_hubs']}"
                    self.tree_text.insert(tk.END, f"  {status_char} {v['description']:<25s} {v['status']:<8s} {hops_str}{tiers_str}{hubs_str}\\n")
                else:
                    self.tree_text.insert(tk.END, f"  {status_char} {v['description']:<25s} {v['status']:<8s} {hops_str}{tiers_str}\\n")
        '''
            
            # Replace the method in the lines
            new_lines = lines[:method_start] + [line + '\\n' for line in new_method.split('\\n')] + lines[method_end:]
            
            # Write back
            with open('src/gui.py', 'w') as f:
                f.writelines(new_lines)
            
            fixes_applied.append("_update_tree_display method updated")
            print("   ✓ FIXED: _update_tree_display method rewritten")
        else:
            print("✓ PASS: _update_tree_display method appears to be updated correctly")
    
    # Fix 4: Create CRLF version
    print("\n=== Converting to CRLF line endings ===")
    try:
        with open('src/gui_crlf.py', 'w', newline='\\r\\n') as f:
            with open('src/gui.py', 'r') as src:
                f.write(src.read())
        
        crlf_size = os.path.getsize('src/gui_crlf.py')
        original_size = os.path.getsize('src/gui.py')
        
        print(f"✓ PASS: Created src/gui_crlf.py ({crlf_size} bytes)")
        
        if crlf_size != original_size:
            print(f"   Note: File size changed from {original_size} to {crlf_size} bytes due to line ending conversion")
    except Exception as e:
        print(f"❌ FAIL: Could not create CRLF version: {e}")
    
    # Verify fixes
    print("\n=== Verification ===")
    
    with open('src/gui.py', 'r') as f:
        final_content = f.read()
    
    # Check 1: No text_color
    if 'text_color' not in final_content:
        print("✓ PASS: No text_color found in final version")
    else:
        print("❌ FAIL: text_color still exists!")
    
    # Check 2: No temp_log.txt
    if 'temp_log.txt' not in final_content:
        print("✓ PASS: No temp_log.txt found in final version")
    else:
        print("❌ FAIL: temp_log.txt still exists!")
    
    # Check 3: VERDICT and tree structure
    if 'def _update_tree_display' in final_content:
        method_match = re.search(r'def _update_tree_display.*?(?=\\n    def |\\n\\n|$)', final_content, re.DOTALL)
        if method_match:
            method_text = method_match.group(0)
            if 'VERDICT' in method_text and 'Full USB & Display Tree' in method_text:
                print("✓ PASS: _update_tree_display has VERDICT and tree structure")
            else:
                print("❌ FAIL: _update_tree_display missing VERDICT or tree structure")
    
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
