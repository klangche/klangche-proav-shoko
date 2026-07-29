#!/usr/bin/env python3
"""
Analyze the current GUI structure and identify what's missing
compared to CLI requirements.
"""

def analyze_gui():
    print("=" * 70)
    print("GUI ANALYSIS - Missing CLI Features")
    print("=" * 70)
    
    try:
        with open('src/gui.py', 'r') as f:
            content = f.read()
    except FileNotFoundError:
        print("ERROR: src/gui.py not found!")
        return False
    
    # Check for CLI features that should be in GUI
    cli_features = [
        # TAGGING support
        ("Tagg xxx.tree", "TAGGING section for tree"),
        ("Tagg xxx.score", "TAGGING section for score"),
        ("Tag xxx.warnings", "TAGGING section for warnings"),
        
        # Headers and sections
        ("Full USB & Display Tree", "Full tree header"),
        ("PER PORT", "PER PORT section header"),
        ("VERDICT", "VERDICT section"),
        
        # CLI specific methods
        ("def _print_tag(self, tag: str)", "_print_tag method"),
        ("def _print_verdict", "_print_verdict method"),
        ("def _print_section_header", "_print_section_header method"),
        ("def _print_stability_port", "_print_stability_port method"),
    ]
    
    print("\nChecking for CLI features in GUI:")
    print("-" * 70)
    
    found = []
    missing = []
    
    for pattern, description in cli_features:
        if pattern in content:
            print(f"  ✓ {description}")
            found.append(pattern)
        else:
            print(f"  ✗ {description}")
            missing.append(pattern)
    
    print("\n" + "-" * 70)
    print(f"Summary: {len(found)} found, {len(missing)} missing")
    
    if missing:
        print("\nMISSING CLI FEATURES:")
        for pattern, description in cli_features:
            if pattern in missing:
                print(f"  - {description}: {pattern}")
    
    print("\n" + "=" * 70)
    print("CLI vs GUI Structure Analysis")
    print("=" * 70)
    
    # Check _update_tree_display method
    method_start = content.find('    def _update_tree_display(self, usb_tree, hops_data, stability, displays):')
    
    if method_start != -1:
        print("\nFound _update_tree_display method")
        
        method_end = content.find('\n    def ', method_start + 100)
        if method_end == -1:
            method_end = len(content)
        
        method_content = content[method_start:method_end]
        
        # Check what's in the method
        checks = [
            ('"Full USB & Display Tree"', "Full USB & Display Tree header"),
            ('Tagg xxx.tree', "TAGGING tree section"),
            ('Tagg xxx.score', "TAGGING score section"),
            ('Tag xxx.warnings', "TAGGING warnings section"),
            ('"PER PORT"', "PER PORT section"),
            ('VERDICT', "VERDICT section"),
        ]
        
        print("\nContents of _update_tree_display method:")
        for pattern, description in checks:
            if pattern in method_content:
                print(f"  ✓ {description}")
            else:
                print(f"  ✗ {description}")
    
    # Check current layout
    print("\n" + "=" * 70)
    print("Current GUI Layout")
    print("=" * 70)
    
    if 'middle_container = tk.Frame(main_frame, bg=self.colors[\'bg\'])' in content:
        print("\n✓ Has middle_container with resizable panels")
    
    if 'self.left_frame = tk.Frame(middle_container' in content:
        print("✓ Has left_frame (USB tree)")
    
    if 'self.right_frame = tk.Frame(middle_container' in content:
        print("✓ Has right_frame (Live Log)")
    
    print("\n" + "=" * 70)
    print("RECOMMENDATIONS")
    print("=" * 70)
    
    print("\nTo match CLI format, GUI needs:")
    print("1. TAGGING support (Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings)")
    print("2. 'Full USB & Display Tree' header")
    print("3. VERDICT sections")
    print("4. PER PORT header and sections")
    print("5. _print_tag method for machine-parseable tags")
    print("6. Better layout integration with left/right panels")
    
    print("\n" + "=" * 70)
    print("NEXT STEPS")
    print("=" * 70)
    print("\nRun a fix script to add these features:")
    print("  python3 scripts/fix_gui_final.py")
    
    return True

if __name__ == '__main__':
    analyze_gui()
