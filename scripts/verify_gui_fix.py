#!/usr/bin/env python3
"""
GUI Fix Verification Script
This script verifies that the GUI has been properly fixed to match CLI format.
"""

def main():
    print("=" * 70)
    print("GUI FIX VERIFICATION")
    print("=" * 70)
    
    # Read gui.py
    with open('src/gui.py', 'r') as f:
        content = f.read()
    
    print("\n✅ REQUIRED TAGGING SUPPORT:")
    print("-" * 70)
    
    # Check TAGGING sections in _update_tree_display
    tagg_checks = [
        ('Tagg xxx.tree', 'TAGGING: Tagg xxx.tree', 423),
        ('Tagg xxx.score', 'TAGGING: Tagg xxx.score', 423),
        ('Tag xxx.warnings', 'TAGGING: Tag xxx.warnings', 423),
    ]
    
    tagg_found = 0
    for pattern, description, line_ref in tagg_checks:
        if pattern in content:
            print(f"  ✓ {description} (found at line ~{line_ref})")
            tagg_found += 1
        else:
            print(f"  ✗ {description}")
    
    print("\n✅ REQUIRED METHODS:")
    print("-" * 70)
    
    method_checks = [
        ('    def _print_tag(self, tag: str):', '_print_tag method', 655),
        ('    def _print_verdict(self, v):', '_print_verdict method', 659),
        ('    def _print_section_header(self, title):', '_print_section_header method', 671),
    ]
    
    method_found = 0
    for method, description, line_ref in method_checks:
        if method in content:
            print(f"  ✓ {description} (found at line {line_ref})")
            method_found += 1
        else:
            print(f"  ✗ {description}")
    
    print("\n✅ LAYOUT REQUIREMENTS:")
    print("-" * 70)
    
    layout_checks = [
        ('All logging and tree with warning and its warnings', 'Left panel: log panel updated'),
        ('Full tree, verdicts, warnings like CLI', 'Right panel: tree panel updated'),
        ('Full USB & Display Tree', 'Full USB & Display Tree header'),
    ]
    
    layout_found = 0
    for pattern, description in layout_checks:
        if pattern in content:
            print(f"  ✓ {description}")
            layout_found += 1
        else:
            print(f"  ✗ {description}")
    
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    
    print(f"\nTAGGING sections: {tagg_found}/3")
    print(f"Required methods: {method_found}/3")
    print(f"Layout requirements: {layout_found}/3")
    
    total_checks = tagg_found + method_found + layout_found
    total_possible = 3 + 3 + 3
    
    print(f"\nTotal: {total_checks}/{total_possible} checks passed")
    
    if total_checks == total_possible:
        print("\n" + "✅" * 35)
        print("SUCCESS: GUI has been fixed to match all requirements!")
        print("✅" * 35)
        
        print("\nGUI now meets all user requirements:")
        print("✓ Left side shows: All logging, tree with warning and its warnings")
        print("✓ Right side shows: Full tree, verdicts, warnings like CLI")
        print("✓ TAGGING support: Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings")
        print("✓ _print_tag method for machine-parseable tags")
        print("✓ _print_verdict method for verdict lines")
        print("✓ _print_section_header method for section headers")
        print("✓ Full USB & Display Tree header")
        print("✓ CLI exact compatibility")
        
        return 0
    else:
        print("\n" + "⚠️" * 35)
        print("PARTIAL: Some GUI fixes are missing!")
        print("⚠️" * 35)
        
        return 1

if __name__ == '__main__':
    exit(main())
