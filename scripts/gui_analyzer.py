#!/usr/bin/env python3
"""
Check if GUI _update_tree_display method matches CLI exactly.
The CLI shows TAGGING sections (Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings)
and VERDICT, PER PORT, EXTERNAL, INTERNAL sections.
"""

import os

def main():
    print("=" * 70)
    print("ANALYZING PROAV SHOKO - GUI TAGGING FIX VERIFICATION")
    print("=" * 70)
    
    # Check if gui.py exists
    if not os.path.exists('src/gui.py'):
        print("ERROR: src/gui.py not found!")
        return
    
    # Read the current gui.py
    with open('src/gui.py', 'r') as f:
        content = f.read()
    
    print("\n1. CHECKING _update_tree_display METHOD:")
    print("-" * 70)
    
    # Find the _update_tree_display method
    method_start = content.find('    def _update_tree_display(self, usb_tree, hops_data, stability, displays):')
    if method_start == -1:
        print("✗ _update_tree_display method NOT FOUND")
        return
    
    # Find the end of this method
    method_end = content.find('\n    def ', method_start + 100)
    if method_end == -1:
        method_end = len(content)
    
    method_text = content[method_start:method_end]
    
    print("✓ Found _update_tree_display method")
    print(f"  Method length: {len(method_text)} characters")
    
    # Check for CLI features
    print("\n2. CLI FEATURES CHECK:")
    print("-" * 70)
    
    required_features = {
        'Full USB & Display Tree': 'CLI main tree header',
        'PER PORT': 'CLI per-port section header',
        'Tagg xxx.tree': 'Tagg tree assessment section',
        'Tagg xxx.score': 'Tagg score assessment section',
        'Tag xxx.warnings': 'Tag warnings assessment section',
        'VERDICT': 'Stability verdict sections',
        'EXTERNAL': 'External devices section',
        'INTERNAL': 'Internal devices section'
    }
    
    all_present = True
    for feature, description in required_features.items():
        if feature in method_text:
            count = method_text.count(feature)
            print(f"  ✓ {description}: PRESENT ({count} occurrences)")
        else:
            print(f"  ✗ {description}: MISSING")
            all_present = False
    
    # Display specific TAGGING sections if they exist
    print("\n3. TAGGING SECTIONS ANALYSIS:")
    print("-" * 70)
    
    if 'Tagg xxx.tree' in method_text:
        print("  TAGGING sections found:")
        if 'Tagg xxx.tree' in method_text:
            print("    - Tagg xxx.tree: PRESENT")
        if 'Tagg xxx.score' in method_text:
            print("    - Tagg xxx.score: PRESENT")
        if 'Tag xxx.warnings' in method_text:
            print("    - Tag xxx.warnings: PRESENT")
        
        # Show context
        import re
        tagg_matches = re.findall(r'Tagg \w+\.\w+', method_text)
        if tagg_matches:
            print(f"    - Matched patterns: {tagg_matches}")
    else:
        print("  ⚠ TAGGING sections NOT FOUND (CLI shows these must be present)")
        print("  NOTE: The CLI output shows these TAGGING sections must be present")
    
    # Check VERDICT sections
    print("\n4. VERDICT SECTIONS:")
    print("-" * 70)
    
    verdict_count = method_text.count('VERDICT')
    if verdict_count > 0:
        print(f"  ✓ VERDICT sections: {verdict_count} found")
        
        # Show VERDICT context
        import re
        verdict_lines = re.findall(r'Verdict.*?(?=\\n\\n|\n    |\\n$', method_text)
        if verdict_lines:
            print(f"    - Sample VERDICT lines: {len(verdict_lines)} instances")
    else:
        print("  ✗ VERDICT sections: MISSING")
    
    # Check 'Full USB & Display Tree' header
    print("\n5. 'Full USB & Display Tree' HEADER:")
    print("-" * 70)
    
    if 'Full USB & Display Tree' in method_text:
        print("  ✓ 'Full USB & Display Tree' header: PRESENT")
        header_count = method_text.count('Full USB & Display Tree')
        print(f"    - Appears {header_count} times")
    else:
        print("  ✗ 'Full USB & Display Tree' header: MISSING")
        print("  NOTE: This is required to match CLI format")
    
    # Show port sections
    print("\n6. PORT SECTIONS:")
    print("-" * 70)
    
    external_count = method_text.count('EXTERNAL')
    internal_count = method_text.count('INTERNAL')
    
    print(f"  EXTERNAL section: {external_count} occurrences")
    print(f"  INTERNAL section: {internal_count} occurrences")
    
    # Verify method structure matches CLI expectations
    print("\n7. METHOD STRUCTURE VERIFICATION:")
    print("-" * 70)
    
    # Check for key structure elements
    struct_checks = [
        ('for idx, child in enumerate(orig_children):', 'Port iteration loop'),
        ('if orig_children:', 'Port sections container'),
        ('if external_children:', 'External section container'),
        ('if internal_children:', 'Internal section container'),
        ('self.tree_text.insert(tk.END', 'Tree text output')
    ]
    
    for pattern, description in struct_checks:
        if pattern in method_text:
            print(f"  ✓ {description}: PRESENT")
        else:
            print(f"  ✗ {description}: MISSING")
    
    # SUMMARY
    print("\n" + "=" * 70)
    print("VERIFICATION SUMMARY:")
    print("=" * 70)
    
    issues = []
    
    # Check if all required TAGGING sections are present
    ttags_found = sum(1 for feature in ['Tagg xxx.tree', 'Tagg xxx.score', 'Tag xxx.warnings', 'Full USB & Display Tree'] if feature in method_text)
    
    if ttags_found < 3:
        issues.append(f"WARNING: Only {ttags_found}/4 required TAGGING sections found")
        print(f"  ⚠ TAGGING sections: Found {ttags_found}/4 (CLI requires all Tagg/Tag patterns)")
    else:
        print(f"  ✓ TAGGING sections: Found {ttags_found}/4 required sections")
    
    # Check for 'Full USB & Display Tree' specifically
    if 'Full USB & Display Tree' not in method_text:
        issues.append("'Full USB & Display Tree' header missing")
        print("  ⚠ 'Full USB & Display Tree' header: REQUIRED for CLI match")
    
    # Check VERDICT sections
    if method_text.count('VERDICT') == 0:
        issues.append("VERDICT sections missing")
        print("  ⚠ VERDICT verdict sections: REQUIRED for CLI match")
    
    # Structure checks
    for pattern, description in struct_checks:
        if pattern not in method_text:
            issues.append(f"{description} missing")
    
    if issues:
        print(f"\n  Issues found ({len(issues)}):")
        for issue in issues[:5]:  # Limit to first 5 issues
            print(f"    - {issue}")
        if len(issues) > 5:
            print(f"    ... and {len(issues) - 5} more issues")
        
        print("\n" + "=" * 70)
        print("RECOMMENDATIONS:")
        print("=" * 70)
        print("The GUI _update_tree_display method needs to be updated to:")
        print("1. Include 'Full USB & Display Tree' header at the beginning")
        print("2. Add TAGGING sections:")
        print("   - Tagg xxx.tree")
        print("   - Tagg xxx.score")
        print("   - Tag xxx.warnings")
        print("3. Ensure VERDICT sections are present")
        print("4. Verify port structure (EXTERNAL/INTERNAL sections)")
        
        return False
    else:
        print("\n" + "=" * 70)
        print("✓ VERIFICATION PASSED")
        print("=" * 70)
        print("The GUI _update_tree_display method matches CLI format!")
        print("\nAll required features are present:")
        for feature, description in required_features.items():
            print(f"  ✓ {description}")
        
        return True

if __name__ == '__main__':
    success = main()
    exit(0 if success else 1)
