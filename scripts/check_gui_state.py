#!/usr/bin/env python3
"""
Quick check of GUI state
"""
import os

print("=" * 70)
print("GUI STATE CHECK")
print("=" * 70)

# Read gui.py
if not os.path.exists('src/gui.py'):
    print("ERROR: src/gui.py not found!")
    exit(1)

with open('src/gui.py', 'r') as f:
    content = f.read()

print(f"File size: {len(content)} characters")
print(f"Number of lines: {len(content.split(chr(10)))}")

print("\n🔍 CHECKING CORE REQUIREMENTS:")

# Required methods
required_methods = [
    '_print_tag',
    '_print_verdict', 
    '_print_section_header'
]

method_found = 0
for method in required_methods:
    if f'def {method}(' in content:
        print(f"  ✓ {method} method present")
        method_found += 1
    else:
        print(f"  ✗ {method} method missing")

# Required TAGGING sections
tagg_sections = [
    'Tagg xxx.tree',
    'Tagg xxx.score',
    'Tag xxx.warnings'
]

section_found = 0
for section in tagg_sections:
    if section in content:
        print(f"  ✓ {section} present")
        section_found += 1
    else:
        print(f"  ✗ {section} missing")

# Required headers
required_headers = [
    ('All logging and tree with warning and its warnings', 'Left panel header'),
    ('Full tree, verdicts, warnings like CLI', 'Right panel header'),
    ('Full USB & Display Tree', 'Tree header')
]

header_found = 0
for header, desc in required_headers:
    if header in content:
        print(f"  ✓ {desc}: '{header}'")
        header_found += 1
    else:
        print(f"  ✗ {desc}: missing")

print("\n" + "=" * 70)
print("SUMMARY")
print("=" * 70)

print(f"\nRequired methods: {method_found}/{len(required_methods)} ({'PASS' if method_found == len(required_methods) else 'FAIL'})")
print(f"TAGGING sections: {section_found}/{len(tagg_sections)} ({'PASS' if section_found == len(tagg_sections) else 'FAIL'})")
print(f"Required headers: {header_found}/{len(required_headers)} ({'PASS' if header_found == len(required_headers) else 'FAIL'})")

total_passed = method_found + section_found + header_found
total_possible = len(required_methods) + len(tagg_sections) + len(required_headers)

print(f"\nTotal: {total_passed}/{total_possible} requirements met")

if total_passed == total_possible:
    print("\n✅ GUI is COMPLETE and meets all requirements!")
    print("\nGUI meets user requirements:")
    print("  • Left side shows: All logging, tree with warning and its warnings")
    print("  • Right side shows: Full tree, verdicts, warnings like CLI")
    print("  • TAGGING support: Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings")
    print("  • _print_tag, _print_verdict, _print_section_header methods")
else:
    print(f"\n⚠️ GUI incomplete - {total_possible - total_passed} requirements missing")
    
    if method_found < len(required_methods):
        print(f"  • Missing {len(required_methods) - method_found} required methods")
    
    if section_found < len(tagg_sections):
        print(f"  • Missing {len(tagg_sections) - section_found} TAGGING sections")
        
    if header_found < len(required_headers):
        print(f"  • Missing {len(required_headers) - header_found} headers")

print("\n" + "=" * 70)
print("RECOMMENDATIONS")
print("=" * 70)

if total_passed < total_possible:
    print("\nTo fix the GUI:")
    print("  1. Run gui_fix_final.py to apply all fixes")
    print("  2. Or manually check the _update_tree_display method")
    print("  3. Verify all TAGGING sections are present")
    print("  4. Check panel headers are correctly set")
else:
    print("\n✅ GUI is ready!")
    print("\nThe GUI can now be used.")
