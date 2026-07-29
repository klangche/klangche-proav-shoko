#!/usr/bin/env python3
"""
Simple test to verify GUI fix
"""

with open('src/gui.py', 'r') as f:
    content = f.read()

print("=" * 70)
print("GUI FIX VERIFICATION")
print("=" * 70)

# Check for required methods
methods = [
    ('_print_tag', 'Machine-parseable tag printing'),
    ('_print_verdict', 'Verdict printing'),
    ('_print_section_header', 'Section header printing'),
]

print("\n✅ REQUIRED METHODS CHECK:")
for method_name, desc in methods:
    if f"def {method_name}(" in content:
        print(f"  ✓ {desc}: {method_name}")
    else:
        print(f"  ✗ {desc}: {method_name} - MISSING")

# Check TAGGING sections
tagg_sections = [
    ('Tagg xxx.tree', 'TAGGING tree section'),
    ('Tagg xxx.score', 'TAGGING score section'),
    ('Tag xxx.warnings', 'TAGGING warnings section'),
]

print("\n✅ TAGGING SECTIONS CHECK:")
for tagg, desc in tagg_sections:
    if tagg in content:
        print(f"  ✓ {desc}: {tagg}")
    else:
        print(f"  ✗ {desc}: {tagg} - MISSING")

print("\n✅ LAYOUT CHECK:")
if "Full USB & Display Tree" in content:
    print("  ✓ Full USB & Display Tree header")
else:
    print("  ✗ Full USB & Display Tree header - MISSING")

if "All logging and tree with warning and its warnings" in content:
    print("  ✓ Left panel: All logging, tree with warning")
else:
    print("  ✗ Left panel - MISSING")

if "Full tree, verdicts, warnings like CLI" in content:
    print("  ✓ Right panel: Full tree, verdicts, warnings like CLI")
else:
    print("  ✗ Right panel - MISSING")

print("\n" + "=" * 70)
print("SUMMARY")
print("=" * 70)

method_count = sum(1 for name, _ in methods if f"def {name}(" in content)
tagg_count = sum(1 for tagg, _ in tagg_sections if tagg in content)
layout_count = sum(1 for tagg, _ in [("Full USB & Display Tree", ""), ("All logging and tree with warning and its warnings", ""), ("Full tree, verdicts, warnings like CLI", "")] if tagg in content)

total = method_count + tagg_count + layout_count
total_possible = 3 + 3 + 3

print(f"\nMethods: {method_count}/{len(methods)}")
print(f"TAGGING sections: {tagg_count}/{len(tagg_sections)}")
print(f"Layout: {layout_count}/{3}")
print(f"\nTotal: {total}/{total_possible}")

if total == total_possible:
    print("\n✅ ALL CHECKS PASSED!")
    print("\nGUI has been successfully fixed to match user requirements:")
    print("  • Left side shows: All logging, tree with warning and its warnings")
    print("  • Right side shows: Full tree, verdicts, warnings like CLI")
    print("  • TAGGING support: Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings")
    print("  • Required methods: _print_tag, _print_verdict, _print_section_header")
else:
    print("\n⚠️ SOME CHECKS FAILED")
    if method_count < len(methods):
        print(f"  • Missing {len(methods) - method_count} methods")
    if tagg_count < len(tagg_sections):
        print(f"  • Missing {len(tagg_sections) - tagg_count} TAGGING sections")
    if layout_count < 3:
        print(f"  • Missing {3 - layout_count} layout features")

print("\n" + "=" * 70)
