#!/usr/bin/env python3
"""Check GUI compatibility with CLI"""

with open('src/gui.py', 'r') as f:
    content = f.read()

print("=== GUI CLI COMPATIBILITY CHECK ===")
print()

# Check TAGGING support
has_tagg_tree = 'Tagg xxx.tree' in content
has_tagg_score = 'Tagg xxx.score' in content
has_tag_warnings = 'Tag xxx.warnings' in content
tagging_ok = has_tagg_tree and has_tagg_score and has_tag_warnings
print("TAGGING support (Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings): PASS" if tagging_ok else "TAGGING support (Tagg xxx.tree, Tagg xxx.score, Tag xxx.warnings): FAIL")

# Check 'Full USB & Display Tree' header
has_full_tree = 'Full USB & Display Tree' in content
print("Full USB & Display Tree header: PASS" if has_full_tree else "Full USB & Display Tree header: FAIL")

# Check VERDICT sections - count how many
verdict_count = content.count('VERDICT')
print(f"VERDICT sections: {verdict_count} found (at least 2 needed): PASS" if verdict_count >= 2 else f"VERDICT sections: {verdict_count} found (at least 2 needed): FAIL")

# Check PER PORT header
has_per_port = 'PER PORT' in content
print("PER PORT section header: PASS" if has_per_port else "PER PORT section header: FAIL")

# Check for _print_tag method
has_print_tag = 'def _print_tag(self, tag: str)' in content
print("_print_tag method: PASS" if has_print_tag else "_print_tag method: FAIL")

# Check for _print_verdict method
has_print_verdict = 'def _print_verdict(self, v)' in content
print("_print_verdict method: PASS" if has_print_verdict else "_print_verdict method: FAIL")

# Check for _print_section_header method
has_print_section_header = 'def _print_section_header(self, title)' in content
print("_print_section_header method: PASS" if has_print_section_header else "_print_section_header method: FAIL")

# Check layout - right side is tree
has_right_side_tree = 'self.right_frame = tk.Frame(middle_container' in content and 'tree_frame = tk.Frame(self.right_frame' in content
print("Right side contains tree (left side is log): PASS" if has_right_side_tree else "Right side contains tree (left side is log): FAIL")

print()
print("=== SUMMARY ===")
requirements = [tagging_ok, has_full_tree, verdict_count >= 2, has_per_port, has_print_tag, has_print_verdict, has_print_section_header, has_right_side_tree]
passed = sum(requirements)
total = len(requirements)
print(f"Requirements met: {passed}/{total}")

if passed == total:
    print("\nALL CLI REQUIREMENTS MET!")
else:
    print("\nSOME CLI REQUIREMENTS MISSING!")

print("\n=== DETAILED BREAKDOWN ===")

ttag_tree_count = content.count('Tagg xxx.tree')
ttag_score_count = content.count('Tagg xxx.score')
tag_warnings_count = content.count('Tag xxx.warnings')
print("\nTAGGING sections found:")
print(f"  - Tagg xxx.tree: {ttag_tree_count}")
print(f"  - Tagg xxx.score: {ttag_score_count}")
print(f"  - Tag xxx.warnings: {tag_warnings_count}")

print("\nVERDICT sections found:")
lines = content.split('\n')
for i, line in enumerate(lines):
    if 'VERDICT' in line and i+1 < len(lines) and 'VERDICT' in lines[i+1]:
        print(f"  Line {i+1}: {line.strip()[:100]}")

# Check layout configuration
print("\n=== LAYOUT CONFIGURATION ===")
print("Left side: should show logging (log panel)")
left_has_log = 'self.log_text = tk.Text' in content and 'self.log_text.pack' in content
print(f"  Log panel exists: {left_has_log}")
print("Right side: should show full tree, verdicts, warnings like cli")
right_has_tree = 'tree_frame = tk.Frame(self.right_frame)' in content
print(f"  Tree panel exists on right: {right_has_tree}")
EOF
