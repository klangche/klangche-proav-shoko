#!/usr/bin/env python3
"""
Update the _update_tree_display method in src/gui.py to match CLI exactly
"""

import re

# Read the current file
with open('src/gui.py', 'r') as f:
    content = f.read()

# Find the _update_tree_display method
method_start = content.find('    def _update_tree_display(self, usb_tree, hops_data, stability, displays):')
if method_start == -1:
    print('ERROR: Could not find _update_tree_display method')
    exit(1)

# Find the end of the method (next method or end of file)
lines = content.split('\n')
method_start_line = None
for i, line in enumerate(lines):
    if '    def _update_tree_display(' in line:
        method_start_line = i
        break

if method_start_line is None:
    print('ERROR: Could not find _update_tree_display method')
    exit(1)

# Find the end of this method (next method at same indentation level)
method_lines = []
indent_level = None
for i in range(method_start_line, len(lines)):
    line = lines[i]
    if i == method_start_line:
        method_lines.append(line)
        # Extract indent level (number of leading spaces)
        indent_level = len(line) - len(line.lstrip())
        continue
    
    # Check if we've reached the end of this method
    if line.strip() and len(line) - len(line.lstrip()) == 0 and 'def ' in line:
        # Next method at top level - end this method
        break
    
    method_lines.append(line)

# Get the current method
method_text = '\n'.join(method_lines)

print('Current _update_tree_display method:')
print('=' * 60)
print(method_text)
print('=' * 60)

# Check if it has VERDICT and Full USB & Display Tree
if 'VERDICT' in method_text and 'Full USB & Display Tree' in method_text:
    print('\n✓ Method appears to have VERDICT and "Full USB & Display Tree"')
    
    # Check specific structure
    if 'for idx, child in enumerate(orig_children):' in method_text:
        print('✓ Has port iteration loop')
    else:
        print('✗ Missing port iteration loop')
else:
    print('\n✗ Method missing VERDICT or "Full USB & Display Tree"')
    print('This needs to be fixed')

print('\n' + '=' * 60)
print('Method needs to match CLI exactly:')
print('=' * 60)
print('  - "Full USB & Display Tree" at the beginning')
print('  - Port structure with proper indentation')
print('  - VERDICT sections for stability assessment')
print('  - Per-port details (external and internal)')
