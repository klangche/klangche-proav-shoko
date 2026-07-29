import sys

# Check if text_color exists in gui.py
with open('src/gui.py', 'r') as f:
    content = f.read()

if 'text_color' in content:
    print('ERROR: text_color still found in gui.py')
    sys.exit(1)
else:
    print('SUCCESS: No text_color in gui.py')
    sys.exit(0)