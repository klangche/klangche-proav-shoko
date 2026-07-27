import os
path = r'C:\Users\linus\AppData\Local\Temp\proav-shoko_report_20260726_150523.html'
if os.path.exists(path):
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    checks = [
        ('Platform info', 'Windows 10.0.26200' in content),
        ('Full tree', 'Full USB' in content),
        ('Tree structure', '└──' in content or '├──' in content),
        ('Per port', 'PER PORT' in content),
        ('External', 'EXTERNAL' in content),
        ('Internal', 'INTERNAL' in content),
        ('Overall rating', 'Overall rating' in content or 'OVERALL RATING' in content),
        ('Stability table', 'Windows' in content and 'STABLE' in content),
    ]
    for name, result in checks:
        print(name + ': ' + ('OK' if result else 'MISSING'))
else:
    print('File not found')