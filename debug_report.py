import os
path = r'C:\Users\linus\AppData\Local\Temp\proav-shoko_report_20260726_150523.html'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()
# Find the body content
start = content.find('<body>')
end = content.find('</body>')
if start != -1 and end != -1:
    body = content[start:end]
    with open('report_body.txt', 'w', encoding='utf-8') as out:
        out.write(body[:5000])
    print('Written to report_body.txt')