import os

files = [
    r'c:\ NEXO SDK SFERA\HrAppka Import Pracowników\HrAppka Import Pracowników.csproj',
    r'c:\ NEXO SDK SFERA\HrAppka Import Pracowników\Konfiguracja.xml'
]

for path in files:
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    new_content = content.replace('1.0.8.19', '1.0.8.20')
    with open(path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    print(f"Updated {os.path.basename(path)} -> 1.0.8.20: {'1.0.8.20' in new_content}")
