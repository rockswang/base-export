"""
Create English-named dummy attachments for verification test
"""
import os, pathlib

attachments_dir = os.path.join('tests', 'attachments_en')
pathlib.Path(attachments_dir).mkdir(parents=True, exist_ok=True)

files = {
    'report.docx': b'Fake DOCX - relative path test',
    'data.xlsx': b'Fake XLSX - relative path test',
    'notes.txt': b'Fake TXT - relative path test\nLine 2\nLine 3',
    'image.png': b'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',  # 1x1 red PNG
}

for name, content in files.items():
    with open(os.path.join(attachments_dir, name), 'wb') as f:
        f.write(content)
    print(f'  OK {name}')
