import os
for f in os.listdir('tests/attachments'):
    p = os.path.join('tests/attachments', f)
    print(f'  {f} ({os.path.getsize(p)}B)')
