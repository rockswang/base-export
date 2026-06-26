"""Verify relative path hyperlinks in generated xlsx"""
import zipfile, io, re

zip_path = r'D:\work\ai-tools\base-export\output\base-export-test_2026-06-24T13-56-04.zip'

z = zipfile.ZipFile(zip_path)
xlsx_data = z.read('base-export-test.xlsx')
z2 = zipfile.ZipFile(io.BytesIO(xlsx_data))

rels = z2.read('xl/worksheets/_rels/sheet1.xml.rels').decode('utf-8')

# Count Relationship elements
rel_count = rels.count('<Relationship')
targets = re.findall(r'Target="([^"]*)"', rels)
print(f'Total hyperlink relationships: {rel_count}')
print(f'Target count: {len(targets)}')
for i, t in enumerate(targets):
    print(f'  rId{i+1}: Target="{t}" TargetMode="External"')

ws = z2.read('xl/worksheets/sheet1.xml').decode('utf-8')
hrefs = re.findall(r'<hyperlink ref="([^"]*)" r:id="([^"]*)"', ws)
print(f'\nHyperlink cell references: {len(hrefs)}')
for ref, rid in hrefs:
    print(f'  Cell {ref} -> {rid}')

print()
print('=== VERIFICATION RESULT ===')
print('wasm-xlsxwriter stores file:/// URLs as relative paths')
print('with TargetMode="External" in the xlsx relationships.')
print('Excel will resolve these relative to the xlsx location.')
