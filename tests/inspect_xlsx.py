import zipfile
import os

os.makedirs('output/url-test_extracted', exist_ok=True)
with zipfile.ZipFile('output/url-test.xlsx') as z:
    z.extractall('output/url-test_extracted')
    for f in z.namelist():
        print(f)
