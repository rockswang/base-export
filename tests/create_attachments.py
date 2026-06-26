"""创建模拟附件文件用于测试"""
import os

attachments_dir = 'tests/attachments'
os.makedirs(attachments_dir, exist_ok=True)

files = {
    '项目方案.pdf': b'%PDF-1.4 fake pdf content for testing relative path hyperlinks' + b'\n' * 500,
    '需求文档.docx': b'PK\x03\x04 fake docx content for testing' + b'\n' * 300,
    '数据报表.xlsx': b'PK\x03\x04 fake xlsx content for testing' + b'\n' * 400,
    '会议纪要.txt': '会议时间：2024年1月15日\n参会人员：张三、李四\n会议内容：讨论项目进展\n'.encode('utf-8'),
}

for name, content in files.items():
    with open(os.path.join(attachments_dir, name), 'wb') as f:
        f.write(content)
    print(f'  OK {name}')
