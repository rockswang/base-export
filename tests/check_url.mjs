import xlsxInit from 'wasm-xlsxwriter/web';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { readFileSync, writeFileSync } from 'fs';

const __dirname = dirname(fileURLToPath(import.meta.url));
const wasmPath = resolve(__dirname, '..', 'node_modules', 'wasm-xlsxwriter', 'web', 'wasm_xlsxwriter_bg.wasm');
await xlsxInit({ module_or_path: readFileSync(wasmPath) });

const { Workbook, Url } = await import('wasm-xlsxwriter/web');

// Try to craft a relative URL that writeUrl accepts
// Test: does "file:///" followed by a relative path produce a relative link in Excel?
const wb = new Workbook();
const ws = wb.addWorksheet();
ws.setName('Test');

// Write some data
ws.write(0, 0, '文件名');
ws.write(0, 1, '超链接');

ws.write(1, 0, '项目方案.pdf');
ws.writeUrl(1, 1, new Url('file:///attachments/项目方案.pdf').setText('打开文件'));
// This writes file:///attachments/项目方案.pdf as absolute path from drive root

ws.write(2, 0, '需求文档.docx');
// Use "internal:" as workaround? No, that's for internal workbook links

ws.write(3, 0, '会议纪要.txt');
ws.writeUrl(3, 1, new Url('file:///./attachments/会议纪要.txt').setText('打开文件'));
// This might be interpreted as "current directory + attachments/会议纪要.txt"

const buf = wb.saveToBufferSync();
const outPath = resolve(__dirname, '..', 'output', 'url-test.xlsx');
writeFileSync(outPath, buf);
console.log(`✓ Saved to ${outPath}`);
console.log(`  Size: ${buf.length} bytes`);
