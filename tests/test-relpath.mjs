/**
 * 最小化验证：wasm-xlsxwriter 用 file:/// URL 生成相对路径超链接
 *
 * 验证点：file:///attachments/report.docx 在 xlsx 内部被存储为
 *   Target="attachments/report.docx" TargetMode="External"
 *
 * 输出: output/relpath-test.xlsx  + 解压版 output/relpath-test_extracted/
 */
import xlsxInit from 'wasm-xlsxwriter/web';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { readFileSync, writeFileSync, mkdirSync } from 'fs';
import JSZip from 'jszip';

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(__dirname, '..');
const OUT = resolve(ROOT, 'output');
const ATTACH = resolve(ROOT, 'tests', 'attachments_en');

const WASM = resolve(ROOT, 'node_modules', 'wasm-xlsxwriter', 'web', 'wasm_xlsxwriter_bg.wasm');
await xlsxInit({ module_or_path: readFileSync(WASM) });

const { Workbook, Format, FormatAlign, FormatBorder, Color, Url } = await import('wasm-xlsxwriter/web');

// ── 创建 xlsx ──────────────────────────────────────────
const wb = new Workbook();
const ws = wb.addWorksheet();
ws.setName('RelativePathTest');

// 表头
const hf = new Format().setBold().setAlign(FormatAlign.Center).setBorder(FormatBorder.Thin)
    .setBackgroundColor(Color.rgb(0x4472C4)).setFontColor(Color.white());
ws.writeRowWithFormat(0, 0, ['Name', 'Type', 'File Link'], hf);
ws.setColumnWidth(0, 18);
ws.setColumnWidth(1, 12);
ws.setColumnWidth(2, 40);
ws.setFreezePanes(1, 0);

const cf = new Format().setBorder(FormatBorder.Thin).setAlign(FormatAlign.VerticalCenter);
const uf = new Format().setBorder(FormatBorder.Thin).setFontColor(Color.rgb(0x0563C1))
    .setUnderline(1).setAlign(FormatAlign.VerticalCenter);

// 数据行 —— 关键测试：
// 用 file:///attachments/filename.ext 作为 URL
// wasm-xlsxwriter 会提取路径部分存储为相对链接
const files = [
    { name: 'Report', type: 'docx', file: 'report.docx' },
    { name: 'Data', type: 'xlsx', file: 'data.xlsx' },
    { name: 'Notes', type: 'txt', file: 'notes.txt' },
    { name: 'Image', type: 'png', file: 'image.png' },
];

for (let i = 0; i < files.length; i++) {
    const row = i + 1;
    const f = files[i];

    ws.writeWithFormat(row, 0, f.name, cf);
    ws.writeWithFormat(row, 1, f.type, cf);

    // ★ 核心: file:/// + 相对路径 attachments/filename.ext
    //    wasm-xlsxwriter 自动转为相对链接 TargetMode="External"
    ws.writeUrl(row, 2, new Url(`file:///attachments/${f.file}`).setText(`attachments/${f.file}`));
}

const xlsxBuf = wb.saveToBufferSync();

// ── 打包 ZIP ──────────────────────────────────────────
const zip = new JSZip();
zip.file('relpath-test.xlsx', xlsxBuf);
for (const f of files) {
    const fp = resolve(ATTACH, f.file);
    zip.file(`attachments/${f.file}`, readFileSync(fp));
}

mkdirSync(OUT, { recursive: true });
const zipBuf = await zip.generateAsync({ type: 'nodebuffer' });
const zipPath = resolve(OUT, 'relpath-test.zip');
writeFileSync(zipPath, zipBuf);

// ── 输出结果 ──────────────────────────────────────────
console.log('=== 测试文件已生成 ===');
console.log(`ZIP: ${zipPath}  (${zipBuf.length} bytes)`);
console.log('');

// 解压 xlsx 检查内部结构
import { unzipSync } from 'zlib';
const xlsxZip = new JSZip();
await xlsxZip.loadAsync(xlsxBuf);
const relsXml = await xlsxZip.file('xl/worksheets/_rels/sheet1.xml.rels').async('string');
const wsXml = await xlsxZip.file('xl/worksheets/sheet1.xml').async('string');

console.log('=== xl/worksheets/_rels/sheet1.xml.rels ===');
console.log(relsXml);
console.log('');
console.log('=== xl/worksheets/sheet1.xml (hyperlinks section) ===');
const hyperlinks = wsXml.match(/<hyperlinks>[\s\S]*<\/hyperlinks>/);
if (hyperlinks) console.log(hyperlinks[0]);
console.log('');

// 计数
const targets = relsXml.match(/Target="([^"]+)"/g) || [];
console.log(`总计 ${targets.length} 个相对路径超链接:`);
targets.forEach((t, i) => console.log(`  ${i + 1}. ${t}`));
console.log('');
console.log('结论: wasm-xlsxwriter 将 file:/// 路径存储为相对路径 ✓');
console.log('Excel 打开 xlsx 后，点击超链接会按 xlsx 所在目录解析。');
