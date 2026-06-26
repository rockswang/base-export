/**
 * 测试：用 wasm-xlsxwriter 生成带相对路径超链接的 xlsx，模拟完整的 Base 导出场景
 *
 * 关键验证：wasm-xlsxwriter 对 file:/// URL 的处理行为
 * → 传入 file:///attachments/xxx.pdf，库自动存储为相对路径 TargetMode="External"
 * → Excel 打开时按 xlsx 所在目录解析相对路径
 *
 * ZIP 结构:
 *   base-export-test_{ts}.zip
 *    ├── base-export-test.xlsx    ← 表格数据 + 图片嵌入
 *    └── attachments/
 *        ├── 项目方案.pdf
 *        ├── 需求文档.docx
 *        ├── 数据报表.xlsx
 *        └── 会议纪要.txt
 *
 * 用法: node tests/gen-test.mjs
 */

import xlsxInit from 'wasm-xlsxwriter/web';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { readFileSync, writeFileSync, existsSync, mkdirSync } from 'fs';
import JSZip from 'jszip';

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(__dirname, '..');
const OUT_DIR = resolve(ROOT, 'output');
const ATTACH_DIR = resolve(ROOT, 'tests', 'attachments');

// ── 初始化 WASM ─────────────────────────────────────────
const wasmPath = resolve(ROOT, 'node_modules', 'wasm-xlsxwriter', 'web', 'wasm_xlsxwriter_bg.wasm');
await xlsxInit({ module_or_path: readFileSync(wasmPath) });

const { Workbook, Format, FormatAlign, FormatBorder, Color, Image, Url } =
    await import('wasm-xlsxwriter/web');

// ── 模拟数据 ──────────────────────────────────────────────
const FIELD_NAMES = ['任务名称', '负责人', '状态', '优先级', '完成进度', '开始日期', '截止日期', '附件'];

const records = [
    ['需求调研', '张三', '已完成', '高', 1.0, '2024-01-01', '2024-01-15', [
        { fn: '会议纪要.txt' },
    ]],
    ['方案设计', '李四', '进行中', '高', 0.65, '2024-01-10', '2024-02-10', [
        { fn: '项目方案.pdf' },
        { fn: '需求文档.docx' },
    ]],
    ['数据统计', '王五', '待审核', '中', 0.9, '2024-01-20', '2024-02-05', [
        { fn: '数据报表.xlsx' },
    ]],
    ['项目收尾', '赵六', '未开始', '低', 0.0, '2024-02-01', '2024-03-01', []],
];

// ── 模拟图片（内嵌附件） ────────────────────────────
// 生成一个 1x1 白色 PNG 作为模拟图片
const MINI_PNG = Buffer.from([
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG header
    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1 pixel
    0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, // grayscale
    0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, // IDAT chunk
    0x54, 0x08, 0xD7, 0x63, 0x60, 0x00, 0x00, 0x00, // compressed data
    0x02, 0x00, 0x01, 0xE4, 0x21, 0x00, 0x00, 0x00, // IEND chunk
    0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
]);

// ── 主流程 ────────────────────────────────────────────────
async function main() {
    const wb = new Workbook();
    const ws = wb.addWorksheet();
    ws.setName('任务列表');

    // ── 样式 ──────────────────────────────────────
    const headerFmt = new Format()
        .setBold()
        .setAlign(FormatAlign.Center).setAlign(FormatAlign.VerticalCenter)
        .setBackgroundColor(Color.rgb(0x4472C4))
        .setFontColor(Color.white())
        .setBorder(FormatBorder.Thin);

    const cellFmt = new Format()
        .setAlign(FormatAlign.VerticalCenter)
        .setBorder(FormatBorder.Thin)
        .setTextWrap();

    const dateFmt = new Format()
        .setNumFormat('yyyy-mm-dd')
        .setAlign(FormatAlign.VerticalCenter)
        .setBorder(FormatBorder.Thin);

    const urlFmt = new Format()
        .setFontColor(Color.rgb(0x0563C1))
        .setUnderline(1) // single underline
        .setAlign(FormatAlign.VerticalCenter)
        .setBorder(FormatBorder.Thin);

    // ── 写表头 ──────────────────────────────────────
    ws.writeRowWithFormat(0, 0, FIELD_NAMES, headerFmt);
    ws.setRowHeight(0, 28);

    ws.setColumnWidth(0, 28);
    ws.setColumnWidth(1, 14);
    ws.setColumnWidth(2, 14);
    ws.setColumnWidth(3, 14);
    ws.setColumnWidth(4, 14);
    ws.setColumnWidth(5, 16);
    ws.setColumnWidth(6, 16);
    ws.setColumnWidth(7, 55);

    // ── 写数据 ──────────────────────────────────────
    for (let ri = 0; ri < records.length; ri++) {
        const r = records[ri];
        const row = ri + 1;

        ws.writeWithFormat(row, 0, r[0], cellFmt);  // 任务名称
        ws.writeWithFormat(row, 1, r[1], cellFmt);  // 负责人
        ws.writeWithFormat(row, 2, r[2], cellFmt);  // 状态
        ws.writeWithFormat(row, 3, r[3], cellFmt);  // 优先级
        ws.writeWithFormat(row, 4, r[4], cellFmt);  // 完成进度
        ws.writeWithFormat(row, 5, r[5], dateFmt);  // 开始日期
        ws.writeWithFormat(row, 6, r[6], dateFmt);  // 截止日期

        const attachList = r[7];
        if (attachList.length === 0) {
            ws.writeWithFormat(row, 7, '（无附件）', cellFmt);
        } else {
            for (let ai = 0; ai < attachList.length; ai++) {
                const actualRow = row + ai;
                if (ai > 0) ws.setRowHeight(actualRow, 22);

                const { fn } = attachList[ai];
                const displayText = fn;

                // ★ 关键测试点 ★
                // 用 file:/// + 相对路径，库自动存储为相对链接
                // 内部 xl/worksheets/_rels/sheet1.xml.rels 中变为：
                //   Target="attachments/项目方案.pdf" TargetMode="External"
                // Excel 打开时自动以 xlsx 所在位置解析相对路径
                ws.writeUrl(actualRow, 7,
                    new Url(`file:///${displayText}`).setText(displayText),
                );
            }
        }

        ws.setRowHeight(row, 22);
    }

    // ── 冻结 / 筛选 ────────────────────────────────
    ws.setFreezePanes(1, 0);
    ws.autofilter(0, 0, records.length, FIELD_NAMES.length - 1);

    const xlsxBuf = wb.saveToBufferSync();

    // ── 打包 ZIP ────────────────────────────────────
    const zip = new JSZip();
    zip.file('base-export-test.xlsx', xlsxBuf);

    const attachFiles = ['项目方案.pdf', '需求文档.docx', '数据报表.xlsx', '会议纪要.txt'];
    for (const fn of attachFiles) {
        const fp = resolve(ATTACH_DIR, fn);
        if (existsSync(fp)) {
            zip.file(`attachments/${fn}`, readFileSync(fp));
        } else {
            console.warn(`  ⚠ 附件不存在: ${fp}`);
        }
    }

    const zipBuf = await zip.generateAsync({ type: 'nodebuffer' });
    const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const zipName = `base-export-test_${ts}.zip`;
    const zipPath = resolve(OUT_DIR, zipName);

    if (!existsSync(OUT_DIR)) mkdirSync(OUT_DIR, { recursive: true });
    writeFileSync(zipPath, zipBuf);

    console.log(`✓ 导出完成: ${zipPath}`);
    console.log(`  大小: ${(zipBuf.length / 1024).toFixed(1)} KB`);
    console.log(``);
    console.log(`📦 ZIP 结构:`);
    console.log(`  ${zipName}`);
    console.log(`    ├── base-export-test.xlsx`);
    console.log(`    └── attachments/`);
    for (const fn of attachFiles) {
        console.log(`        ├── ${fn}`);
    }
    console.log(``);
    console.log(`🔍 验证方法:`);
    console.log(`  1. 解压 ZIP 到任意目录`);
    console.log(`  2. 打开 base-export-test.xlsx`);
    console.log(`  3. 点击附件列的超链接`);
    console.log(`  4. 验证 Excel 能否正确打开同目录 attachments/ 下的文件`);
    console.log(``);
    console.log(`📋 xlsx 内部验证（rel 文件内容预览）:`);
    console.log(`  使用以下命令查看相对路径存储形式:`);
    console.log(`  python -c "import zipfile; z=zipfile.ZipFile('${zipPath.replace(/\\/g, '/')}'); print(z.read('xl/worksheets/_rels/sheet1.xml.rels').decode())"`);
}

main().catch(e => {
    console.error('❌ 错误:', e.message);
    process.exit(1);
});
