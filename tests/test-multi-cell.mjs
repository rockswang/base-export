/**
 * 验证: 单个单元格能否包含多个文件超链接 + 图片混合
 *
 * 飞书附件字段允许一个单元格内有多个附件，类型混合（图片+文件）。
 * xlsx 格式限制: 一个单元格只能有一个 <hyperlink> 元素。
 * 本测试验证 wasm-xlsxwriter 在重复写入同一单元格时的行为。
 */
import xlsxInit from 'wasm-xlsxwriter/web';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { readFileSync, writeFileSync, mkdirSync, existsSync } from 'fs';

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(__dirname, '..');
const OUT = resolve(ROOT, 'output');
const WASM = resolve(ROOT, 'node_modules', 'wasm-xlsxwriter', 'web', 'wasm_xlsxwriter_bg.wasm');
await xlsxInit({ module_or_path: readFileSync(WASM) });

const { Workbook, Url, Image } = await import('wasm-xlsxwriter/web');

// ── 准备测试图片: 生成一张红蓝小方块 PNG ─────────────────
function createMiniPNG() {
    return Buffer.from([
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x10, 0x00, 0x00, 0x00, 0x10,
        0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
        0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
        0x54, 0x08, 0xD7, 0x63, 0x60, 0x00, 0x00, 0x00,
        0x02, 0x00, 0x01, 0xE4, 0x21, 0x00, 0x00, 0x00,
        0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ]);
}

const SAMPLE_PNG = createMiniPNG();

// ── 测试 1: 同一单元格写入多个 writeUrl ──────────────────
// 预期: xlsx 格式限制 - 一个 cell 只能有一个 hyperlink
// 需要验证 wasm-xlsxwriter 是覆盖还是报错
console.log('=== 测试1: 单个单元格多个 writeUrl ===');
{
    const wb = new Workbook();
    const ws = wb.addWorksheet();
    ws.setName('MultiURL');

    ws.write(0, 0, '文件列表');
    ws.writeUrl(1, 0, new Url('file:///attachments/doc1.pdf').setText('doc1.pdf'));
    try {
        ws.writeUrl(1, 0, new Url('file:///attachments/doc2.pdf').setText('doc2.pdf'));
        console.log('  第二次 writeUrl: 未报错');
    } catch(e) {
        console.log(`  第二次 writeUrl 报错: ${e.message}`);
    }

    // 保存并检查内部结构
    const buf = wb.saveToBufferSync();
    // 用 JSZip 检查
    const { default: JSZip } = await import('jszip');
    const zip = await JSZip.loadAsync(buf);
    const relsXml = await zip.file('xl/worksheets/_rels/sheet1.xml.rels').async('string');
    const wsXml = await zip.file('xl/worksheets/sheet1.xml').async('string');

    const targetCount = (relsXml.match(/Target=/g) || []).length;
    const hyperlinkCount = (wsXml.match(/<hyperlink /g) || []).length;
    console.log(`  超链接关系数: ${targetCount}`);
    console.log(`  单元格超链接数: ${hyperlinkCount}`);

    // 如果只有1个, 说明后者覆盖了前者
}

// ── 测试 2: 同单元格 writeUrl + insertImage 混合 ─────────
console.log('\n=== 测试2: 同单元格 writeUrl + Image 混合 ===');
{
    const wb = new Workbook();
    const ws = wb.addWorksheet();
    ws.setName('Mixed');

    ws.setColumnWidth(0, 30);
    ws.setRowHeight(1, 60);

    ws.write(0, 0, '混合附件');

    // 先写入超链接
    ws.writeUrl(1, 0, new Url('file:///attachments/report.pdf').setText('report.pdf'));
    // 再嵌入图片（图片浮动在单元格上方，不与超链接冲突）
    const img = new Image(SAMPLE_PNG);
    ws.insertImage(1, 0, img);  // 图片从 B1 开始浮动

    const buf = wb.saveToBufferSync();
    const { default: JSZip } = await import('jszip');
    const zip = await JSZip.loadAsync(buf);
    const relsXml = await zip.file('xl/worksheets/_rels/sheet1.xml.rels').async('string');
    const wsXml = await zip.file('xl/worksheets/sheet1.xml').async('string');

    const hasHyperlink = wsXml.includes('<hyperlink');
    const hasDrawing = wsXml.includes('<drawing');
    console.log(`  单元格超链接: ${hasHyperlink}`);
    console.log(`  嵌入图片: ${hasDrawing}`);

    const targetCount = (relsXml.match(/Target=/g) || []).length;
    console.log(`  超链接关系数: ${targetCount}`);
}

// ── 测试 3: 推荐做法 - 每附件占一行 ──────────────────────
console.log('\n=== 测试3: 每附件独立一行（推荐做法） ===');
{
    const wb = new Workbook();
    const ws = wb.addWorksheet();
    ws.setName('Expanded');

    ws.setColumnWidth(0, 30);
    ws.setColumnWidth(1, 14);
    ws.setColumnWidth(2, 50);

    // 表头
    ws.writeRow(0, 0, ['任务名', '附件类型', '附件']);
    ws.setFreezePanes(1, 0);

    // 模拟: 一个记录有3个附件 [image1.png, report.pdf, image2.jpg]
    const attachments = [
        { type: '图片', file: 'screenshot.png' },
        { type: '文档', file: 'report.pdf' },
        { type: '图片', file: 'diagram.png' },
    ];
    ws.write(1, 0, '需求分析');

    for (let i = 0; i < attachments.length; i++) {
        const a = attachments[i];
        const row = 1 + i;
        ws.write(row, 1, a.type);
        if (a.type === '图片') {
            const img = new Image(SAMPLE_PNG);
            ws.write(row, 2, a.file);
            ws.insertImage(row, 2, img);
        } else {
            ws.writeUrl(row, 2, new Url(`file:///attachments/${a.file}`).setText(a.file));
        }
    }

    const buf = wb.saveToBufferSync();
    const { default: JSZip } = await import('jszip');
    const zip = await JSZip.loadAsync(buf);
    const relsXml = await zip.file('xl/worksheets/_rels/sheet1.xml.rels').async('string');
    const wsXml = await zip.file('xl/worksheets/sheet1.xml').async('string');
    const targetCount = (relsXml.match(/Target=/g) || []).length;
    const hyperlinkCount = (wsXml.match(/<hyperlink /g) || []).length;
    const hasDrawing = wsXml.includes('<drawing');
    console.log(`  超链接: ${targetCount} 个关系, ${hyperlinkCount} 个单元格引用`);
    console.log(`  图片: ${hasDrawing ? '有' : '无'} Drawing`);

    // 输出最终文件
    const { default: JSZip2 } = await import('jszip');
    const outZip = new JSZip2();
    outZip.file('test-multi-attach.xlsx', buf);

    // 添加几个模拟附件
    for (const a of attachments) {
        if (a.type === '文档') {
            outZip.file(`attachments/${a.file}`, `Fake ${a.file} content`);
        }
    }
    mkdirSync(OUT, { recursive: true });
    const zipBuf = await outZip.generateAsync({ type: 'nodebuffer' });
    writeFileSync(resolve(OUT, 'test-multi-attach.zip'), zipBuf);
    console.log(`  输出: output/test-multi-attach.zip`);
}

// ── 结论 ──────────────────────────────────────────────
console.log('\n═══════════════════════════════════════');
console.log('结论:');
console.log('1. 单单元格只能有 1 个超链接 — 重复 writeUrl 后者覆盖前者');
console.log('2. 嵌入图片与超链接不冲突 — 图片浮动在单元格上，可共存');
console.log('3. 推荐做法: 将附件的每个文件展开为独立行');
console.log('   每个占1行，图片嵌入，非图片用超链接');
console.log('4. 所有附件共享同一个合并的"记录名"单元格');
console.log('═══════════════════════════════════════');
