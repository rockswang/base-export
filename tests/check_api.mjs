// Use the default export from wasm-xlsxwriter/web - it's a factory function
// that can accept the wasm module path
import xlsxInit from 'wasm-xlsxwriter/web';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { readFileSync } from 'fs';

const __dirname = dirname(fileURLToPath(import.meta.url));
const wasmPath = resolve(__dirname, '..', 'node_modules', 'wasm-xlsxwriter', 'web', 'wasm_xlsxwriter_bg.wasm');

console.log('WASM path:', wasmPath);

try {
    // Try passing wasm module directly
    await xlsxInit({ module_or_path: readFileSync(wasmPath) });
    console.log('✓ xlsxInit with Buffer succeeded');
} catch(e) {
    console.log('Failed with Buffer:', e.message);
    
    // Try with URL path
    try {
        const wasmUrl = new URL(`file:///${wasmPath.replace(/\\/g, '/')}`);
        await xlsxInit(wasmUrl);
        console.log('✓ xlsxInit with file URL succeeded');
    } catch(e2) {
        console.log('Failed with file URL:', e2.message);
    }
}

const { Format, FormatAlign, FormatBorder, Color, Url, Workbook } = await import('wasm-xlsxwriter/web');
console.log('✓ imported all symbols');

const fmt = new Format();
fmt.setBold()
   .setBackgroundColor(Color.rgb(0x4472C4))
   .setFontColor(Color.white())
   .setAlign(FormatAlign.Center)
   .setAlign(FormatAlign.VerticalCenter)
   .setBorder(FormatBorder.Thin);
console.log('✓ Format created');

const url = new Url('attachments/test.pdf').setText('test.pdf (123 B)');
console.log('✓ Url created');

const wb = new Workbook();
const ws = wb.addWorksheet();
ws.writeUrl(0, 0, new Url('attachments/test.pdf').setText('click me'));
console.log('✓ worksheet writeUrl');

const buf = wb.saveToBufferSync();
console.log(`✓ workbook saved: ${buf.length} bytes`);
