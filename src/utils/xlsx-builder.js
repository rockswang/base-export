import initWasm, { Workbook, Url } from 'wasm-xlsxwriter/web';

export class XlsxBuilder {
    async init() {
        await initWasm();
        return this;
    }

    build({ rows }) {
        const wb = new Workbook();
        const ws = wb.addWorksheet();
        ws.setName('导出数据');

        // 表头
        ws.writeRow(0, 0, rows.header || []);

        // 数据行
        for (let i = 0; i < rows.data.length; i++) {
            const row = rows.data[i];
            const r = i + 1;
            ws.setRowHeight(r, row.height || 20);

            for (let c = 0; c < row.cells.length; c++) {
                const cell = row.cells[c];
                if (cell === null || cell === undefined) continue;

                if (typeof cell === 'object') {
                    if (cell._type === 'url') {
                        ws.writeUrl(r, c, new Url(`file:///${cell._path ?? ''}`).setText(cell._name ?? ''));
                    } else {
                        ws.write(r, c, String(cell));
                    }
                } else {
                    ws.write(r, c, cell);
                }
            }
        }

        return wb.saveToBufferSync();
    }
}
