import('wasm-xlsxwriter/web').then(m => {
    console.log('=== FormatAlign ===');
    console.log(Object.entries(m.FormatAlign).map(([k,v]) => `  ${k}=${v}`).join('\n'));
    console.log('=== FormatBorder ===');
    console.log(Object.entries(m.FormatBorder).map(([k,v]) => `  ${k}=${v}`).join('\n'));
    console.log('=== Color ===');
    console.log(Object.keys(m.Color).filter(k => isNaN(Number(k))).map(k => `  ${k}`).join('\n'));
});
