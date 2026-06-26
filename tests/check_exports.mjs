import('wasm-xlsxwriter/web').then(m => {
    const keys = Object.keys(m).sort();
    console.log(keys.join('\n'));
});
