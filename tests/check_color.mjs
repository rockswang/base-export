import('wasm-xlsxwriter/web').then(m => {
    console.log('Color type:', typeof m.Color);
    if (typeof m.Color === 'function') {
        const proto = Object.getOwnPropertyNames(m.Color.prototype);
        console.log('Color prototype methods:', proto.join(', '));
        const staticMethods = Object.getOwnPropertyNames(m.Color).filter(k => k !== 'prototype' && k !== 'length' && k !== 'name');
        console.log('Color static:', staticMethods.join(', '));
    }
    // Try to create a Color
    try { console.log('Color.White:', m.Color.White); } catch(e) {}
    // Check if it accepts hex strings
    console.log('Format.setBackgroundColor(0x4472C4):');
    const fmt = new m.Format();
    fmt.setBackgroundColor(0x4472C4);
    console.log('  OK');
});
