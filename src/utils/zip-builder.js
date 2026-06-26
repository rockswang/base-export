import JSZip from 'jszip';

export class ZipBuilder {
    constructor() {
        this.zip = new JSZip();
    }

    addXlsx(name, buffer) {
        this.zip.file(name, buffer);
        return this;
    }

    addAttachment(filePath, buffer) {
        this.zip.file(`attachments/${filePath}`, buffer);
        return this;
    }

    async generate() {
        return this.zip.generateAsync({ type: 'blob' });
    }
}
