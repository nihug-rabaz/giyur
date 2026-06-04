class SummonsExportService {
  constructor(headers) {
    this.headers = headers;
  }

  async generateDocx(templateId, row, rowIndex) {
    const buffer = await TemplateStore.getBuffer(templateId);
    const mapping = await TemplateStore.getMapping(templateId);
    const engine = new DocxTemplateEngine(buffer);
    return engine.fill(row, this.headers, mapping);
  }

  async downloadOne(templateId, row, rowIndex, templateNum) {
    const blob = await this.generateDocx(templateId, row, rowIndex);
    const name = this._rowFilename(row, rowIndex, templateNum);
    FileDownloadHelper.download(blob, `${name}.docx`);
  }

  async downloadAll(templateId, rows, templateNum) {
    for (let i = 0; i < rows.length; i++) {
      await this.downloadOne(templateId, rows[i], i, templateNum);
      await this._delay(350);
    }
  }

  async printOne(templateId, row, rowIndex) {
    const blob = await this.generateDocx(templateId, row, rowIndex);
    await DocxPrintRenderer.printBlobs([blob]);
  }

  async printAll(templateId, rows) {
    const blobs = [];
    for (let i = 0; i < rows.length; i++) {
      blobs.push(await this.generateDocx(templateId, rows[i], i));
    }
    await DocxPrintRenderer.printBlobs(blobs);
  }

  _rowFilename(row, rowIndex, templateNum) {
    const labelHeader = this.headers.find((h) => h.label.includes("שם")) || this.headers[0];
    const base = labelHeader ? row[labelHeader.id] : "";
    return FileDownloadHelper.sanitizeFilename(`${base || "זימון"}_${rowIndex + 1}_תבנית${templateNum}`);
  }

  _delay(ms) {
    return new Promise((r) => setTimeout(r, ms));
  }
}

window.SummonsExportService = SummonsExportService;
