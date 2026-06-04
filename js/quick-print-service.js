class QuickPrintService {
  constructor(config, headers) {
    this.config = config || {};
    this.headers = headers || [];
    this.fields = this._normalizeFields(this.config.fields);
  }

  _normalizeFields(fields) {
    return (fields || []).map((f) =>
      typeof f === "string" ? { tag: f, column: f } : { tag: f.tag, column: f.column ?? f.tag }
    );
  }

  // Maps every configured tag to a table column id, with auto-match as a fallback.
  buildMapping() {
    const mapping = {};
    this.fields.forEach(({ tag, column }) => {
      mapping[tag] = this._resolveColumnId(column, tag);
    });
    return mapping;
  }

  _resolveColumnId(column, tag) {
    const exact = this.headers.find((h) => h.label === column);
    if (exact) return exact.id;
    const partial = this.headers.find((h) => column && h.label.includes(column));
    if (partial) return partial.id;
    return window.PlaceholderMapper ? PlaceholderMapper.autoMatch(tag, this.headers) : "";
  }

  async loadTemplate() {
    const url = chrome.runtime.getURL(this.config.templatePath);
    const res = await fetch(url);
    if (!res.ok) throw new Error("לא ניתן לטעון את תבנית ההדפסה המהירה");
    return res.arrayBuffer();
  }

  // Keeps the original DOCX design by generating real Word files from the configured template.
  async printAll(rows) {
    const buffer = await this.loadTemplate();
    const mapping = this.buildMapping();
    if (this.config.outputMode === "browserPrint") {
      return this._browserPrint(rows, buffer, mapping);
    }
    return this._downloadDocx(rows, buffer, mapping);
  }

  async _browserPrint(rows, buffer, mapping) {
    const blobs = [];
    for (const row of rows) {
      blobs.push(await new DocxTemplateEngine(buffer).fill(row, this.headers, mapping));
    }
    return DocxPrintRenderer.printBlobs(blobs);
  }

  async _downloadDocx(rows, buffer, mapping) {
    if (rows.length === 1) {
      const blob = await new DocxTemplateEngine(buffer).fill(rows[0], this.headers, mapping);
      FileDownloadHelper.download(blob, `${this._filename(rows[0], 0)}.docx`);
      return;
    }

    const zip = new JSZip();
    for (let i = 0; i < rows.length; i++) {
      const blob = await new DocxTemplateEngine(buffer).fill(rows[i], this.headers, mapping);
      zip.file(`${this._filename(rows[i], i)}.docx`, await blob.arrayBuffer());
    }
    const out = await zip.generateAsync({ type: "blob" });
    FileDownloadHelper.download(out, "הדפסה_מהירה_תבנית_1.zip");
  }

  _filename(row, index) {
    const labelHeader = this.headers.find((h) => h.label.includes("שם")) || this.headers[0];
    const base = labelHeader ? row[labelHeader.id] : "";
    return FileDownloadHelper.sanitizeFilename(`${base || "זימון"}_${index + 1}_תבנית1`);
  }
}

window.QuickPrintService = QuickPrintService;
