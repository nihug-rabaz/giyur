class QuickPrintService {
  constructor(config, headers) {
    this.config = config || {};
    this.headers = headers || [];
    this.templateFieldMap = this.config.templateFieldMap || {};
    this.fields = this._normalizeFields(this.config.fields);
  }

  _normalizeFields(fields) {
    return (fields || []).map((field) => {
      if (typeof field === "string") return { tag: field, column: field, source: "column" };
      if (field.source === "system") {
        return { tag: field.tag, source: "system", internal: field.internal || "printDay" };
      }
      return { tag: field.tag, column: field.column ?? field.tag, source: "column" };
    });
  }

  // Maps every configured tag to a table column id, with auto-match as a fallback.
  buildMapping() {
    const mapping = {};
    this.fields.forEach(({ tag, column, source }) => {
      if (source === "system") {
        mapping[tag] = tag;
        return;
      }
      mapping[tag] = this._resolveColumnId(column, tag);
    });
    return mapping;
  }

  _enrichRow(row, date = new Date()) {
    let enriched = SystemDateService.applyTemplateFieldMap(row, this.templateFieldMap, date);
    this.fields.forEach(({ tag, source, internal }) => {
      if (source === "system") enriched[tag] = SystemDateService.valueOf(internal, date);
    });
    return enriched;
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
      blobs.push(await new DocxTemplateEngine(buffer).fill(this._enrichRow(row), this.headers, mapping));
    }
    return DocxPrintRenderer.printBlobs(blobs);
  }

  async _downloadDocx(rows, buffer, mapping) {
    if (rows.length === 1) {
      const blob = await new DocxTemplateEngine(buffer).fill(this._enrichRow(rows[0]), this.headers, mapping);
      FileDownloadHelper.download(blob, `${this._filename(rows[0], 0, rows.length)}.docx`);
      return;
    }

    const zip = new JSZip();
    for (let i = 0; i < rows.length; i++) {
      const blob = await new DocxTemplateEngine(buffer).fill(this._enrichRow(rows[i]), this.headers, mapping);
      zip.file(`${this._filename(rows[i], i, rows.length)}.docx`, await blob.arrayBuffer());
    }
    const out = await zip.generateAsync({ type: "blob" });
    FileDownloadHelper.download(out, this.config.exportZipName || "הדפסה_מהירה_תבנית_1.zip");
  }

  _filename(row, index, total = 1) {
    const base = FileDownloadHelper.sanitizeFilename(this.config.exportBaseName || "תבנית");
    if (total === 1) return base;
    return `${base}_${index + 1}`;
  }
}

window.QuickPrintService = QuickPrintService;
