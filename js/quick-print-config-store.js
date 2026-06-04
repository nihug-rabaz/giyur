class QuickPrintConfigStore {
  static KEY = "quickPrintConfig";

  // Static defaults declared in quick-print-config.js, normalized to {tag, column} fields.
  static defaults() {
    const base = window.QUICK_PRINT_CONFIG || {};
    return {
      templatePath: base.templatePath || "templates/template1.docx",
      outputMode: base.outputMode || "browserPrint",
      fields: (base.fields || []).map((f) =>
        typeof f === "string" ? { tag: f, column: f } : { tag: f.tag, column: f.column ?? f.tag }
      ),
    };
  }

  // Saved admin config merged over the static defaults.
  static async get() {
    const def = this.defaults();
    const data = await chrome.storage.local.get(this.KEY);
    const saved = data[this.KEY];
    if (!saved) return def;
    return {
      templatePath: saved.templatePath || def.templatePath,
      outputMode: saved.outputMode || def.outputMode,
      fields: Array.isArray(saved.fields) ? saved.fields : def.fields,
    };
  }

  static async save(config) {
    await chrome.storage.local.set({ [this.KEY]: config });
  }

  static async reset() {
    await chrome.storage.local.remove(this.KEY);
  }
}

window.QuickPrintConfigStore = QuickPrintConfigStore;
