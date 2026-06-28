// Aggregates every persisted setting into one portable snapshot and restores it.
class SettingsBackupStore {
  static TYPE = "summons-extension-settings";
  static VERSION = 2;
  static RUNTIME_KEYS = new Set(["summonsTableData"]);

  static _isTemplateKey(key) {
    return key.startsWith("custom_") || key.startsWith("mapping_");
  }

  static _mergeConfig(stored, override) {
    if (!stored && !override) return null;
    return { ...(stored || {}), ...(override || {}) };
  }

  static _templatesFrom(all) {
    const templates = {};
    Object.keys(all)
      .filter((key) => this._isTemplateKey(key))
      .forEach((key) => { templates[key] = all[key]; });
    return templates;
  }

  static _extraFrom(all) {
    const extra = {};
    Object.keys(all).forEach((key) => {
      if (this.RUNTIME_KEYS.has(key)) return;
      if (key === QuickPrintConfigStore.KEY) return;
      if (key === SharePointConfigStore.KEY) return;
      if (this._isTemplateKey(key)) return;
      extra[key] = all[key];
    });
    return extra;
  }

  // Confirms every non-runtime storage key appears in the export payload.
  static verifyPayload(payload, all = {}) {
    const missing = [];
    Object.keys(all).forEach((key) => {
      if (this.RUNTIME_KEYS.has(key)) return;
      if (key === QuickPrintConfigStore.KEY) {
        if (!payload.quickPrint) missing.push(key);
        return;
      }
      if (key === SharePointConfigStore.KEY) {
        if (!payload.sharePoint) missing.push(key);
        return;
      }
      if (this._isTemplateKey(key)) {
        if (!payload.templates?.[key]) missing.push(key);
        return;
      }
      if (!payload.extra?.[key]) missing.push(key);
    });
    return { complete: missing.length === 0, missing };
  }

  // Builds a full snapshot from storage; `overrides` lets the admin form inject unsaved edits.
  static async export(overrides = {}) {
    const all = await chrome.storage.local.get(null);
    const templates = this._templatesFrom(all);
    const extra = this._extraFrom(all);
    const sharePoint = this._mergeConfig(all[SharePointConfigStore.KEY], overrides.sharePoint);
    const quickPrint = this._mergeConfig(all[QuickPrintConfigStore.KEY], overrides.quickPrint);
    const payload = {
      type: this.TYPE,
      version: this.VERSION,
      exportedAt: new Date().toISOString(),
      quickPrint,
      sharePoint,
      templates,
    };
    if (Object.keys(extra).length) payload.extra = extra;
    const check = this.verifyPayload(payload, all);
    if (!check.complete) console.warn("Settings export missing keys:", check.missing);
    return payload;
  }

  // Restores a snapshot produced by export(), writing only recognized settings keys.
  static async import(data) {
    if (!data || typeof data !== "object") throw new Error("קובץ הגדרות לא תקין");
    const writes = {};
    if (data.quickPrint) writes[QuickPrintConfigStore.KEY] = data.quickPrint;
    if (data.sharePoint) writes[SharePointConfigStore.KEY] = data.sharePoint;
    if (data.templates && typeof data.templates === "object") {
      Object.entries(data.templates)
        .filter(([key]) => this._isTemplateKey(key))
        .forEach(([key, value]) => { writes[key] = value; });
    }
    if (data.extra && typeof data.extra === "object") {
      Object.entries(data.extra).forEach(([key, value]) => { writes[key] = value; });
    }
    if (!Object.keys(writes).length) throw new Error("קובץ הגדרות לא תקין");
    await chrome.storage.local.set(writes);
  }
}

window.SettingsBackupStore = SettingsBackupStore;
