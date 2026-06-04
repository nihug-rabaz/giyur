class TemplateStore {
  static async getBuffer(templateId) {
    const custom = await this._getCustom(templateId);
    if (custom) return custom;
    const url = chrome.runtime.getURL(`templates/${templateId}.docx`);
    const res = await fetch(url);
    if (!res.ok) throw new Error("לא ניתן לטעון את התבנית");
    return res.arrayBuffer();
  }

  static async saveCustom(templateId, arrayBuffer) {
    const bytes = Array.from(new Uint8Array(arrayBuffer));
    await chrome.storage.local.set({ [this._customKey(templateId)]: bytes });
  }

  static async clearCustom(templateId) {
    await chrome.storage.local.remove(this._customKey(templateId));
  }

  static async getMapping(templateId) {
    const key = this._mappingKey(templateId);
    const data = await chrome.storage.local.get(key);
    return data[key] || {};
  }

  static async saveMapping(templateId, mapping) {
    await chrome.storage.local.set({ [this._mappingKey(templateId)]: mapping });
  }

  static _customKey(templateId) { return `custom_${templateId}`; }

  static _mappingKey(templateId) { return `mapping_${templateId}`; }

  static async _getCustom(templateId) {
    const key = this._customKey(templateId);
    const data = await chrome.storage.local.get(key);
    const bytes = data[key];
    if (!bytes?.length) return null;
    return new Uint8Array(bytes).buffer;
  }
}

window.TemplateStore = TemplateStore;
