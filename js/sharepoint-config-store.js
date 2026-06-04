class SharePointConfigStore {
  static KEY = "sharepointConfig";

  // Editable defaults derived from the static sharepoint-field-map.js file.
  static defaults() {
    const lookup = window.SHAREPOINT_LOOKUP || {};
    const join = lookup.join || {};
    return {
      siteUrl: lookup.siteUrl || "https://rabaz.army.idf/sites/Giur/Giuron",
      listTitle: lookup.listTitle || "ניהול מתגיירים",
      matchTableColumn: join.tableColumnDisplay || "שיוך לתהליך",
      matchListFieldInternal: join.listFieldInternal || "",
      matchListFieldDisplay: join.listFieldDisplay || "שם התהליך",
      fieldMap: { ...(window.SHAREPOINT_FIELD_MAP || {}) },
    };
  }

  // Raw saved values merged over defaults (what the admin window edits).
  static async getRaw() {
    const data = await chrome.storage.local.get(this.KEY);
    return { ...this.defaults(), ...(data[this.KEY] || {}) };
  }

  // Config shaped for SharePointLookupService.
  static async get() {
    const c = await this.getRaw();
    return {
      siteUrl: c.siteUrl,
      listTitle: c.listTitle,
      listUrl: this.buildListUrl(c.siteUrl, c.listTitle),
      join: {
        tableColumnDisplay: c.matchTableColumn,
        listFieldDisplay: c.matchListFieldDisplay,
        listFieldInternal: c.matchListFieldInternal,
      },
      fieldMap: c.fieldMap || {},
    };
  }

  static buildListUrl(siteUrl, listTitle) {
    const base = (siteUrl || "").replace(/\/$/, "");
    return `${base}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle || "")}')/items`;
  }

  static async save(raw) {
    await chrome.storage.local.set({ [this.KEY]: raw });
  }

  static async reset() {
    await chrome.storage.local.remove(this.KEY);
  }
}

window.SharePointConfigStore = SharePointConfigStore;
