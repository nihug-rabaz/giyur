class SharePointConfigStore {
  static KEY = "sharepointConfig";

  // Secondary related lists configured in the admin window (each is an independent screen).
  static PROFILES = ["summons", "sessions"];

  // Editable defaults derived from the static sharepoint-field-map.js file.
  static defaults() {
    const lookup = window.SHAREPOINT_LOOKUP || {};
    const join = lookup.join || {};
    const base = {
      siteUrl: lookup.siteUrl || "https://rabaz.army.idf/sites/Giur/Giuron",
      listTitle: lookup.listTitle || "ניהול מתגיירים",
      matchTableColumn: join.tableColumnDisplay || "שיוך לתהליך (מ.א מתגייר/ת)",
      matchListFieldInternal: join.listFieldInternal || "Title",
      matchListFieldDisplay: join.listFieldDisplay || "שם התהליך (מ.א מתגייר/ת)",
      fieldMap: { ...(window.SHAREPOINT_FIELD_MAP || {}) },
      displayFields: [...(window.SHAREPOINT_DISPLAY_FIELDS || [])],
    };
    this.PROFILES.forEach((p) => Object.assign(base, this._profileDefaults(p)));
    return base;
  }

  static _profileDefaults(prefix) {
    return {
      [`${prefix}ListTitle`]: "",
      [`${prefix}DateFieldInternal`]: "",
      [`${prefix}LookupFieldInternal`]: "",
      [`${prefix}DisplayFields`]: [],
      [`${prefix}BaseDisplayFields`]: [],
      [`${prefix}BasePosition`]: "before",
      [`${prefix}TypeFieldInternal`]: "",
      [`${prefix}Types`]: [],
      [`${prefix}Templates`]: [],
      [`${prefix}LocationFieldInternal`]: "",
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
    const config = {
      siteUrl: c.siteUrl,
      listTitle: c.listTitle,
      listUrl: this.buildListUrl(c.siteUrl, c.listTitle),
      join: {
        tableColumnDisplay: c.matchTableColumn,
        listFieldDisplay: c.matchListFieldDisplay,
        listFieldInternal: c.matchListFieldInternal,
      },
      fieldMap: c.fieldMap || {},
      displayFields: c.displayFields || [],
    };
    this.PROFILES.forEach((p) => { config[p] = this._buildProfile(c, p); });
    return config;
  }

  // Shapes a secondary-list profile (summons / sessions) from its prefixed raw keys.
  static _buildProfile(c, prefix) {
    return {
      siteUrl: c.siteUrl,
      listTitle: c[`${prefix}ListTitle`] || "",
      listUrl: this.buildListUrl(c.siteUrl, c[`${prefix}ListTitle`] || ""),
      dateField: c[`${prefix}DateFieldInternal`] || "",
      lookupField: c[`${prefix}LookupFieldInternal`] || "",
      displayFields: c[`${prefix}DisplayFields`] || [],
      baseDisplayFields: c[`${prefix}BaseDisplayFields`] || [],
      basePosition: c[`${prefix}BasePosition`] || "before",
      typeField: c[`${prefix}TypeFieldInternal`] || "",
      types: this.normalizeTypes(c[`${prefix}Types`]),
      templates: this.normalizeTemplates(c[`${prefix}Templates`]),
      locationField: c[`${prefix}LocationFieldInternal`] || "",
    };
  }

  // Flat print templates (buttons without tabs): [{ name, path }].
  static normalizeTemplates(templates) {
    return (templates || [])
      .map((t) => ({ name: t.name || "", path: t.path || "" }))
      .filter((t) => t.path);
  }

  // Normalizes summons types to { name, templates: [{ name, path }] }, upgrading the old single-path shape.
  static normalizeTypes(types) {
    return (types || []).map((t) => ({
      name: t.name || "",
      templates: Array.isArray(t.templates) && t.templates.length
        ? t.templates.map((x) => ({ name: x.name || "", path: x.path || "" })).filter((x) => x.path)
        : (t.templatePath ? [{ name: "", path: t.templatePath }] : []),
    }));
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
