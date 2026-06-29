class TemplateConfig {
  static normalize(entry = {}) {
    const legacy = (entry.path || "").trim();
    return {
      name: entry.name || "",
      printPath: (entry.printPath || legacy).trim(),
      exportPath: (entry.exportPath || legacy).trim(),
    };
  }

  static resolvePath(template, outputMode) {
    if (!template) return "";
    const isPrint = outputMode === "browserPrint";
    const chosen = isPrint ? template.printPath : template.exportPath;
    return (chosen || template.path || "").trim();
  }

  static hasAnyPath(template) {
    return Boolean(this.resolvePath(template, "browserPrint") || this.resolvePath(template, "download"));
  }

  static actionItems(template) {
    const items = [];
    if (this.resolvePath(template, "browserPrint")) {
      items.push({ label: "הדפסה", value: "browserPrint", variant: "" });
    }
    if (this.resolvePath(template, "download")) {
      items.push({ label: "ייצוא ל-Word", value: "download", variant: "export" });
    }
    return items;
  }
}

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
      itemPrintTemplates: [],
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
      [`${prefix}ColumnLayout`]: [],
      [`${prefix}TypeFieldInternal`]: "",
      [`${prefix}Types`]: [],
      [`${prefix}Templates`]: [],
      [`${prefix}LocationFieldInternal`]: "",
      [`${prefix}JudgeFieldInternal`]: "",
      [`${prefix}SummonedFieldInternal`]: "",
      [`${prefix}TemplateFieldMap`]: {},
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
      itemPrintTemplates: this.normalizeItemPrintTemplates(c.itemPrintTemplates),
    };
    this.PROFILES.forEach((p) => { config[p] = this._buildProfile(c, p); });
    return config;
  }

  // Shapes a secondary-list profile (summons / sessions) from its prefixed raw keys.
  static _buildProfile(c, prefix) {
    const columnLayout = this.resolveColumnLayout(c, prefix);
    return {
      siteUrl: c.siteUrl,
      listTitle: c[`${prefix}ListTitle`] || "",
      listUrl: this.buildListUrl(c.siteUrl, c[`${prefix}ListTitle`] || ""),
      dateField: c[`${prefix}DateFieldInternal`] || "",
      lookupField: c[`${prefix}LookupFieldInternal`] || "",
      columnLayout,
      displayFields: columnLayout.filter((col) => col.source === "list").map((col) => col.internal),
      baseDisplayFields: columnLayout.filter((col) => col.source === "base").map((col) => col.internal),
      basePosition: c[`${prefix}BasePosition`] || "before",
      typeField: c[`${prefix}TypeFieldInternal`] || "",
      types: this.normalizeTypes(c[`${prefix}Types`]),
      templates: this.normalizeTemplates(c[`${prefix}Templates`]),
      locationField: c[`${prefix}LocationFieldInternal`] || "",
      judgeField: c[`${prefix}JudgeFieldInternal`] || "",
      summonedField: c[`${prefix}SummonedFieldInternal`] || "",
      templateFieldMap: this.normalizeTemplateFieldMap(c[`${prefix}TemplateFieldMap`]),
    };
  }

  static normalizeColumnLayout(entries) {
    return (entries || [])
      .filter((entry) => entry?.internal && (entry.source === "base" || entry.source === "list"))
      .map((entry) => ({ source: entry.source === "base" ? "base" : "list", internal: entry.internal }));
  }

  static migrateColumnLayout(c, prefix) {
    const list = (c[`${prefix}DisplayFields`] || []).map((internal) => ({ source: "list", internal }));
    const base = (c[`${prefix}BaseDisplayFields`] || []).map((internal) => ({ source: "base", internal }));
    if (!list.length && !base.length) return [];
    if (!base.length) return list;
    if (!list.length) return base;
    return (c[`${prefix}BasePosition`] || "before") === "after" ? [...list, ...base] : [...base, ...list];
  }

  static resolveColumnLayout(c, prefix) {
    const saved = c[`${prefix}ColumnLayout`];
    if (Array.isArray(saved) && saved.length) return this.normalizeColumnLayout(saved);
    return this.migrateColumnLayout(c, prefix);
  }

  static normalizeTemplateFieldMap(map) {
    const systemFields = new Set([
      "printDay", "printMonth", "printYear",
      "gregorianDay", "gregorianMonth", "gregorianYear",
      "hebrewDay", "hebrewMonth", "hebrewYear",
      "printDate", "gregorianDate", "hebrewDate",
    ]);
    const out = {};
    Object.entries(map || {}).forEach(([tag, value]) => {
      if (!tag) return;
      const source = typeof value === "object" ? value.source : "list";
      const internal = typeof value === "object" ? value.internal : value;
      if (source === "system") {
        if (systemFields.has(internal)) out[tag] = { source: "system", internal };
        return;
      }
      if (internal) out[tag] = { source: source === "base" ? "base" : "list", internal };
    });
    return out;
  }

  static normalizeTemplates(templates) {
    return (templates || [])
      .map((t) => TemplateConfig.normalize(t))
      .filter((t) => TemplateConfig.hasAnyPath(t));
  }

  static normalizeItemPrintTemplates(templates) {
    return (templates || [])
      .map((t, index) => {
        const entry = TemplateConfig.normalize(t);
        return {
          id: t.id || `tpl_${index}`,
          name: entry.name,
          printPath: entry.printPath,
          exportPath: entry.exportPath,
          templateFieldMap: this.normalizeTemplateFieldMap(t.templateFieldMap),
        };
      })
      .filter((t) => TemplateConfig.hasAnyPath(t));
  }

  static normalizeTypes(types) {
    return (types || []).map((t) => ({
      name: t.name || "",
      templates: Array.isArray(t.templates) && t.templates.length
        ? t.templates.map((x) => TemplateConfig.normalize(x)).filter((x) => TemplateConfig.hasAnyPath(x))
        : (t.templatePath ? [TemplateConfig.normalize({ path: t.templatePath })] : []),
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
window.TemplateConfig = TemplateConfig;
