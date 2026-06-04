class AdminConfigPage {
  constructor() {
    this.statusEl = document.getElementById("status");
    this.listEl = document.getElementById("fieldList");
    this.templateEl = document.getElementById("templatePath");
    this.datalist = document.getElementById("columnOptions");

    this.spSiteUrlEl = document.getElementById("spSiteUrl");
    this.spListTitleEl = document.getElementById("spListTitle");
    this.spMatchColumnEl = document.getElementById("spMatchColumn");
    this.spMatchFieldEl = document.getElementById("spMatchField");
    this.spListEl = document.getElementById("spFieldList");
    this.fieldsInfoEl = document.getElementById("fieldsInfo");

    this.config = null;
    this.spConfig = null;
    this.spFields = [];
    this._bindActions();
    this._init();
  }

  // Binding happens before async loads so the buttons always work even if a load fails.
  async _init() {
    try {
      await this._loadColumns();
      await this._loadConfig();
      await this._loadSpConfig();
    } catch (err) {
      console.error(err);
      this._status("שגיאה בטעינת ההגדרות — אפשר עדיין לערוך ולשמור", "error");
    }
  }

  // ---- Quick-print fields (tag <-> table column) ----

  async _loadColumns() {
    const data = await chrome.storage.local.get("summonsTableData");
    const headers = data.summonsTableData?.headers || [];
    this.datalist.innerHTML = "";
    headers.forEach((h) => {
      const opt = document.createElement("option");
      opt.value = h.label;
      this.datalist.appendChild(opt);
    });
  }

  async _loadConfig() {
    this.config = await QuickPrintConfigStore.get();
    this.templateEl.value = this.config.templatePath;
    this.listEl.innerHTML = "";
    this.config.fields.forEach((f) => this._addRow(f.tag, f.column));
    if (!this.config.fields.length) this._addRow("", "");
  }

  _addRow(tag = "", column = "") {
    const row = document.createElement("div");
    row.className = "field-row";

    const tagInput = this._input("field-input tag", "שם התג בוורד", tag);
    const colInput = this._input("field-input column", "שם העמודה בטבלה", column);
    colInput.setAttribute("list", "columnOptions");

    row.append(tagInput, colInput, this._removeBtn(row));
    this.listEl.appendChild(row);
  }

  // ---- SharePoint mapping (tag <-> SP field) ----

  async _loadSpConfig() {
    this.spConfig = await SharePointConfigStore.getRaw();
    this.spSiteUrlEl.value = this.spConfig.siteUrl;
    this.spListTitleEl.value = this.spConfig.listTitle;
    this.spMatchColumnEl.value = this.spConfig.matchTableColumn;
    this._populateSelect(this.spMatchFieldEl, this.spConfig.matchListFieldInternal);

    this.spListEl.innerHTML = "";
    const entries = Object.entries(this.spConfig.fieldMap || {});
    entries.forEach(([tag, internal]) => this._addSpRow(tag, internal));
    if (!entries.length) this._addSpRow("", "");
  }

  _addSpRow(tag = "", internal = "") {
    const row = document.createElement("div");
    row.className = "field-row";

    const tagInput = this._input("field-input sp-tag", "שם התג בוורד", tag);
    const select = document.createElement("select");
    select.className = "field-input sp-field";
    this._populateSelect(select, internal);

    row.append(tagInput, select, this._removeBtn(row));
    this.spListEl.appendChild(row);
  }

  // Fills a select with the loaded SP fields, keeping the currently selected internal name.
  _populateSelect(select, selectedInternal) {
    select.innerHTML = "";
    select.appendChild(new Option("— ללא —", ""));

    const known = new Set();
    this.spFields.forEach((f) => {
      known.add(f.internal);
      const opt = new Option(f.title, f.internal);
      if (f.internal === selectedInternal) opt.selected = true;
      select.appendChild(opt);
    });

    if (selectedInternal && !known.has(selectedInternal)) {
      const opt = new Option(selectedInternal, selectedInternal);
      opt.selected = true;
      select.appendChild(opt);
    }
  }

  async _loadFields() {
    const siteUrl = this.spSiteUrlEl.value.trim();
    const listTitle = this.spListTitleEl.value.trim();
    if (!siteUrl || !listTitle) {
      this._setFieldsInfo("יש למלא כתובת אתר וכותרת רשימה");
      return;
    }
    this._setFieldsInfo("טוען שדות...");
    try {
      const raw = await this._fetchFields(siteUrl, listTitle);
      this.spFields = raw
        .filter((f) => !f.Hidden)
        .map((f) => ({ title: f.Title || f.InternalName, internal: f.InternalName || f.StaticName }))
        .filter((f) => f.internal)
        .sort((a, b) => String(a.title).localeCompare(String(b.title), "he"));
      this._refreshSelects();
      this._setFieldsInfo(`נטענו ${this.spFields.length} שדות`);
    } catch (err) {
      console.error(err);
      this._setFieldsInfo(`טעינת השדות נכשלה — ${err?.message || "בדוק כתובת והרשאות"}`);
    }
  }

  async _fetchFields(siteUrl, listTitle) {
    try {
      return await this._fetchFieldsViaSharePointTab(siteUrl, listTitle);
    } catch (err) {
      console.warn("SharePoint tab fetch failed, falling back to extension fetch", err);
      const listUrl = SharePointConfigStore.buildListUrl(siteUrl, listTitle);
      return new SharePointLookupService({ listUrl }).fetchFields();
    }
  }

  async _fetchFieldsViaSharePointTab(siteUrl, listTitle) {
    const origin = new URL(siteUrl).origin;
    const tabs = await chrome.tabs.query({ url: `${origin}/*` });
    const tab = tabs.find((t) => t.url?.startsWith(siteUrl)) || tabs[0];
    if (!tab?.id) throw new Error("פתח לשונית SharePoint מחוברת ונסה שוב");

    const [{ result }] = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: async (url, title) => {
        const base = url.replace(/\/$/, "");
        const endpoint =
          `${base}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')` +
          "/fields?$select=Title,InternalName,StaticName,Hidden,TypeAsString&$format=json";
        const res = await fetch(endpoint, {
          credentials: "include",
          headers: { Accept: "application/json;odata=verbose" },
        });
        const text = await res.text();
        if (!res.ok) throw new Error(`SharePoint fields ${res.status}`);
        if (text.trim().startsWith("<")) throw new Error("SharePoint החזיר HTML במקום JSON");
        const json = JSON.parse(text);
        const body = json.d || json;
        return body.results || body.value || [];
      },
      args: [siteUrl, listTitle],
    });

    return result || [];
  }

  _refreshSelects() {
    this._populateSelect(this.spMatchFieldEl, this.spMatchFieldEl.value);
    this.spListEl.querySelectorAll(".sp-field").forEach((select) =>
      this._populateSelect(select, select.value)
    );
  }

  // ---- Shared helpers ----

  _input(className, placeholder, value) {
    const el = document.createElement("input");
    el.className = className;
    el.placeholder = placeholder;
    el.value = value;
    return el;
  }

  _removeBtn(row) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "btn btn--sm btn--danger";
    btn.textContent = "הסר";
    btn.addEventListener("click", () => row.remove());
    return btn;
  }

  _collectQuickPrint() {
    const fields = [];
    this.listEl.querySelectorAll(".field-row").forEach((row) => {
      const tag = row.querySelector(".tag").value.trim();
      const column = row.querySelector(".column").value.trim();
      if (tag) fields.push({ tag, column: column || tag });
    });
    return {
      templatePath: this.templateEl.value.trim() || "templates/template1.docx",
      outputMode: this.config?.outputMode || "browserPrint",
      fields,
    };
  }

  _collectSharePoint() {
    const fieldMap = {};
    this.spListEl.querySelectorAll(".field-row").forEach((row) => {
      const tag = row.querySelector(".sp-tag").value.trim();
      const internal = row.querySelector(".sp-field").value.trim();
      if (tag && internal) fieldMap[tag] = internal;
    });
    const matchOption = this.spMatchFieldEl.selectedOptions[0];
    return {
      siteUrl: this.spSiteUrlEl.value.trim(),
      listTitle: this.spListTitleEl.value.trim(),
      matchTableColumn: this.spMatchColumnEl.value.trim(),
      matchListFieldInternal: this.spMatchFieldEl.value,
      matchListFieldDisplay: matchOption ? matchOption.textContent : "",
      fieldMap,
    };
  }

  _bindActions() {
    document.getElementById("addField").addEventListener("click", () => this._addRow("", ""));
    document.getElementById("addSpField").addEventListener("click", () => this._addSpRow("", ""));
    document.getElementById("loadFields").addEventListener("click", () => this._loadFields());
    document.getElementById("save").addEventListener("click", () => this._save());
    document.getElementById("reset").addEventListener("click", () => this._reset());
    document.getElementById("exportSettings").addEventListener("click", () => this._export());
    document.getElementById("importSettings").addEventListener("click", () =>
      document.getElementById("importFile").click()
    );
    document.getElementById("importFile").addEventListener("change", (e) => {
      this._import(e.target.files[0]);
      e.target.value = "";
    });
  }

  // Downloads a JSON snapshot of the settings currently shown in the form.
  _export() {
    const data = {
      type: "summons-extension-settings",
      version: 1,
      exportedAt: new Date().toISOString(),
      quickPrint: this._collectQuickPrint(),
      sharePoint: this._collectSharePoint(),
    };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "הגדרות-תוסף-זימונים.json";
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    this._status("ההגדרות יוצאו לקובץ", "ok");
  }

  // Reads a settings file, persists it, and reloads the form.
  async _import(file) {
    if (!file) return;
    try {
      const data = JSON.parse(await file.text());
      if (data.quickPrint) await QuickPrintConfigStore.save(data.quickPrint);
      if (data.sharePoint) await SharePointConfigStore.save(data.sharePoint);
      await this._loadConfig();
      await this._loadSpConfig();
      this._status("ההגדרות יובאו והוחלו בהצלחה", "ok");
    } catch (err) {
      console.error(err);
      this._status("ייבוא נכשל — קובץ הגדרות לא תקין", "error");
    }
  }

  async _save() {
    try {
      await QuickPrintConfigStore.save(this._collectQuickPrint());
      await SharePointConfigStore.save(this._collectSharePoint());
      this._status("ההגדרות נשמרו בהצלחה", "ok");
    } catch (err) {
      console.error(err);
      this._status(err?.message || "שמירת ההגדרות נכשלה", "error");
    }
  }

  async _reset() {
    await QuickPrintConfigStore.reset();
    await SharePointConfigStore.reset();
    this.spFields = [];
    await this._loadConfig();
    await this._loadSpConfig();
    this._status("ההגדרות אופסו לברירת מחדל", "ok");
  }

  _setFieldsInfo(text) {
    this.fieldsInfoEl.textContent = text;
  }

  _status(text, type) {
    this.statusEl.textContent = text;
    this.statusEl.className = `visible ${type}`;
  }
}

new AdminConfigPage();
