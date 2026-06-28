class ItemRow {
  constructor(page) {
    this.page = page;
    this.values = {};
    this.idValue = "";
    this.item = null;
    this.tr = document.createElement("tr");
    this._build();
  }

  _build() {
    const idTd = document.createElement("td");
    idTd.className = "cell-id";
    this.idInput = document.createElement("input");
    this.idInput.type = "text";
    this.idInput.placeholder = "מזהה";
    this.idInput.addEventListener("keydown", (e) => { if (e.key === "Enter") this.fill(); });
    this.idInput.addEventListener("blur", () => this.fill());
    idTd.appendChild(this.idInput);
    this.tr.appendChild(idTd);

    this.inputs = {};
    this.page.columns.forEach((c) => {
      const td = document.createElement("td");
      const input = document.createElement("input");
      input.type = "text";
      input.placeholder = c.tag;
      if (c.readOnly) input.readOnly = true;
      input.addEventListener("input", () => this._setRowStatus(this.item ? "נערך" : "לא נטען"));
      this.inputs[c.tag] = input;
      td.appendChild(input);
      this.tr.appendChild(td);
    });

    const statusTd = document.createElement("td");
    this.statusEl = document.createElement("span");
    this.statusEl.className = "row-status";
    statusTd.appendChild(this.statusEl);
    this.tr.appendChild(statusTd);

    const actionsTd = document.createElement("td");
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "btn btn--sm btn--danger";
    remove.textContent = "הסר";
    remove.addEventListener("click", () => this.page.removeRow(this));
    actionsTd.appendChild(remove);
    const save = document.createElement("button");
    save.type = "button";
    save.className = "btn btn--sm btn--success";
    save.textContent = "שמור";
    save.addEventListener("click", () => this.save());
    actionsTd.appendChild(save);
    const create = document.createElement("button");
    create.type = "button";
    create.className = "btn btn--sm btn--primary";
    create.textContent = "צור";
    create.addEventListener("click", () => this.create());
    actionsTd.appendChild(create);
    this.tr.appendChild(actionsTd);
  }

  // Looks the entered id up in the list and writes the matched item values into the row cells.
  async fill() {
    const id = this.idInput.value.trim();
    if (!id || id === this.idValue) return;
    this.idValue = id;
    this._setRowStatus("טוען...");
    try {
      const matches = await this.page.lookup.search(this.page.searchInternal(), id);
      const item = matches[0];
      this.item = item || null;
      this.page.columns.forEach((c) => {
        const raw = item ? this.page.lookup.valueOf(item, c.internal) : "";
        const value = SharePointLookupService.formatValue(raw);
        this.values[c.tag] = value;
        this.inputs[c.tag].value = value;
      });
      this._setRowStatus(item ? "נטען" : "לא נמצא");
    } catch (err) {
      console.error(err);
      this._setRowStatus("שגיאה");
      this.page._setStatus(err?.message || "חיפוש נכשל", "error");
    }
  }

  async save() {
    if (!this.item) {
      await this.fill();
      if (!this.item) return this._setRowStatus("לא נמצא פריט לשמירה");
    }
    const changes = this._changes();
    if (!Object.keys(changes).length) return this._setRowStatus("אין שינויים");
    this._setRowStatus("שומר...");
    try {
      await this.page.lookup.update(this.item, changes);
      this.page.columns.forEach((c) => { this.values[c.tag] = this.inputs[c.tag].value; });
      this._setRowStatus("נשמר");
    } catch (err) {
      console.error(err);
      this._setRowStatus("שמירה נכשלה");
      this.page._setStatus(`שמירה נכשלה — ${err?.message || "בדוק הרשאות ושדות"}`, "error");
    }
  }

  // Creates a brand-new list item from the row inputs (id input maps to the chosen search field).
  async create() {
    const payload = this._createPayload();
    if (!Object.keys(payload).length) return this._setRowStatus("אין נתונים ליצירה");
    this._setRowStatus("יוצר...");
    try {
      const created = await this.page.lookup.create(payload);
      this.item = created;
      this.page.columns.forEach((c) => { this.values[c.tag] = this.inputs[c.tag].value; });
      const newId = created?.id ?? created?.ID ?? created?.Id;
      if (newId !== undefined) this.idValue = String(newId);
      this._setRowStatus(newId !== undefined ? `נוצר (ID ${newId})` : "נוצר");
    } catch (err) {
      console.error(err);
      this._setRowStatus("יצירה נכשלה");
      this.page._setStatus(`יצירה נכשלה — ${err?.message || "בדוק הרשאות ושדות"}`, "error");
    }
  }

  _createPayload() {
    const payload = {};
    const idField = this.page.searchInternal();
    const idVal = this.idInput.value.trim();
    if (idField && idVal) payload[idField] = idVal;
    this.page.columns.forEach((c) => {
      const value = this.inputs[c.tag].value.trim();
      if (value !== "") payload[c.internal] = value;
    });
    return payload;
  }

  _changes() {
    const changes = {};
    this.page.columns.forEach((c) => {
      const next = this.inputs[c.tag].value.trim();
      const prev = String(this.values[c.tag] ?? "").trim();
      if (next !== prev) changes[c.internal] = next;
    });
    return changes;
  }

  exportRow(idLabel) {
    const row = { [idLabel]: this.idInput.value.trim() };
    this.page.columns.forEach((c) => { row[c.tag] = this.inputs[c.tag]?.value ?? ""; });
    return row;
  }

  _setRowStatus(text) { this.statusEl.textContent = text; }
}

class NewItemForm {
  constructor(page) {
    this.page = page;
    this.overlay = null;
    this.inputs = {};
  }

  // Opens a modal listing every writable field of the configured list for a fresh item.
  open() {
    const fields = this.page.lookup.creatableFields();
    if (!fields.length) return this.page._setStatus("אין שדות זמינים — טען שדות מהרשימה", "error");
    this._render(fields);
  }

  _render(fields) {
    this.fields = fields;
    this.overlay = this._el("div", "modal-overlay");
    this.overlay.addEventListener("click", (e) => { if (e.target === this.overlay) this.close(); });

    const modal = this._el("div", "modal");
    modal.appendChild(this._header());
    modal.appendChild(this._body(fields));
    this.statusEl = this._el("div", "modal-status");
    modal.appendChild(this.statusEl);
    modal.appendChild(this._footer());

    this.overlay.appendChild(modal);
    document.body.appendChild(this.overlay);
    this._escHandler = (e) => { if (e.key === "Escape") this.close(); };
    document.addEventListener("keydown", this._escHandler);
    modal.querySelector("input")?.focus();
  }

  _header() {
    const header = this._el("div", "modal-header");
    const title = this._el("h2", "modal-title");
    title.textContent = "יצירת פריט חדש";
    const close = this._el("button", "modal-close");
    close.type = "button";
    close.textContent = "✕";
    close.addEventListener("click", () => this.close());
    header.appendChild(title);
    header.appendChild(close);
    return header;
  }

  _body(fields) {
    const body = this._el("div", "modal-body");
    const grid = this._el("div", "form-grid");
    fields.forEach((f) => {
      const wrap = this._el("div", `form-field${f.required ? " is-required" : ""}`);
      const label = document.createElement("label");
      label.textContent = f.title;
      const control = this._control(f);
      this.inputs[f.internal] = control;
      wrap.appendChild(label);
      wrap.appendChild(control);
      grid.appendChild(wrap);
    });
    body.appendChild(grid);
    return body;
  }

  _control(f) {
    if (f.type === "Lookup" || f.type === "LookupMulti") return this._lookupSelect(f);
    const input = document.createElement("input");
    input.type = "text";
    input.placeholder = f.type || "";
    input.required = Boolean(f.required);
    return input;
  }

  // Renders a lookup field as a dropdown, filled asynchronously with the target list's values.
  _lookupSelect(f) {
    const select = document.createElement("select");
    select.required = Boolean(f.required);
    if (f.type === "LookupMulti") select.multiple = true;
    const loading = document.createElement("option");
    loading.value = "";
    loading.textContent = "טוען ערכים…";
    select.appendChild(loading);
    this.page.lookup.lookupOptions(f)
      .then((options) => this._fillSelect(select, options, f))
      .catch(() => { loading.textContent = "שגיאה בטעינת ערכים"; });
    return select;
  }

  _fillSelect(select, options, f) {
    select.innerHTML = "";
    if (f.type !== "LookupMulti") {
      const blank = document.createElement("option");
      blank.value = "";
      blank.textContent = f.required ? "— בחר —" : "— ללא —";
      select.appendChild(blank);
    }
    options
      .slice()
      .sort((a, b) => String(a.label).localeCompare(String(b.label), "he"))
      .forEach((o) => {
        const opt = document.createElement("option");
        opt.value = String(o.id);
        opt.textContent = o.label || String(o.id);
        select.appendChild(opt);
      });
  }

  _valueOf(f) {
    const el = this.inputs[f.internal];
    if (!el) return "";
    if (el.multiple) return Array.from(el.selectedOptions).map((o) => o.value).filter(Boolean).join(",");
    return el.value.trim();
  }

  _footer() {
    const footer = this._el("div", "modal-footer");
    const submit = this._el("button", "btn btn--primary");
    submit.type = "button";
    submit.textContent = "צור פריט";
    submit.addEventListener("click", () => this._submit(submit));
    const cancel = this._el("button", "btn");
    cancel.type = "button";
    cancel.textContent = "ביטול";
    cancel.addEventListener("click", () => this.close());
    footer.appendChild(submit);
    footer.appendChild(cancel);
    return footer;
  }

  _payload() {
    const payload = {};
    (this.fields || []).forEach((f) => {
      const value = this._valueOf(f);
      if (value) payload[f.entityProperty || f.internal] = value;
    });
    return payload;
  }

  async _submit(btn) {
    const missing = (this.fields || []).filter((f) => f.required && !this._valueOf(f));
    if (missing.length) return this._setStatus(`חסרים שדות חובה: ${missing.map((f) => f.title).join(", ")}`, "error");
    const payload = this._payload();
    if (!Object.keys(payload).length) return this._setStatus("מלא לפחות שדה אחד", "error");
    btn.disabled = true;
    this._setStatus("יוצר פריט...", "info");
    try {
      const created = await this.page.lookup.create(payload);
      const skipped = created?._skipped || [];
      if (skipped.length) {
        this._setStatus(`נוצר, אך לא נשמרו שדות (ערך לא נמצא): ${skipped.join(", ")}`, "error");
        this.page._onItemCreated(created?.id ?? created?.ID ?? created?.Id);
        return;
      }
      this.page._onItemCreated(created?.id ?? created?.ID ?? created?.Id);
      this.close();
    } catch (err) {
      console.error(err);
      this._setStatus(`יצירה נכשלה — ${err?.message || "בדוק הרשאות ושדות"}`, "error");
      btn.disabled = false;
    }
  }

  _setStatus(text, type = "") {
    this.statusEl.textContent = text;
    this.statusEl.className = `modal-status ${type}`;
  }

  close() {
    if (this._escHandler) document.removeEventListener("keydown", this._escHandler);
    this.overlay?.remove();
    this.overlay = null;
  }

  _el(tag, className) {
    const el = document.createElement(tag);
    el.className = className;
    return el;
  }
}

class SharePointListBuilderPage {
  constructor() {
    this.lookup = new SharePointItemLookup();
    this.columns = [];
    this.rows = [];
    this.fieldSelect = document.getElementById("searchField");
    this.statusEl = document.getElementById("status");
    this.tableHead = document.getElementById("tableHead");
    this.tableBody = document.getElementById("tableBody");
    this.emptyState = document.getElementById("emptyState");
    this._bind();
    this._init();
  }

  _bind() {
    document.getElementById("reloadFields").addEventListener("click", () => this._loadFields());
    document.getElementById("addRow").addEventListener("click", () => this.addRow());
    document.getElementById("clearRows").addEventListener("click", () => this._clear());
    document.getElementById("saveAll").addEventListener("click", () => this._saveAll());
    document.getElementById("createAll").addEventListener("click", () => this._createAll());
    document.getElementById("newItem").addEventListener("click", () => new NewItemForm(this).open());
    this.fieldSelect.addEventListener("change", () => this._updateIdHeader());
    document.getElementById("exportT1").addEventListener("click", () => this._print(1));
    document.getElementById("exportT2").addEventListener("click", () => this._print(2));
    document.getElementById("exportT3").addEventListener("click", () => this._print(3));
  }

  searchInternal() { return this.fieldSelect.value; }

  async _init() {
    try {
      const config = await this.lookup.init();
      if (!config.listTitle && !config.listUrl) {
        this._setStatus("לא הוגדרה רשימת SharePoint — הגדר אותה במסך ההגדרות", "error");
        return;
      }
      await this._loadFields();
      if (!this.rows.length) this.addRow();
    } catch (err) {
      this._setStatus(err?.message || "שגיאה בטעינה", "error");
    }
  }

  async _loadFields() {
    this._setStatus("טוען שדות מהרשימה...", "info");
    try {
      const fields = await this.lookup.loadFields();
      this._populateFieldSelect(fields);
      this._computeColumns();
      this._renderHead();
      this._rebuildRows();
      this._setStatus(`נטענו ${fields.length} שדות. הוסף שורות והזן מזהים.`, "ok");
    } catch (err) {
      this._setStatus(`טעינת שדות נכשלה — ${err?.message || "בדוק הגדרות והרשאות"}`, "error");
    }
  }

  _populateFieldSelect(fields) {
    this.fieldSelect.innerHTML = "";
    fields.forEach((f) => this.fieldSelect.appendChild(new Option(f.title, f.internal)));
    const match = this.lookup.config?.join?.listFieldInternal;
    if (match && fields.some((f) => f.internal === match)) this.fieldSelect.value = match;
  }

  // Columns honor the admin-selected display fields (empty = all). Word-tagged fields keep their
  // export tag and are always included so printing keeps working.
  _computeColumns() {
    const byInternal = new Map(this.lookup.fields.map((f) => [f.internal, f]));
    const tagByInternal = new Map(
      Object.entries(this.lookup.config?.fieldMap || {}).map(([tag, internal]) => [internal, tag])
    );
    const display = (this.lookup.config?.displayFields || [])
      .filter((id) => byInternal.has(id) || tagByInternal.has(id));

    const chosen = display.length ? [...display] : this.lookup.fields.map((f) => f.internal);
    tagByInternal.forEach((_tag, internal) => { if (!chosen.includes(internal)) chosen.push(internal); });

    const seenTags = new Set();
    this.columns = chosen
      .map((internal) => {
        const f = byInternal.get(internal);
        return {
          tag: tagByInternal.get(internal) || f?.title || internal,
          internal,
          readOnly: Boolean(f?.readOnly),
        };
      })
      .filter((c) => {
        if (seenTags.has(c.tag)) return false;
        seenTags.add(c.tag);
        return true;
      });
  }

  _idLabel() {
    const field = this.lookup.fields.find((f) => f.internal === this.searchInternal());
    return field ? field.title : "מזהה";
  }

  _renderHead() {
    this.tableHead.innerHTML = "";
    const tr = document.createElement("tr");
    const idTh = document.createElement("th");
    idTh.id = "idHeader";
    idTh.textContent = `מזהה (${this._idLabel()})`;
    tr.appendChild(idTh);
    this.columns.forEach((c) => {
      const th = document.createElement("th");
      th.textContent = c.tag;
      tr.appendChild(th);
    });
    ["סטטוס", "פעולות"].forEach((t) => {
      const th = document.createElement("th");
      th.textContent = t;
      tr.appendChild(th);
    });
    this.tableHead.appendChild(tr);
  }

  _updateIdHeader() {
    const th = document.getElementById("idHeader");
    if (th) th.textContent = `מזהה (${this._idLabel()})`;
  }

  addRow(idValue = "") {
    const row = new ItemRow(this);
    if (idValue) row.idInput.value = idValue;
    this.rows.push(row);
    this.tableBody.appendChild(row.tr);
    this.emptyState.hidden = true;
    if (idValue) row.fill();
    else row.idInput.focus();
    return row;
  }

  removeRow(row) {
    row.tr.remove();
    this.rows = this.rows.filter((r) => r !== row);
    this.emptyState.hidden = this.rows.length > 0;
  }

  _clear() {
    this.rows.forEach((r) => r.tr.remove());
    this.rows = [];
    this.emptyState.hidden = false;
  }

  // Recreates rows after a field reload, preserving the entered ids and refreshing values.
  _rebuildRows() {
    const ids = this.rows.map((r) => r.idInput.value.trim());
    this.rows = [];
    this.tableBody.innerHTML = "";
    ids.forEach((id) => this.addRow(id));
    this.emptyState.hidden = this.rows.length > 0;
  }

  async _print(templateNum) {
    await Promise.all(this.rows.map((r) => r.fill()));
    const idLabel = this._idLabel();
    const rows = this.rows.filter((r) => r.idInput.value.trim()).map((r) => r.exportRow(idLabel));
    if (!rows.length) return this._setStatus("אין שורות עם מזהה להדפסה", "error");

    this._setStatus("מכין הדפסה...", "info");
    try {
      const templatePath = await this._templatePath(templateNum);
      const cfg = templateNum === 1 ? await QuickPrintConfigStore.get() : { fields: [] };
      const service = new QuickPrintService(
        { templatePath, outputMode: "browserPrint", fields: cfg.fields || [] },
        this._exportHeaders(idLabel)
      );
      await service.printAll(rows);
      this._setStatus(`נשלח להדפסה (תבנית ${templateNum})`, "ok");
    } catch (err) {
      this._setStatus(`הדפסה נכשלה — ${err?.message || "בדוק את התבנית"}`, "error");
    }
  }

  async _saveAll() {
    const targets = this.rows.filter((r) => r.idInput.value.trim());
    if (!targets.length) return this._setStatus("אין שורות לשמירה", "error");
    this._setStatus("שומר שינויים...", "info");
    await Promise.all(targets.map((r) => r.save()));
    this._setStatus("שמירה הסתיימה", "ok");
  }

  // Confirms a modal creation; the cache was invalidated so later searches include the new item.
  _onItemCreated(newId) {
    this._setStatus(newId !== undefined ? `נוצר פריט חדש (ID ${newId})` : "פריט חדש נוצר", "ok");
  }

  async _createAll() {
    const targets = this.rows.filter((r) => !r.item && Object.keys(r._createPayload()).length);
    if (!targets.length) return this._setStatus("אין שורות חדשות ליצירה", "error");
    this._setStatus("יוצר פריטים...", "info");
    await Promise.all(targets.map((r) => r.create()));
    this._setStatus("יצירה הסתיימה", "ok");
  }

  async _templatePath(num) {
    if (num === 1) {
      const cfg = await QuickPrintConfigStore.get();
      return cfg.templatePath || "templates/template1.docx";
    }
    return `templates/template${num}.docx`;
  }

  _exportHeaders(idLabel) {
    return [{ id: idLabel, label: idLabel }, ...this.columns.map((c) => ({ id: c.tag, label: c.tag }))];
  }

  _setStatus(text, type = "") {
    this.statusEl.textContent = text;
    this.statusEl.className = type ? `visible ${type}` : "visible";
  }
}

new SharePointListBuilderPage();
