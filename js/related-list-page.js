// Generic screen for a related SharePoint list: type tabs, date-range + location filtering,
// optional base-list join, Word report, and per-template printing. Configured by a profile key.
class RelatedListPage {
  constructor(options = {}) {
    this.profileKey = options.profileKey || "summons";
    this.noun = options.noun || "פריטים";
    this.listService = null;
    this.baseService = null;
    this.baseFields = [];
    this.config = null;
    this.profile = null;
    this.listColumns = [];
    this.baseColumns = [];
    this.columns = [];
    this.rows = [];
    this.rowLocations = [];
    this.locationOptions = [];
    this.locationLabels = new Map();
    this.activeType = null;
    this.dateFromEl = document.getElementById("dateFrom");
    this.dateToEl = document.getElementById("dateTo");
    this.peopleModeEl = document.getElementById("peopleMode");
    this.typeTabsEl = document.getElementById("typeTabs");
    this.locationFilterEl = document.getElementById("locationFilter");
    this.printButtonsEl = document.getElementById("printButtons");
    this.statusEl = document.getElementById("status");
    this.tableHead = document.getElementById("tableHead");
    this.tableBody = document.getElementById("tableBody");
    this.emptyState = document.getElementById("emptyState");
    this.countEl = document.getElementById("resultCount");
    this._bind();
    this._init();
  }

  _bind() {
    document.getElementById("searchBtn").addEventListener("click", () => this._search());
    document.getElementById("reloadFields").addEventListener("click", () => this._loadFields());
    document.getElementById("downloadReport").addEventListener("click", () => this._downloadReport());
    [this.dateFromEl, this.dateToEl].forEach((el) =>
      el.addEventListener("keydown", (e) => { if (e.key === "Enter") this._search(); })
    );
  }

  async _init() {
    try {
      this.config = await SharePointConfigStore.get();
      this.profile = this.config[this.profileKey] || {};
      this.listService = new RelatedListService(this.profile);
      if (!this.listService.configured) {
        this._setStatus(`לא הוגדרה רשימת ${this.noun} — הגדר אותה במסך ההגדרות`, "error");
        return;
      }
      this._renderTabs();
      await this._loadFields();
    } catch (err) {
      this._setStatus(err?.message || "שגיאה בטעינה", "error");
    }
  }

  // Renders a tab per configured type; the active tab drives both filtering and the templates.
  _renderTabs() {
    const types = this.profile.types || [];
    this.typeTabsEl.innerHTML = "";
    this.activeType = types[0] || null;
    if (types.length) {
      const fragment = document.createDocumentFragment();
      types.forEach((type) => {
        const tab = document.createElement("button");
        tab.type = "button";
        tab.className = "tab";
        tab.textContent = type.name;
        tab.classList.toggle("is-active", type === this.activeType);
        tab.addEventListener("click", () => this._selectType(type, tab));
        fragment.appendChild(tab);
      });
      this.typeTabsEl.appendChild(fragment);
    }
    this._renderPrintButtons();
  }

  _selectType(type, tab) {
    this.activeType = type;
    this.typeTabsEl.querySelectorAll(".tab").forEach((t) => t.classList.remove("is-active"));
    tab.classList.add("is-active");
    this._renderPrintButtons();
    if (this.dateFromEl.value || this.dateToEl.value) this._search();
  }

  // One print button per template: the active type's templates, or the profile's flat templates.
  _renderPrintButtons() {
    this.printButtonsEl.innerHTML = "";
    const templates = this.activeType?.templates?.length
      ? this.activeType.templates
      : (this.profile.templates || []);
    if (!templates.length) return;
    const fragment = document.createDocumentFragment();
    templates.forEach((t, i) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "btn btn--accent";
      btn.textContent = t.name || `הדפסת תבנית ${i + 1}`;
      btn.addEventListener("click", () => this._printTemplate(t.path, btn.textContent));
      fragment.appendChild(btn);
    });
    this.printButtonsEl.appendChild(fragment);
  }

  // Builds the location filter checkboxes (all checked) from the location field's options
  // (works for both Lookup fields — by id — and Choice fields — by text).
  async _loadLocationOptions() {
    const field = this.profile.locationField;
    if (!field) { this.locationFilterEl.hidden = true; return; }
    try {
      this.locationOptions = await this.listService.locationOptions();
    } catch {
      this.locationOptions = [];
    }
    this.locationLabels = new Map(this.locationOptions.map((o) => [o.value, o.label]));
    this.locationFilterEl.innerHTML = "";
    if (!this.locationOptions.length) { this.locationFilterEl.hidden = true; return; }
    const label = document.createElement("span");
    label.className = "loc-label";
    label.textContent = "מיקומים:";
    const fragment = document.createDocumentFragment();
    fragment.appendChild(label);
    this.locationOptions.forEach((opt) => {
      const wrap = document.createElement("label");
      wrap.className = "loc-option";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.value = opt.value;
      cb.checked = true;
      wrap.append(cb, document.createTextNode(opt.label));
      fragment.appendChild(wrap);
    });
    this.locationFilterEl.appendChild(fragment);
    this.locationFilterEl.hidden = false;
  }

  _selectedLocations() {
    return [...this.locationFilterEl.querySelectorAll("input:checked")].map((cb) => cb.value);
  }

  // Server-side location filter; null when all (or none) are checked so nothing is narrowed.
  _locationsFilter() {
    if (!this.profile.locationField || !this.locationOptions.length) return null;
    const selected = this._selectedLocations();
    if (!selected.length || selected.length === this.locationOptions.length) return null;
    return { values: selected };
  }

  async _loadFields() {
    if (!this.listService?.configured) return;
    this._setStatus("טוען שדות מהרשימה...", "info");
    try {
      await this.listService.loadFields();
      this.listColumns = this._computeColumns(this.listService.fields, this.profile.displayFields);
      this._render(this.listColumns, []);
      await this._loadLocationOptions();
      this._setStatus(`הרשימה: ${this.profile.listTitle}. בחר טאב וטווח תאריכים והצג.`, "ok");
    } catch (err) {
      this._setStatus(`טעינת שדות נכשלה — ${err?.message || "בדוק הגדרות והרשאות"}`, "error");
    }
  }

  // Columns honor the admin-selected display fields (empty = all list fields, in list order).
  _computeColumns(fields, displayFields) {
    const byInternal = new Map(fields.map((f) => [f.internal, f]));
    const display = (displayFields || []).filter((id) => byInternal.has(id));
    const chosen = display.length ? display : fields.map((f) => f.internal);
    return chosen.map((internal) => ({ tag: byInternal.get(internal)?.title || internal, internal }));
  }

  _peopleMode() { return this.peopleModeEl.checked; }

  _typeFilter() {
    if (!this.activeType || !this.profile.typeField) return null;
    return { field: this.profile.typeField, value: this.activeType.name };
  }

  async _search() {
    if (!this.listService?.configured) {
      return this._setStatus(`לא הוגדרה רשימת ${this.noun} — הגדר אותה במסך ההגדרות`, "error");
    }
    const from = this.dateFromEl.value;
    const to = this.dateToEl.value;
    if (!from && !to) return this._setStatus("בחר לפחות תאריך אחד לסינון", "error");
    this._setStatus(`מחפש ${this.noun}...`, "info");
    try {
      const items = await this.listService.search({
        from, to, type: this._typeFilter(), locations: this._locationsFilter(),
      });
      if (this._peopleMode()) await this._showPeople(items);
      else await this._showItems(items);
    } catch (err) {
      this._setStatus(`חיפוש נכשל — ${err?.message || "בדוק את שדה התאריך וההרשאות"}`, "error");
    }
  }

  // List rows, optionally enriched with base-list columns joined via the Lookup field.
  async _showItems(items) {
    const listRead = (it, internal) => this.listService.valueOf(it, internal);
    const wantsBase = (this.profile.baseDisplayFields || []).length > 0;
    const lookupField = this.profile.lookupField;
    this.rowLocations = this._locationsOf(items);

    if (!wantsBase) {
      this._render(this.listColumns, this._buildRows(items, this.listColumns, listRead));
      return this._setStatus(this._foundMsg(items.length), "ok");
    }
    if (!lookupField) {
      this._render(this.listColumns, this._buildRows(items, this.listColumns, listRead));
      return this._setStatus(`${this._foundMsg(items.length)}. להוספת עמודות בסיס הגדר שדה קישור (Lookup).`, "error");
    }

    await this._ensureBaseFields();
    const baseExtra = this._computeColumns(this.baseFields, this.profile.baseDisplayFields);
    const { merged, base } = this._mergeColumns(this.listColumns, baseExtra, this.profile.basePosition);
    const baseMap = await this._baseMapFor(items, lookupField);
    const baseRead = (it, internal) => this.baseService.readValue(it, internal);

    const rows = items.map((item) => {
      const row = this._rowFrom(this.listColumns, item, listRead);
      const baseItem = baseMap.get(Number(this.listService.lookupId(item, lookupField)));
      base.forEach((c) => { row[c.tag] = baseItem ? SharePointLookupService.formatValue(baseRead(baseItem, c.internal)) : ""; });
      return row;
    });
    this._render(merged, rows);
    this._setStatus(`${this._foundMsg(items.length)} (משולבים עם רשימת הבסיס)`, "ok");
  }

  // Switches the view to the linked base-list files of people who have an item in the range.
  async _showPeople(items) {
    const lookupField = this.profile.lookupField;
    if (!lookupField) throw new Error("לא הוגדר שדה קישור (Lookup) לרשומת התיק בהגדרות");
    this.rowLocations = [];
    await this._ensureBaseFields();
    const baseRead = (it, internal) => this.baseService.readValue(it, internal);
    const baseMap = await this._baseMapFor(items, lookupField);
    const people = [...baseMap.values()];
    this._render(this.baseColumns, this._buildRows(people, this.baseColumns, baseRead));
    if (!people.length) return this._setStatus(`${this._foundMsg(items.length)}, אך ללא קישור לתיקים`, "error");
    this._setStatus(`נמצאו ${people.length} תיקים (מתוך ${items.length} ${this.noun} בטווח)`, "ok");
  }

  _foundMsg(count) {
    const label = this.activeType ? ` מסוג "${this.activeType.name}"` : "";
    return `נמצאו ${count} ${this.noun}${label}`;
  }

  _locationsOf(items) {
    if (!this.profile.locationField) return [];
    return items.map((it) => this.listService.locationKeyOf(it));
  }

  // Splits the current rows into a table per location when results span more than one location.
  _reportGroups() {
    if (this._peopleMode() || !this.profile.locationField) return this.rows;
    if (this.rowLocations.length !== this.rows.length) return this.rows;
    const groups = new Map();
    this.rows.forEach((row, index) => {
      const loc = this.rowLocations[index] || "";
      if (!groups.has(loc)) groups.set(loc, []);
      groups.get(loc).push(row);
    });
    if (groups.size <= 1) return this.rows;
    return [...groups.entries()].map(([key, rows]) => ({
      title: this.locationLabels.get(key) || key || "ללא מיקום",
      rows,
    }));
  }

  // Fetches the linked base-list items for an item set in one batched query, keyed by item id.
  async _baseMapFor(items, lookupField) {
    const ids = items.map((it) => this.listService.lookupId(it, lookupField)).filter((id) => id != null);
    const people = await this.baseService.fetchItemsByIds(ids);
    return new Map(people.map((p) => [Number(p.ID ?? p.Id), p]));
  }

  async _ensureBaseFields() {
    if (this.baseService) return;
    this.baseService = new SharePointLookupService(this.config, this.config.fieldMap || {});
    const raw = await this.baseService.fetchFields();
    this.baseFields = raw
      .filter((f) => !f.Hidden)
      .map((f) => ({ title: f.Title || f.InternalName, internal: f.InternalName || f.StaticName }))
      .filter((f) => f.internal);
    this.baseColumns = this._computeColumns(this.baseFields, this.config.displayFields);
  }

  // Combines list + base columns (base before/after per config), renaming colliding base tags.
  _mergeColumns(listCols, baseCols, position = "before") {
    const used = new Set(listCols.map((c) => c.tag));
    const base = baseCols.map((c) => {
      const tag = used.has(c.tag) ? `${c.tag} (בסיס)` : c.tag;
      used.add(tag);
      return { tag, internal: c.internal };
    });
    const merged = position === "after" ? [...listCols, ...base] : [...base, ...listCols];
    return { merged, base };
  }

  _buildRows(items, columns, read) {
    return items.map((item) => this._rowFrom(columns, item, read));
  }

  _rowFrom(columns, item, read) {
    const row = {};
    columns.forEach((c) => { row[c.tag] = SharePointLookupService.formatValue(read(item, c.internal)); });
    return row;
  }

  _render(columns, rows) {
    this.columns = columns;
    this.rows = rows;
    this._renderHead();
    this._renderRows();
  }

  _renderHead() {
    this.tableHead.innerHTML = "";
    const tr = document.createElement("tr");
    this.columns.forEach((c) => {
      const th = document.createElement("th");
      th.textContent = c.tag;
      tr.appendChild(th);
    });
    const fragment = document.createDocumentFragment();
    fragment.appendChild(tr);
    this.tableHead.appendChild(fragment);
  }

  _renderRows() {
    this.tableBody.innerHTML = "";
    const fragment = document.createDocumentFragment();
    this.rows.forEach((row) => {
      const tr = document.createElement("tr");
      this.columns.forEach((c) => tr.appendChild(this._cell(c, row[c.tag] ?? "")));
      fragment.appendChild(tr);
    });
    this.tableBody.appendChild(fragment);
    this.countEl.textContent = this.rows.length ? `${this.rows.length} שורות` : "";
    this.emptyState.hidden = this.rows.length > 0;
    if (!this.rows.length) this.emptyState.textContent = "לא נמצאו תוצאות בטווח שנבחר.";
  }

  // Long values are clipped with an ellipsis; clicking the cell opens the full text in a modal.
  _cell(column, value) {
    const td = document.createElement("td");
    td.textContent = value;
    if (String(value).length > RelatedListPage.LONG_TEXT_LENGTH) {
      td.classList.add("cell--clip");
      td.title = "לחץ לצפייה בטקסט המלא";
      td.addEventListener("click", () => this._openTextModal(column.tag, value));
    }
    return td;
  }

  _openTextModal(title, text) {
    this._ensureModal();
    this.modalTitleEl.textContent = title;
    this.modalBodyEl.textContent = text;
    this.modalEl.style.display = "flex";
  }

  _closeModal() {
    if (this.modalEl) this.modalEl.style.display = "none";
  }

  _ensureModal() {
    if (this.modalEl) return;
    const overlay = document.createElement("div");
    overlay.className = "modal-overlay";
    overlay.style.display = "none";
    const box = document.createElement("div");
    box.className = "modal";
    const head = document.createElement("div");
    head.className = "modal-header";
    this.modalTitleEl = document.createElement("h3");
    this.modalTitleEl.className = "modal-title";
    const closeBtn = document.createElement("button");
    closeBtn.type = "button";
    closeBtn.className = "modal-close";
    closeBtn.textContent = "✕";
    closeBtn.addEventListener("click", () => this._closeModal());
    head.append(this.modalTitleEl, closeBtn);
    this.modalBodyEl = document.createElement("div");
    this.modalBodyEl.className = "modal-body modal-text";
    box.append(head, this.modalBodyEl);
    overlay.appendChild(box);
    overlay.addEventListener("click", (e) => { if (e.target === overlay) this._closeModal(); });
    document.addEventListener("keydown", (e) => { if (e.key === "Escape") this._closeModal(); });
    document.body.appendChild(overlay);
    this.modalEl = overlay;
  }

  _activeHeaders() {
    return this.columns.map((c) => ({ id: c.tag, label: c.tag }));
  }

  // The chosen date range, formatted for the report title (single date, range, or open-ended).
  _dateLabel() {
    const from = this.dateFromEl.value;
    const to = this.dateToEl.value;
    if (from && to) return from === to ? this._fmtDate(from) : `${this._fmtDate(from)} - ${this._fmtDate(to)}`;
    if (from) return `מ-${this._fmtDate(from)}`;
    if (to) return `עד ${this._fmtDate(to)}`;
    return "";
  }

  _fmtDate(value) {
    const [y, m, d] = value.split("-");
    return `${d}.${m}.${y}`;
  }

  async _downloadReport() {
    if (!this.rows.length) return this._setStatus("אין נתונים לדוח — הצג קודם תוצאות", "error");
    this._setStatus("מכין דוח Word...", "info");
    try {
      const base = this.activeType ? `דוח ${this.activeType.name}` : `דוח ${this.noun}`;
      const dateLabel = this._dateLabel();
      const title = dateLabel ? `${base} — ${dateLabel}` : base;
      await new TableReportService({ title, fileName: `${title}.docx` }).download(this._activeHeaders(), this._reportGroups());
      this._setStatus("הדוח הורד", "ok");
    } catch (err) {
      this._setStatus(`יצירת הדוח נכשלה — ${err?.message || "נסה שוב"}`, "error");
    }
  }

  // Prints every displayed row with the chosen template for the active type.
  async _printTemplate(templatePath, label) {
    if (!this.rows.length) return this._setStatus("אין שורות להדפסה — הצג קודם תוצאות", "error");
    if (!templatePath) return this._setStatus("לא הוגדר נתיב תבנית — הגדר בהגדרות", "error");
    this._setStatus("מכין הדפסה...", "info");
    try {
      const service = new QuickPrintService(
        { templatePath, outputMode: "browserPrint", fields: [] },
        this._activeHeaders()
      );
      await service.printAll(this.rows);
      this._setStatus(`נשלח להדפסה (${label})`, "ok");
    } catch (err) {
      this._setStatus(`הדפסה נכשלה — ${err?.message || "בדוק את התבנית"}`, "error");
    }
  }

  _setStatus(text, type = "") {
    this.statusEl.textContent = text;
    this.statusEl.className = type ? `visible ${type}` : "visible";
  }
}

RelatedListPage.LONG_TEXT_LENGTH = 80;
window.RelatedListPage = RelatedListPage;
