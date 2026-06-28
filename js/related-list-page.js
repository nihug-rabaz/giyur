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
    this.printRows = [];
    this.rowLocations = [];
    this.rowJudges = [];
    this.rowDates = [];
    this.rowSortKeys = [];
    this.selectedIndexes = new Set();
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
    this.selectionBar = new RowSelectionBar({
      onTemplate: (template, event) => this._openTemplateActionModal(
        template.path,
        template.name || "תבנית",
        this._selectedPrintRows(),
        this._selectedCount(),
        event
      ),
      onClear: () => this._clearSelection(),
    });
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
  _activeTemplates() {
    return this.activeType?.templates?.length
      ? this.activeType.templates
      : (this.profile.templates || []);
  }

  _renderPrintButtons() {
    this.printButtonsEl.innerHTML = "";
    const templates = this._activeTemplates();
    if (!templates.length) return;
    const fragment = document.createDocumentFragment();
    templates.forEach((t, i) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "btn btn--accent";
      btn.textContent = t.name || `תבנית ${i + 1}`;
      btn.addEventListener("click", (e) => this._openTemplateActionModal(t.path, btn.textContent, null, null, e));
      fragment.appendChild(btn);
    });
    this.printButtonsEl.appendChild(fragment);
    this._updateSelectionBar();
  }

  // Builds the location filter checkboxes (all unchecked) from the location field's options
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
    fragment.appendChild(this._buildLocationToggle());
    this.locationOptions.forEach((opt) => {
      const wrap = document.createElement("label");
      wrap.className = "loc-option";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.value = opt.value;
      cb.checked = false;
      wrap.append(cb, document.createTextNode(opt.label));
      fragment.appendChild(wrap);
    });
    this.locationFilterEl.appendChild(fragment);
    this.locationFilterEl.hidden = false;
  }

  // Toggle button that checks/unchecks every location checkbox at once.
  _buildLocationToggle() {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "btn btn--sm loc-toggle";
    btn.textContent = "סמן הכל";
    btn.addEventListener("click", () => {
      const boxes = [...this.locationFilterEl.querySelectorAll("input[type=checkbox]")];
      const checkAll = boxes.some((cb) => !cb.checked);
      boxes.forEach((cb) => { cb.checked = checkAll; });
      btn.textContent = checkAll ? "נקה הכל" : "סמן הכל";
    });
    return btn;
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
      this.listColumns = this._listColumnsOnly(this.listService.fields);
      this._render(this._previewColumns(this.listService.fields), []);
      await this._loadLocationOptions();
      this._setStatus(`הרשימה: ${this.profile.listTitle}. בחר טאב וטווח תאריכים והצג.`, "ok");
    } catch (err) {
      this._setStatus(`טעינת שדות נכשלה — ${err?.message || "בדוק הגדרות והרשאות"}`, "error");
    }
  }

  _listColumnsOnly(fields) {
    const layout = this.profile.columnLayout || [];
    const listLayout = layout.filter((col) => col.source === "list");
    return this._computeColumns(fields, listLayout.map((col) => col.internal));
  }

  _previewColumns(listFields, baseFields = []) {
    return this._columnsFromLayout(listFields, baseFields, this.profile.columnLayout || []).columns;
  }

  _columnsFromLayout(listFields, baseFields, layout) {
    if (!layout.length) {
      const columns = this._computeColumns(listFields, []);
      return { columns, listColumns: columns, baseColumns: [] };
    }

    const listByInternal = new Map(listFields.map((field) => [field.internal, field]));
    const baseByInternal = new Map(baseFields.map((field) => [field.internal, field]));
    const usedTags = new Set();
    const columns = [];
    const listColumns = [];
    const baseColumns = [];

    layout.forEach((entry) => {
      const field = entry.source === "base"
        ? baseByInternal.get(entry.internal)
        : listByInternal.get(entry.internal);
      if (!field) return;
      let tag = field.title || entry.internal;
      if (usedTags.has(tag)) tag = entry.source === "base" ? `${tag} (בסיס)` : `${tag} (רשימה)`;
      usedTags.add(tag);
      const column = {
        tag,
        internal: entry.internal,
        source: entry.source,
        fieldType: field.type || "",
        dateOnly: Boolean(field.dateOnly),
      };
      columns.push(column);
      if (entry.source === "base") baseColumns.push(column);
      else listColumns.push(column);
    });

    return { columns, listColumns, baseColumns };
  }

  // Columns honor the admin-selected display fields (empty = all list fields, in list order).
  _computeColumns(fields, displayFields) {
    const byInternal = new Map(fields.map((f) => [f.internal, f]));
    const display = (displayFields || []).filter((id) => byInternal.has(id));
    const chosen = display.length ? display : fields.map((f) => f.internal);
    return chosen.map((internal) => {
      const field = byInternal.get(internal);
      return {
        tag: field?.title || internal,
        internal,
        fieldType: field?.type || "",
        dateOnly: Boolean(field?.dateOnly),
      };
    });
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
    const layout = this.profile.columnLayout || [];
    const wantsBaseInLayout = layout.some((col) => col.source === "base");
    const wantsBaseMapping = this._mappedFields().some((field) => field.source === "base");
    const wantsBase = wantsBaseInLayout || wantsBaseMapping;
    const lookupField = this.profile.lookupField;
    this.rowLocations = this._locationsOf(items);
    this.rowJudges = this._judgesOf(items);
    this.rowDates = this._datesOf(items);
    this.rowSortKeys = items.map((item) => this.listService.itemSortKey(item));

    if (!wantsBase) {
      const { columns } = this._columnsFromLayout(this.listService.fields, [], layout);
      const rows = items.map((item) => {
        const row = this._rowFromPlan(columns, item, { listRead });
        return { row, printRow: this._withTemplateMappings(row, { listItem: item, listRead }) };
      });
      this._render(columns, rows.map((x) => x.row), rows.map((x) => x.printRow));
      return this._setStatus(this._foundMsg(items.length), "ok");
    }
    if (!lookupField) {
      const { columns } = this._columnsFromLayout(this.listService.fields, [], layout);
      const rows = items.map((item) => {
        const row = this._rowFromPlan(columns, item, { listRead });
        return { row, printRow: this._withTemplateMappings(row, { listItem: item, listRead }) };
      });
      this._render(columns, rows.map((x) => x.row), rows.map((x) => x.printRow));
      return this._setStatus(`${this._foundMsg(items.length)}. להוספת עמודות בסיס הגדר שדה קישור (Lookup).`, "error");
    }

    await this._ensureBaseFields();
    const { columns } = this._columnsFromLayout(
      this.listService.fields,
      this.baseFields,
      layout
    );
    const baseMap = await this._baseMapFor(items, lookupField);
    const baseRead = (it, internal) => this.baseService.readValue(it, internal);

    const rows = items.map((item) => {
      const baseItem = baseMap.get(Number(this.listService.lookupId(item, lookupField)));
      const row = this._rowFromPlan(columns, item, { listRead, baseItem, baseRead });
      return { row, printRow: this._withTemplateMappings(row, { listItem: item, baseItem, listRead, baseRead }) };
    });
    this._render(columns, rows.map((x) => x.row), rows.map((x) => x.printRow));
    this._setStatus(`${this._foundMsg(items.length)} (משולבים עם רשימת הבסיס)`, "ok");
  }

  // Switches the view to the linked base-list files of people who have an item in the range.
  async _showPeople(items) {
    const lookupField = this.profile.lookupField;
    if (!lookupField) throw new Error("לא הוגדר שדה קישור (Lookup) לרשומת התיק בהגדרות");
    this.rowLocations = [];
    this.rowJudges = [];
    this.rowDates = [];
    this.rowSortKeys = [];
    await this._ensureBaseFields();
    const baseRead = (it, internal) => this.baseService.readValue(it, internal);
    const baseMap = await this._baseMapFor(items, lookupField);
    const listByBaseId = new Map(items.map((it) => [Number(this.listService.lookupId(it, lookupField)), it]));
    const people = [...baseMap.values()].sort((a, b) => {
      const itemA = listByBaseId.get(Number(a.ID ?? a.Id));
      const itemB = listByBaseId.get(Number(b.ID ?? b.Id));
      return this.listService.itemSortKey(itemA) - this.listService.itemSortKey(itemB);
    });
    const rows = people.map((person) => {
      const row = this._rowFrom(this.baseColumns, person, baseRead);
      const listItem = listByBaseId.get(Number(person.ID ?? person.Id));
      return { row, printRow: this._withTemplateMappings(row, { listItem, baseItem: person, listRead: (it, internal) => this.listService.valueOf(it, internal), baseRead }) };
    });
    this.rowDates = people.map((person) => {
      const listItem = listByBaseId.get(Number(person.ID ?? person.Id));
      return listItem ? this._gregorianDayKey(this.listService.valueOf(listItem, this.profile.dateField)) : "";
    });
    this.rowSortKeys = people.map((person) => {
      const listItem = listByBaseId.get(Number(person.ID ?? person.Id));
      return listItem ? this.listService.itemSortKey(listItem) : 0;
    });
    this._render(this.baseColumns, rows.map((x) => x.row), rows.map((x) => x.printRow));
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

  _judgesOf(items) {
    const field = this.profile.judgeField;
    if (!field) return items.map(() => "");
    return items.map((item) =>
      SharePointLookupService.formatValue(this.listService.valueOf(item, field))
    );
  }

  _datesOf(items) {
    const field = this.profile.dateField;
    if (!field) return items.map(() => "");
    return items.map((item) => this._gregorianDayKey(this.listService.valueOf(item, field)));
  }

  _gregorianDayKey(value) {
    if (!value) return "";
    const d = new Date(value);
    if (Number.isNaN(d.getTime())) {
      const iso = String(value).slice(0, 10);
      return /^\d{4}-\d{2}-\d{2}$/.test(iso) ? iso : "";
    }
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  }

  _rowDateLabel(index) {
    const key = this.rowDates[index] || "";
    return key ? this._fmtDate(key) : "ללא תאריך";
  }

  _rowLocationLabel(index) {
    const key = this.rowLocations[index] || "";
    return this.locationLabels.get(key) || key || "ללא מיקום";
  }

  // Splits rows into separate Word tables by Gregorian day, judge field, and summons location.
  _reportGroups() {
    const hasJudge = !!this.profile.judgeField;
    const hasLocation = !!this.profile.locationField && this.rowLocations.length === this.rows.length;
    const hasDate = !!this.profile.dateField && this.rowDates.length === this.rows.length;
    if (!hasJudge && !hasLocation && !hasDate) {
      return [{ title: null, rows: this.rows, meta: {}, indexes: this.rows.map((_, i) => i) }];
    }

    const buckets = new Map();
    this.rows.forEach((row, index) => {
      const dateKey = hasDate ? (this.rowDates[index] || "ללא-תאריך") : "";
      const dateLabel = hasDate ? this._rowDateLabel(index) : "";
      const judge = hasJudge ? (this.rowJudges[index] || "ללא אב״ד") : "";
      const location = hasLocation ? this._rowLocationLabel(index) : "";
      const key = `${dateKey}\x1f${judge}\x1f${location}`;
      if (!buckets.has(key)) buckets.set(key, { dateKey, dateLabel, judge, location, indexes: [] });
      buckets.get(key).indexes.push(index);
    });

    return [...buckets.values()]
      .sort((a, b) =>
        a.dateKey.localeCompare(b.dateKey)
        || a.location.localeCompare(b.location, "he")
        || a.judge.localeCompare(b.judge, "he"))
      .map((group) => ({
        title: group.location,
        meta: {
          judge: group.judge,
          location: group.location,
          gregorianDate: group.dateLabel,
        },
        indexes: group.indexes,
        rows: this._rowsByIndexes(group.indexes),
      }));
  }

  _rowsByIndexes(indexes) {
    return [...indexes]
      .sort((a, b) => (this.rowSortKeys[a] || 0) - (this.rowSortKeys[b] || 0))
      .map((index) => this.rows[index]);
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
      .map((f) => ({
        title: f.Title || f.InternalName,
        internal: f.InternalName || f.StaticName,
        type: f.TypeAsString || "",
        dateOnly: SharePointLookupService.isDateOnlyField(f),
      }))
      .filter((f) => f.internal);
    this.baseColumns = this._computeColumns(this.baseFields, this.config.displayFields);
  }

  _buildRows(items, columns, read) {
    return items.map((item) => this._rowFrom(columns, item, read));
  }

  _formatCell(value, column) {
    return SharePointLookupService.formatValue(value, {
      fieldType: column?.fieldType,
      dateOnly: column?.dateOnly,
    });
  }

  _rowFromPlan(columns, item, sources) {
    const row = {};
    columns.forEach((column) => {
      if (column.source === "base") {
        const value = sources.baseItem && sources.baseRead
          ? this._formatCell(sources.baseRead(sources.baseItem, column.internal), column)
          : "";
        row[column.tag] = value;
        return;
      }
      row[column.tag] = this._formatCell(sources.listRead(item, column.internal), column);
    });
    return row;
  }

  _rowFrom(columns, item, read) {
    const row = {};
    columns.forEach((c) => { row[c.tag] = this._formatCell(read(item, c.internal), c); });
    return row;
  }

  _render(columns, rows, printRows = rows) {
    this.columns = columns;
    this.rows = rows;
    this.printRows = printRows;
    this._clearSelection(false);
    this._renderHead();
    this._renderRows();
    this._updateSelectionBar();
  }

  _selectedCount() {
    return this.selectedIndexes.size;
  }

  _selectedPrintRows() {
    return [...this.selectedIndexes]
      .sort((a, b) => a - b)
      .map((index) => this.printRows[index] ?? this.rows[index]);
  }

  _clearSelection(updateBar = true) {
    this.selectedIndexes.clear();
    if (this.selectAllEl) this.selectAllEl.checked = false;
    this.tableBody?.querySelectorAll("tr").forEach((tr) => {
      tr.classList.remove("is-selected");
      const cb = tr.querySelector(".row-select");
      if (cb) cb.checked = false;
    });
    if (updateBar) this._updateSelectionBar();
  }

  _toggleRowSelection(index, selected) {
    if (selected) this.selectedIndexes.add(index);
    else this.selectedIndexes.delete(index);
    const tr = this.tableBody?.children[index];
    tr?.classList.toggle("is-selected", selected);
    if (this.selectAllEl) {
      this.selectAllEl.checked = this.rows.length > 0 && this.selectedIndexes.size === this.rows.length;
      this.selectAllEl.indeterminate = this.selectedIndexes.size > 0
        && this.selectedIndexes.size < this.rows.length;
    }
    this._updateSelectionBar();
  }

  _toggleSelectAll(checked) {
    this.selectedIndexes.clear();
    this.tableBody?.querySelectorAll("tr").forEach((tr, index) => {
      const cb = tr.querySelector(".row-select");
      if (cb) cb.checked = checked;
      tr.classList.toggle("is-selected", checked);
      if (checked) this.selectedIndexes.add(index);
    });
    if (this.selectAllEl) this.selectAllEl.indeterminate = false;
    this._updateSelectionBar();
  }

  _updateSelectionBar() {
    this.selectionBar?.update({
      count: this._selectedCount(),
      templates: this._activeTemplates(),
    });
  }

  _renderHead() {
    this.tableHead.innerHTML = "";
    const tr = document.createElement("tr");
    const selectTh = document.createElement("th");
    selectTh.className = "col-select";
    this.selectAllEl = document.createElement("input");
    this.selectAllEl.type = "checkbox";
    this.selectAllEl.className = "row-select-all";
    this.selectAllEl.title = "סמן הכל";
    this.selectAllEl.addEventListener("change", () => this._toggleSelectAll(this.selectAllEl.checked));
    selectTh.appendChild(this.selectAllEl);
    tr.appendChild(selectTh);
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
    this.rows.forEach((row, index) => {
      const tr = document.createElement("tr");
      const selectTd = document.createElement("td");
      selectTd.className = "col-select";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.className = "row-select";
      cb.addEventListener("change", () => this._toggleRowSelection(index, cb.checked));
      selectTd.appendChild(cb);
      tr.appendChild(selectTd);
      this.columns.forEach((c) => tr.appendChild(this._cell(c, row[c.tag] ?? "")));
      fragment.appendChild(tr);
    });
    this.tableBody.appendChild(fragment);
    const selected = this._selectedCount();
    this.countEl.textContent = this.rows.length
      ? (selected ? `${selected} נבחרו מתוך ${this.rows.length} שורות` : `${this.rows.length} שורות`)
      : "";
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
    return this._reportHeaders();
  }

  _reportHeaders() {
    const groupingTags = this._groupingColumnTags();
    const seen = new Set();
    return this.columns
      .filter((column) => !groupingTags.has(column.tag))
      .filter((column) => {
        if (seen.has(column.tag)) return false;
        seen.add(column.tag);
        return true;
      })
      .map((column) => ({ id: column.tag, label: column.tag }));
  }

  _groupingColumnTags() {
    const tags = new Set();
    const internals = new Set([
      this.profile.dateField,
      this.profile.judgeField,
      this.profile.locationField,
    ].filter(Boolean));
    this.columns.forEach((column) => {
      if (internals.has(column.internal)) tags.add(column.tag);
    });
    return tags;
  }

  _mappedFields() {
    return Object.values(this.profile.templateFieldMap || {});
  }

  _templatePrintFields() {
    return Object.entries(this.profile.templateFieldMap || {}).map(([tag, field]) => {
      if (field.source === "system") return { tag, source: "system", internal: field.internal };
      return { tag, column: tag, source: "column" };
    });
  }

  _printHeaders(fields) {
    const headers = this._activeHeaders();
    const used = new Set(headers.map((h) => h.id));
    fields.forEach((field) => {
      if (!used.has(field.tag)) {
        headers.push({ id: field.tag, label: field.tag });
        used.add(field.tag);
      }
    });
    return headers;
  }

  _withTemplateMappings(row, sources) {
    const printRow = { ...row };
    Object.entries(this.profile.templateFieldMap || {}).forEach(([tag, field]) => {
      if (field.source === "system") return;
      const source = field.source === "base" ? "base" : "list";
      const item = source === "base" ? sources.baseItem : sources.listItem;
      const read = source === "base" ? sources.baseRead : sources.listRead;
      printRow[tag] = item && read ? SharePointLookupService.formatValue(read(item, field.internal)) : "";
    });
    return printRow;
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

  _reportMeta(rows, group = {}) {
    return {
      title: "לוח מוזמנים",
      judge: group.meta?.judge || group.judge || "",
      gregorianDate: group.meta?.gregorianDate || this._dateLabel(),
      hebrewDate: this._reportFieldHint(rows, "עברי"),
      location: group.meta?.location || group.title || this._selectedLocationLabel(),
      summonsType: this.activeType?.name || "",
      summonedCount: rows.length,
      printedAt: new Date().toLocaleString("he-IL"),
    };
  }

  _reportFieldHint(rows, keyword) {
    const col = this.columns.find((c) => c.tag.includes(keyword));
    if (!col || !rows.length) return "";
    const val = rows.find((r) => r[col.tag])?.[col.tag];
    return val ? String(val) : "";
  }

  _selectedLocationLabel() {
    const selected = this._selectedLocations();
    if (!selected.length || selected.length === this.locationOptions.length) return "";
    return selected.map((v) => this.locationLabels.get(v) || v).join(", ");
  }

  _normalizeReportGroups() {
    return this._reportGroups().map((group) => ({
      ...group,
      meta: this._reportMeta(group.rows, group),
    }));
  }

  async _downloadReport() {
    if (!this.rows.length) return this._setStatus("אין נתונים לדוח — הצג קודם תוצאות", "error");
    this._setStatus("מכין דוח Word...", "info");
    try {
      const dateLabel = this._dateLabel();
      const fileName = dateLabel ? `לוח-מוזמנים-${dateLabel}.docx` : "לוח-מוזמנים.docx";
      const groups = this._normalizeReportGroups();
      await new ScheduleReportService({ fileName }).download(this._activeHeaders(), groups);
      this._setStatus("הדוח הורד", "ok");
    } catch (err) {
      this._setStatus(`יצירת הדוח נכשלה — ${err?.message || "נסה שוב"}`, "error");
    }
  }

  _openTemplateActionModal(templatePath, label, rowsOverride = null, countOverride = null, anchor = null) {
    const rows = rowsOverride ?? (this.printRows.length ? this.printRows : this.rows);
    if (!rows.length) {
      return this._setStatus(
        rowsOverride ? "לא נבחרו שורות — סמן פריטים בטבלה" : "אין שורות — הצג קודם תוצאות",
        "error"
      );
    }
    if (!templatePath) return this._setStatus("לא הוגדר נתיב תבנית — הגדר בהגדרות", "error");
    if (!this.templateActionModal) this.templateActionModal = new TemplateActionModal();
    const count = countOverride ?? rows.length;
    const hint = rowsOverride
      ? `בחר פעולה עבור ${count} שורות שנבחרו:`
      : `בחר פעולה עבור כל ${count} השורות המוצגות:`;
    this.templateActionModal.open({
      title: label,
      hint,
      anchor,
      onChoose: (mode) => this._runTemplate(templatePath, label, mode, rows),
    });
  }

  async _runTemplate(templatePath, label, outputMode, rowsOverride = null) {
    const isPrint = outputMode === "browserPrint";
    this._setStatus(isPrint ? "מכין הדפסה..." : "מייצא ל-Word...", "info");
    try {
      const fields = this._templatePrintFields();
      const rows = rowsOverride ?? (this.printRows.length ? this.printRows : this.rows);
      const service = new QuickPrintService(
        {
          templatePath,
          outputMode,
          fields,
          templateFieldMap: this.profile.templateFieldMap,
          exportBaseName: label,
          exportZipName: `${FileDownloadHelper.sanitizeFilename(label || "תבנית")}.zip`,
        },
        this._printHeaders(fields)
      );
      await service.printAll(rows);
      this._setStatus(
        isPrint ? `נשלח להדפסה (${label})` : `הורד קובץ Word (${label})`,
        "ok"
      );
    } catch (err) {
      this._setStatus(
        isPrint
          ? `הדפסה נכשלה — ${err?.message || "בדוק את התבנית"}`
          : `ייצוא נכשל — ${err?.message || "בדוק את התבנית"}`,
        "error"
      );
    }
  }

  _setStatus(text, type = "") {
    this.statusEl.textContent = text;
    this.statusEl.className = type ? `visible ${type}` : "visible";
  }
}

RelatedListPage.LONG_TEXT_LENGTH = 80;
window.RelatedListPage = RelatedListPage;
