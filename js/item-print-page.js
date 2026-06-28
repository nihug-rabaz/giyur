// Wizard: search a base-list case, pick it, choose a related list (summons / sessions),
// then see only that case's linked items (filtered by type) and print any template per item.
class ItemPrintPage {
  static PROFILES = [
    { key: "summons", label: "זימונים", noun: "זימונים" },
    { key: "sessions", label: "ישיבות בית דין", noun: "ישיבות" },
  ];

  constructor() {
    this.config = null;
    this.baseFields = [];
    this.itemPrintTemplates = [];
    this._lastMenuAnchor = null;
    this.baseLookup = new SharePointItemLookup();
    this.baseColumns = [];
    this.baseItems = [];
    this.selectedBaseIndex = -1;

    this.models = {};
    this.profileKey = null;
    this.profile = null;
    this.presentTypes = [];
    this.activeType = null;
    this.linkedColumns = [];
    this.linkedRows = [];
    this.linkedPrintRows = [];
    this.linkedTypeValues = [];
    this.selectedLinked = new Set();

    this.menu = new TemplateActionModal();
    this.selectionBar = new RowSelectionBar({
      onTemplate: (template, event) => this._openTemplateMenu(template, event),
      onClear: () => this._clearLinkedSelection(),
    });

    this.searchFieldEl = document.getElementById("searchField");
    this.searchValueEl = document.getElementById("searchValue");
    this.statusEl = document.getElementById("status");
    this.baseCountEl = document.getElementById("baseCount");
    this.baseHead = document.getElementById("baseHead");
    this.baseBody = document.getElementById("baseBody");
    this.baseEmpty = document.getElementById("baseEmpty");

    this.linkedCard = document.getElementById("linkedCard");
    this.linkedTitleEl = document.getElementById("linkedTitle");
    this.linkedSubtitleEl = document.getElementById("linkedSubtitle");
    this.profileTabsEl = document.getElementById("profileTabs");
    this.typeTabsEl = document.getElementById("typeTabs");
    this.linkedStatusEl = document.getElementById("linkedStatus");
    this.linkedWrap = document.getElementById("linkedWrap");
    this.linkedHead = document.getElementById("linkedHead");
    this.linkedBody = document.getElementById("linkedBody");
    this.linkedEmpty = document.getElementById("linkedEmpty");

    this._bind();
    this._init();
  }

  _bind() {
    document.getElementById("searchBtn").addEventListener("click", () => this._searchBase());
    document.getElementById("reloadFields").addEventListener("click", () => this._loadBaseFields());
    this.searchValueEl.addEventListener("keydown", (e) => { if (e.key === "Enter") this._searchBase(); });
  }

  async _init() {
    try {
      this.config = await this.baseLookup.init();
      this.itemPrintTemplates = this.config.itemPrintTemplates || [];
      if (!this.config.listTitle && !this.config.listUrl) {
        return this._setStatus("לא הוגדרה רשימת בסיס — הגדר אותה במסך ההגדרות", "error");
      }
      await this._loadBaseFields();
    } catch (err) {
      this._setStatus(err?.message || "שגיאה בטעינה", "error");
    }
  }

  async _loadBaseFields() {
    this._setStatus("טוען שדות מרשימת הבסיס...", "info");
    try {
      const fields = await this.baseLookup.loadFields();
      this.baseFields = fields.map((f) => ({
        title: f.title,
        internal: f.internal,
        dateOnly: Boolean(f.dateOnly),
      }));
      this._populateSearchField(fields);
      this.baseColumns = this._computeBaseColumns(fields);
      this._renderBaseHead();
      this._setStatus(`בחר שדה חיפוש, הקלד ערך וחפש תיק (${this.config.listTitle || "רשימת בסיס"}).`, "ok");
    } catch (err) {
      this._setStatus(`טעינת שדות נכשלה — ${err?.message || "בדוק הגדרות והרשאות"}`, "error");
    }
  }

  _populateSearchField(fields) {
    this.searchFieldEl.innerHTML = "";
    fields.forEach((f) => this.searchFieldEl.appendChild(new Option(f.title, f.internal)));
    const match = this.config?.join?.listFieldInternal;
    if (match && fields.some((f) => f.internal === match)) this.searchFieldEl.value = match;
  }

  _computeBaseColumns(fields) {
    const byInternal = new Map(fields.map((f) => [f.internal, f]));
    const display = (this.config.displayFields || []).filter((id) => byInternal.has(id));
    const chosen = display.length ? display : fields.slice(0, 6).map((f) => f.internal);
    return chosen.map((internal) => ({
      tag: byInternal.get(internal)?.title || internal,
      internal,
      dateOnly: Boolean(byInternal.get(internal)?.dateOnly),
    }));
  }

  async _searchBase() {
    const value = this.searchValueEl.value.trim();
    if (!value) return this._setStatus("הקלד ערך לחיפוש", "error");
    this._setStatus("מחפש תיק...", "info");
    this._resetLinked();
    try {
      const items = await this.baseLookup.search(this.searchFieldEl.value, value);
      this.baseItems = items || [];
      this.selectedBaseIndex = -1;
      this._renderBaseRows();
      this.baseCountEl.textContent = this.baseItems.length ? `${this.baseItems.length} תיקים נמצאו` : "";
      this._setStatus(
        this.baseItems.length ? "בחר תיק כדי לראות את הפריטים המקושרים" : "לא נמצאו תיקים תואמים",
        this.baseItems.length ? "ok" : "error"
      );
    } catch (err) {
      this._setStatus(`חיפוש נכשל — ${err?.message || "בדוק הרשאות ושדה חיפוש"}`, "error");
    }
  }

  _renderBaseHead() {
    this.baseHead.innerHTML = "";
    const tr = document.createElement("tr");
    const pickTh = document.createElement("th");
    pickTh.className = "col-pick";
    tr.appendChild(pickTh);
    this.baseColumns.forEach((col) => {
      const th = document.createElement("th");
      th.textContent = col.tag;
      tr.appendChild(th);
    });
    this.baseHead.appendChild(tr);
  }

  _renderBaseRows() {
    this.baseBody.innerHTML = "";
    const fragment = document.createDocumentFragment();
    this.baseItems.forEach((item, index) => {
      const tr = document.createElement("tr");
      tr.addEventListener("click", (e) => {
        if (e.target.closest(".col-pick")) return;
        this._selectBase(index, e);
      });
      const pickTd = document.createElement("td");
      pickTd.className = "col-pick";
      const radio = document.createElement("input");
      radio.type = "radio";
      radio.name = "basePick";
      radio.className = "row-pick";
      radio.checked = index === this.selectedBaseIndex;
      radio.addEventListener("click", (e) => { e.stopPropagation(); this._selectBase(index, e); });
      pickTd.appendChild(radio);
      tr.appendChild(pickTd);
      this.baseColumns.forEach((col) => {
        const td = document.createElement("td");
        td.textContent = this._formatBase(item, col);
        tr.appendChild(td);
      });
      tr.classList.toggle("is-selected", index === this.selectedBaseIndex);
      fragment.appendChild(tr);
    });
    this.baseBody.appendChild(fragment);
    this.baseEmpty.hidden = this.baseItems.length > 0;
    if (!this.baseItems.length) this.baseEmpty.textContent = "לא נמצאו תיקים. נסה ערך חיפוש אחר.";
  }

  _formatBase(item, col) {
    return SharePointLookupService.formatValue(this.baseLookup.valueOf(item, col.internal), {
      dateOnly: col.dateOnly,
    });
  }

  // Selecting a base case opens a menu: related lists or direct-print templates from settings.
  _selectBase(index, event) {
    this.selectedBaseIndex = index;
    this._lastMenuAnchor = event;
    this.baseBody.querySelectorAll("tr").forEach((tr, i) => {
      tr.classList.toggle("is-selected", i === index);
      const radio = tr.querySelector(".row-pick");
      if (radio) radio.checked = i === index;
    });
    const items = [
      ...ItemPrintPage.PROFILES.map((p) => ({ label: p.label, value: `profile:${p.key}` })),
      ...this.itemPrintTemplates.map((t) => ({ label: t.name || t.path, value: `template:${t.id}` })),
    ];
    const hasTemplates = this.itemPrintTemplates.length > 0;
    this.menu.open({
      title: this._baseSummary(this.baseItems[index]),
      hint: hasTemplates
        ? "בחר רשימה מקושרת או תבנית להדפסה ישירה מהתיק:"
        : "בחר רשימה להצגת הפריטים המקושרים:",
      anchor: event,
      items,
      onChoose: (value) => this._onBaseMenuChoose(value),
    });
  }

  _onBaseMenuChoose(value) {
    if (value.startsWith("profile:")) {
      this._loadProfile(value.slice(8));
      return;
    }
    if (value.startsWith("template:")) {
      const template = this.itemPrintTemplates.find((t) => t.id === value.slice(9));
      if (template) requestAnimationFrame(() => this._openDirectPrintMenu(template));
    }
  }

  _openDirectPrintMenu(template) {
    const item = this.baseItems[this.selectedBaseIndex];
    if (!item) return;
    if (!template.path) return this._setStatus("לא הוגדר נתיב תבנית", "error");
    this._resetLinked();
    const label = template.name || "תבנית";
    this.menu.open({
      title: label,
      hint: "בחר פעולה:",
      anchor: this._lastMenuAnchor,
      items: TemplateActionModal.PRINT_ITEMS,
      onChoose: (mode) => this._runDirectTemplate(template, mode),
    });
  }

  _buildDirectPrintRow(item, template) {
    const row = {};
    this.baseColumns.forEach((col) => {
      row[col.tag] = this._formatBase(item, col);
    });
    const printRow = { ...row };
    Object.entries(template.templateFieldMap || {}).forEach(([tag, field]) => {
      if (field.source === "system") return;
      if (field.source === "base") {
        printRow[tag] = this._formatBase(item, this._baseFieldColumn(field.internal));
      }
    });
    return printRow;
  }

  _baseFieldColumn(internal) {
    const col = this.baseColumns.find((c) => c.internal === internal);
    if (col) return col;
    const field = this.baseFields.find((f) => f.internal === internal);
    return { internal, tag: field?.title || internal, dateOnly: Boolean(field?.dateOnly) };
  }

  _templatePrintFields(templateFieldMap) {
    return Object.entries(templateFieldMap || {}).map(([tag, field]) => {
      if (field.source === "system") return { tag, source: "system", internal: field.internal };
      return { tag, column: tag, source: "column" };
    });
  }

  _printHeaders(columns, fields) {
    const seen = new Set();
    const headers = columns
      .filter((column) => {
        if (seen.has(column.tag)) return false;
        seen.add(column.tag);
        return true;
      })
      .map((column) => ({ id: column.tag, label: column.tag }));
    const used = new Set(headers.map((h) => h.id));
    fields.forEach((field) => {
      if (!used.has(field.tag)) {
        headers.push({ id: field.tag, label: field.tag });
        used.add(field.tag);
      }
    });
    return headers;
  }

  async _runDirectTemplate(template, outputMode) {
    const item = this.baseItems[this.selectedBaseIndex];
    if (!item) return;
    const row = this._buildDirectPrintRow(item, template);
    const label = template.name || "תבנית";
    const isPrint = outputMode === "browserPrint";
    this._setStatus(isPrint ? "מכין הדפסה..." : "מייצא ל-Word...", "info");
    try {
      const fields = this._templatePrintFields(template.templateFieldMap);
      const service = new QuickPrintService(
        {
          templatePath: template.path,
          outputMode,
          fields,
          templateFieldMap: template.templateFieldMap,
          exportBaseName: label,
          exportZipName: `${FileDownloadHelper.sanitizeFilename(label || "תבנית")}.zip`,
        },
        this._printHeaders(this.baseColumns, fields)
      );
      await service.printAll([row]);
      this._setStatus(isPrint ? `נשלח להדפסה (${label})` : `הורד קובץ Word (${label})`, "ok");
    } catch (err) {
      this._setStatus(
        isPrint
          ? `הדפסה נכשלה — ${err?.message || "בדוק את התבנית"}`
          : `ייצוא נכשל — ${err?.message || "בדוק את התבנית"}`,
        "error"
      );
    }
  }

  _baseSummary(item) {
    const parts = this.baseColumns
      .slice(0, 3)
      .map((col) => this._formatBase(item, col))
      .filter((v) => v);
    return parts.length ? parts.join(" · ") : "תיק נבחר";
  }

  async _loadProfile(key) {
    const item = this.baseItems[this.selectedBaseIndex];
    if (!item) return;
    this.config = await SharePointConfigStore.get();
    this.itemPrintTemplates = this.config.itemPrintTemplates || [];
    this.profileKey = key;
    this.profile = this.config[key] || {};
    const noun = ItemPrintPage.PROFILES.find((p) => p.key === key)?.noun || "פריטים";
    this._renderProfileTabs();
    this.linkedCard.hidden = false;
    this.linkedTitleEl.textContent = `${this._profileLabel(key)} — ${this._baseSummary(item)}`;
    this.linkedSubtitleEl.textContent = "";
    this._setLinkedStatus(`טוען ${noun} מקושרים...`, "info");
    try {
      const model = await this._modelFor(key, true);
      if (!model) return this._setLinkedStatus(`לא הוגדרה רשימת ${noun} — הגדר במסך ההגדרות`, "error");
      const baseId = item.ID ?? item.Id;
      const items = await model.listService.searchByLookup(baseId);
      const built = await model.buildForItems(items, { fixedBaseItem: item });
      this.linkedColumns = built.columns;
      this.linkedRows = built.rows;
      this.linkedPrintRows = built.printRows;
      this.linkedTypeValues = items.map((it) =>
        String(model.listService.valueOf(it, this.profile.typeField) ?? "")
      );
      this._clearLinkedSelection(false);
      this._renderTypeTabs(items.length);
      this._renderLinkedTable();
      this._setLinkedStatus(
        items.length ? `נמצאו ${items.length} ${noun} מקושרים — סמן פריטים והדפס תבנית` : `אין ${noun} מקושרים לתיק זה`,
        items.length ? "ok" : "error"
      );
    } catch (err) {
      this._setLinkedStatus(`טעינה נכשלה — ${err?.message || "בדוק שדה קישור והרשאות"}`, "error");
    }
  }

  async _modelFor(key, refreshProfile = false) {
    const profile = this.config[key] || {};
    if (!this.models[key]) {
      const model = new RelatedListRowModel(profile, this.config);
      if (!model.configured) return null;
      await model.loadFields();
      this.models[key] = model;
      return model;
    }
    if (refreshProfile) {
      this.models[key].profile = profile;
      await this.models[key].loadFields();
    }
    return this.models[key];
  }

  _profileLabel(key) {
    return ItemPrintPage.PROFILES.find((p) => p.key === key)?.label || key;
  }

  _renderProfileTabs() {
    this.profileTabsEl.innerHTML = "";
    const fragment = document.createDocumentFragment();
    ItemPrintPage.PROFILES.forEach((entry) => {
      const tab = document.createElement("button");
      tab.type = "button";
      tab.className = "tab";
      tab.textContent = entry.label;
      tab.classList.toggle("is-active", entry.key === this.profileKey);
      tab.addEventListener("click", () => { if (entry.key !== this.profileKey) this._loadProfile(entry.key); });
      fragment.appendChild(tab);
    });
    this.profileTabsEl.appendChild(fragment);
  }

  // Type tabs show only types that actually have linked items for this case.
  _renderTypeTabs(total) {
    this.typeTabsEl.innerHTML = "";
    this.presentTypes = [];
    this.activeType = null;
    if (!this.profile.typeField || !(this.profile.types || []).length) {
      this.typeTabsEl.hidden = true;
      return;
    }
    const counts = new Map();
    this.linkedTypeValues.forEach((v) => counts.set(v, (counts.get(v) || 0) + 1));
    this.presentTypes = (this.profile.types || []).filter((t) => counts.has(t.name));
    if (!this.presentTypes.length) {
      this.typeTabsEl.hidden = true;
      return;
    }
    this.typeTabsEl.hidden = false;
    this.activeType = this.presentTypes[0];
    const fragment = document.createDocumentFragment();
    this.presentTypes.forEach((type) => {
      const tab = document.createElement("button");
      tab.type = "button";
      tab.className = "tab";
      tab.textContent = `${type.name} (${counts.get(type.name)})`;
      tab.classList.toggle("is-active", type === this.activeType);
      tab.addEventListener("click", () => this._selectType(type, tab));
      fragment.appendChild(tab);
    });
    this.typeTabsEl.appendChild(fragment);
  }

  _selectType(type, tab) {
    this.activeType = type;
    this.typeTabsEl.querySelectorAll(".tab").forEach((t) => t.classList.remove("is-active"));
    tab.classList.add("is-active");
    this._clearLinkedSelection(false);
    this._renderLinkedTable();
  }

  _activeTemplates() {
    if (this.activeType) return this.activeType.templates || [];
    return this.profile.templates || [];
  }

  _visibleIndexes() {
    if (!this.activeType || !this.profile.typeField) {
      return this.linkedRows.map((_, i) => i);
    }
    const indexes = [];
    this.linkedTypeValues.forEach((value, i) => {
      if (value === this.activeType.name) indexes.push(i);
    });
    return indexes;
  }

  _renderLinkedTable() {
    const indexes = this._visibleIndexes();
    this.linkedHead.innerHTML = "";
    const headTr = document.createElement("tr");
    const selectTh = document.createElement("th");
    selectTh.className = "col-select";
    this.selectAllEl = document.createElement("input");
    this.selectAllEl.type = "checkbox";
    this.selectAllEl.title = "סמן הכל";
    this.selectAllEl.addEventListener("change", () => this._toggleSelectAll(this.selectAllEl.checked, indexes));
    selectTh.appendChild(this.selectAllEl);
    headTr.appendChild(selectTh);
    this.linkedColumns.forEach((col) => {
      const th = document.createElement("th");
      th.textContent = col.tag;
      headTr.appendChild(th);
    });
    this.linkedHead.appendChild(headTr);

    this.linkedBody.innerHTML = "";
    const fragment = document.createDocumentFragment();
    indexes.forEach((index) => {
      const row = this.linkedRows[index];
      const tr = document.createElement("tr");
      const selectTd = document.createElement("td");
      selectTd.className = "col-select";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.className = "row-select";
      cb.checked = this.selectedLinked.has(index);
      cb.addEventListener("change", () => this._toggleLinked(index, cb.checked));
      selectTd.appendChild(cb);
      tr.appendChild(selectTd);
      this.linkedColumns.forEach((col) => {
        const td = document.createElement("td");
        td.textContent = row[col.tag] ?? "";
        tr.appendChild(td);
      });
      tr.classList.toggle("is-selected", this.selectedLinked.has(index));
      fragment.appendChild(tr);
    });
    this.linkedBody.appendChild(fragment);

    this.linkedWrap.hidden = false;
    this.linkedEmpty.hidden = indexes.length > 0;
    if (!indexes.length) this.linkedEmpty.textContent = "אין פריטים מסוג זה לתיק.";
    this._updateSelectionBar();
  }

  _toggleLinked(index, selected) {
    if (selected) this.selectedLinked.add(index);
    else this.selectedLinked.delete(index);
    this._syncLinkedRowClasses();
    this._updateSelectionBar();
  }

  _syncLinkedRowClasses() {
    const indexes = this._visibleIndexes();
    [...this.linkedBody.children].forEach((tr, i) => {
      tr.classList.toggle("is-selected", this.selectedLinked.has(indexes[i]));
    });
    if (this.selectAllEl) {
      const visibleSelected = indexes.filter((i) => this.selectedLinked.has(i)).length;
      this.selectAllEl.checked = indexes.length > 0 && visibleSelected === indexes.length;
      this.selectAllEl.indeterminate = visibleSelected > 0 && visibleSelected < indexes.length;
    }
  }

  _toggleSelectAll(checked, indexes) {
    indexes.forEach((i) => { if (checked) this.selectedLinked.add(i); else this.selectedLinked.delete(i); });
    this.linkedBody.querySelectorAll(".row-select").forEach((cb) => { cb.checked = checked; });
    this._syncLinkedRowClasses();
    this._updateSelectionBar();
  }

  _clearLinkedSelection(updateBar = true) {
    this.selectedLinked.clear();
    if (this.selectAllEl) { this.selectAllEl.checked = false; this.selectAllEl.indeterminate = false; }
    this.linkedBody?.querySelectorAll(".row-select").forEach((cb) => { cb.checked = false; });
    this.linkedBody?.querySelectorAll("tr").forEach((tr) => tr.classList.remove("is-selected"));
    if (updateBar) this._updateSelectionBar();
  }

  _updateSelectionBar() {
    this.selectionBar.update({
      count: this.selectedLinked.size,
      templates: this._activeTemplates(),
    });
  }

  _selectedPrintRows() {
    return [...this.selectedLinked]
      .sort((a, b) => a - b)
      .map((index) => this.linkedPrintRows[index] ?? this.linkedRows[index]);
  }

  _openTemplateMenu(template, event) {
    const rows = this._selectedPrintRows();
    if (!rows.length) return this._setLinkedStatus("לא נבחרו פריטים — סמן פריטים בטבלה", "error");
    if (!template.path) return this._setLinkedStatus("לא הוגדר נתיב תבנית", "error");
    const label = template.name || "תבנית";
    this.menu.open({
      title: label,
      hint: `בחר פעולה עבור ${rows.length} פריטים שנבחרו:`,
      anchor: event,
      onChoose: (mode) => this._runTemplate(template.path, label, mode, rows),
    });
  }

  async _runTemplate(templatePath, label, outputMode, rows) {
    const isPrint = outputMode === "browserPrint";
    this._setLinkedStatus(isPrint ? "מכין הדפסה..." : "מייצא ל-Word...", "info");
    try {
      const model = this.models[this.profileKey];
      const fields = model.templatePrintFields();
      const service = new QuickPrintService(
        {
          templatePath,
          outputMode,
          fields,
          templateFieldMap: this.profile.templateFieldMap,
          exportBaseName: label,
          exportZipName: `${FileDownloadHelper.sanitizeFilename(label || "תבנית")}.zip`,
        },
        model.printHeaders(this.linkedColumns, fields)
      );
      await service.printAll(rows);
      this._setLinkedStatus(isPrint ? `נשלח להדפסה (${label})` : `הורד קובץ Word (${label})`, "ok");
    } catch (err) {
      this._setLinkedStatus(
        isPrint
          ? `הדפסה נכשלה — ${err?.message || "בדוק את התבנית"}`
          : `ייצוא נכשל — ${err?.message || "בדוק את התבנית"}`,
        "error"
      );
    }
  }

  _resetLinked() {
    this.profileKey = null;
    this.profile = null;
    this.linkedColumns = [];
    this.linkedRows = [];
    this.linkedPrintRows = [];
    this.linkedTypeValues = [];
    this.selectedLinked.clear();
    this.linkedCard.hidden = true;
    this.linkedWrap.hidden = true;
    this.selectionBar.hide();
  }

  _setStatus(text, type = "") {
    this.statusEl.textContent = text;
    this.statusEl.className = type ? `visible ${type}` : "visible";
  }

  _setLinkedStatus(text, type = "") {
    this.linkedStatusEl.textContent = text;
    this.linkedStatusEl.className = type ? `visible ${type}` : "visible";
  }
}

window.ItemPrintPage = ItemPrintPage;
