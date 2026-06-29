// Reusable "pick fields and order them" widget (used for both the worklist and the summons table).
class OrderedFieldPicker {
  constructor(container) {
    this.container = container;
    this.fields = [];
    this.order = [];
  }

  setFields(fields) { this.fields = fields || []; this.render(); }
  setOrder(order) { this.order = [...(order || [])]; this.render(); }
  getOrder() { return [...this.order]; }

  render() {
    this.container.innerHTML = "";
    this.container.appendChild(this._addRow());

    const list = document.createElement("div");
    list.className = "display-list";
    if (!this.order.length) {
      const hint = document.createElement("p");
      hint.className = "mapping-empty";
      hint.textContent = "לא נבחרו שדות — יוצגו כל השדות. הוסף שדות כדי לקבוע אילו ובאיזה סדר.";
      list.appendChild(hint);
    } else {
      this.order.forEach((internal, i) => list.appendChild(this._row(internal, i)));
    }
    this.container.appendChild(list);
  }

  _addRow() {
    const row = document.createElement("div");
    row.className = "field-row";
    const select = document.createElement("select");
    select.className = "field-input";
    select.appendChild(new Option("— בחר שדה להוספה —", ""));
    this.fields
      .filter((f) => !this.order.includes(f.internal))
      .forEach((f) => select.appendChild(new Option(f.title, f.internal)));

    const add = document.createElement("button");
    add.type = "button";
    add.className = "btn btn--sm";
    add.textContent = "+ הוסף";
    add.addEventListener("click", () => { if (select.value) this._add(select.value); });

    row.append(select, add);
    return row;
  }

  _row(internal, index) {
    const title = this.fields.find((f) => f.internal === internal)?.title || internal;
    const row = document.createElement("div");
    row.className = "field-row";

    const name = document.createElement("span");
    name.className = "display-name";
    name.textContent = title;

    const up = this._iconBtn("↑", () => this._move(index, -1));
    up.disabled = index === 0;
    const down = this._iconBtn("↓", () => this._move(index, 1));
    down.disabled = index === this.order.length - 1;

    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "btn btn--sm btn--danger";
    remove.textContent = "הסר";
    remove.addEventListener("click", () => this._remove(internal));

    row.append(name, up, down, remove);
    return row;
  }

  _iconBtn(text, onClick) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "btn btn--sm";
    btn.textContent = text;
    btn.addEventListener("click", onClick);
    return btn;
  }

  _add(internal) {
    if (!this.order.includes(internal)) this.order.push(internal);
    this.render();
  }

  _remove(internal) {
    this.order = this.order.filter((i) => i !== internal);
    this.render();
  }

  _move(index, dir) {
    const next = index + dir;
    if (next < 0 || next >= this.order.length) return;
    [this.order[index], this.order[next]] = [this.order[next], this.order[index]];
    this.render();
  }
}

class MixedColumnPicker {
  constructor(container) {
    this.container = container;
    this.listFields = [];
    this.baseFields = [];
    this.layout = [];
  }

  setListFields(fields) {
    this.listFields = fields || [];
    this.render();
  }

  setBaseFields(fields) {
    this.baseFields = fields || [];
    this.render();
  }

  setLayout(layout) {
    this.layout = SharePointConfigStore.normalizeColumnLayout(layout);
    this.render();
  }

  getLayout() {
    return this.layout.map((entry) => ({ source: entry.source, internal: entry.internal }));
  }

  render() {
    this.container.innerHTML = "";
    this.container.appendChild(this._addRow());

    const list = document.createElement("div");
    list.className = "display-list";
    if (!this.layout.length) {
      const hint = document.createElement("p");
      hint.className = "mapping-empty";
      hint.textContent = "לא נבחרו עמודות — יוצגו כל שדות הרשימה. הוסף עמודות מרשימת הזימונים/ישיבות ומרשימת הבסיס בכל סדר.";
      list.appendChild(hint);
    } else {
      this.layout.forEach((entry, index) => list.appendChild(this._row(entry, index)));
    }
    this.container.appendChild(list);
  }

  _addRow() {
    const row = document.createElement("div");
    row.className = "field-row";

    const sourceSelect = document.createElement("select");
    sourceSelect.className = "field-input column-source";
    sourceSelect.append(new Option("רשימת זימונים/ישיבות", "list"), new Option("רשימת בסיס", "base"));

    const fieldSelect = document.createElement("select");
    fieldSelect.className = "field-input column-field";

    const syncFields = () => this._populateFieldSelect(fieldSelect, sourceSelect.value);
    sourceSelect.addEventListener("change", syncFields);
    syncFields();

    const add = document.createElement("button");
    add.type = "button";
    add.className = "btn btn--sm";
    add.textContent = "+ הוסף";
    add.addEventListener("click", () => {
      if (!fieldSelect.value) return;
      this._add({ source: sourceSelect.value, internal: fieldSelect.value });
    });

    row.append(sourceSelect, fieldSelect, add);
    return row;
  }

  _populateFieldSelect(select, source) {
    const fields = source === "base" ? this.baseFields : this.listFields;
    select.innerHTML = "";
    select.appendChild(new Option("— בחר שדה —", ""));
    fields
      .filter((field) => !this._isUsed(source, field.internal))
      .forEach((field) => select.appendChild(new Option(field.title, field.internal)));
  }

  _row(entry, index) {
    const row = document.createElement("div");
    row.className = "field-row";

    const badge = document.createElement("span");
    badge.className = "display-name";
    badge.style.minWidth = "88px";
    badge.textContent = entry.source === "base" ? "בסיס" : "רשימה";

    const name = document.createElement("span");
    name.className = "display-name";
    name.textContent = this._fieldTitle(entry);

    const up = this._iconBtn("↑", () => this._move(index, -1));
    up.disabled = index === 0;
    const down = this._iconBtn("↓", () => this._move(index, 1));
    down.disabled = index === this.layout.length - 1;

    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "btn btn--sm btn--danger";
    remove.textContent = "הסר";
    remove.addEventListener("click", () => this._remove(index));

    row.append(badge, name, up, down, remove);
    return row;
  }

  _iconBtn(text, onClick) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "btn btn--sm";
    btn.textContent = text;
    btn.addEventListener("click", onClick);
    return btn;
  }

  _fieldTitle(entry) {
    const pool = entry.source === "base" ? this.baseFields : this.listFields;
    return pool.find((field) => field.internal === entry.internal)?.title || entry.internal;
  }

  _isUsed(source, internal) {
    return this.layout.some((entry) => entry.source === source && entry.internal === internal);
  }

  _add(entry) {
    if (this._isUsed(entry.source, entry.internal)) return;
    this.layout.push({ source: entry.source === "base" ? "base" : "list", internal: entry.internal });
    this.render();
  }

  _remove(index) {
    this.layout.splice(index, 1);
    this.render();
  }

  _move(index, dir) {
    const next = index + dir;
    if (next < 0 || next >= this.layout.length) return;
    [this.layout[index], this.layout[next]] = [this.layout[next], this.layout[index]];
    this.render();
  }
}

// Self-contained editor for one secondary related list (summons / sessions).
// Builds its own card UI and reads/writes the prefixed config keys for that profile.
class RelatedListEditor {
  constructor(container, prefix, labels, host) {
    this.container = container;
    this.prefix = prefix;
    this.labels = labels;
    this.host = host;
    this.useTypeTabs = labels.useTypeTabs !== false;
    this.fields = [];
    this._build();
  }

  _key(suffix) { return `${this.prefix}${suffix}`; }

  _build() {
    this.container.innerHTML = "";
    this.container.append(this._heading(this.labels.title), this._para(this.labels.intro));

    this.listTitleEl = this._input("text-input", this.labels.listPlaceholder, "");
    this.container.append(this._label(this.labels.listLabel), this.listTitleEl);

    this.fieldsInfoEl = this._infoSpan();
    const loadBtn = this._btn(this.labels.loadFieldsText);
    loadBtn.addEventListener("click", () => this.loadFields());
    const loadRow = document.createElement("div");
    loadRow.style.marginTop = "12px";
    loadRow.append(loadBtn, this.fieldsInfoEl);
    this.container.append(loadRow);

    this.dateFieldEl = this._select();
    this.container.append(this._label("שדה תאריך לסינון"), this.dateFieldEl);

    this.lookupFieldEl = this._select();
    this.container.append(this._label("שדה קישור לרשומת התיק (Lookup)"), this.lookupFieldEl);

    this.container.append(this._heading("עמודות הטבלה — סדר ומיקום", true));
    this.container.append(this._para('הוסף עמודות מרשימת הזימונים/ישיבות ומרשימת הבסיס בכל סדר. שדות הבסיס מוצגים לפי שדה הקישור (Lookup). טען שדות בסיס בכרטיס "שיוך מ-SharePoint".'));
    const layoutHost = document.createElement("div");
    this.container.append(layoutHost);
    this.columnLayoutPicker = new MixedColumnPicker(layoutHost);

    this._buildTemplateFieldMapSection();

    if (this.useTypeTabs) this._buildTypeSection();
    else this._buildTemplatesSection();

    this.locationFieldEl = this._select();
    this.container.append(this._heading("עמודת מיקום (אפשרות/בדיקת מידע, לא חובה)", true));
    this.container.append(this._para('תומך בעמודת "אפשרות" (Choice) או "בדיקת מידע" (Lookup). אם תוגדר, יופיע מסנן מיקומים, ובדוח Word תיווצר טבלה נפרדת לכל שילוב של יום (תאריך לועזי), אב״ד ומיקום.'));
    this.container.append(this.locationFieldEl);

    this.judgeFieldEl = this._select();
    this.container.append(this._heading("שדה אב״ד לדוח Word (לא חובה)", true));
    this.container.append(this._para('בחר את השדה שמכיל את שם אב״ד. בדוח Word תיווצר טבלה נפרדת לכל שילוב של יום, אב״ד ומיקום הזימון. אם הטבלה ארוכה — Word ימשיך אותה בעמוד הבא.'));
    this.container.append(this.judgeFieldEl);

    if (this.labels.showSummonedField) {
      this.summonedFieldEl = this._select();
      this.container.append(this._heading('שדה "האם זומן" (לא חובה)', true));
      this.container.append(this._para('בחר עמודת כן/לא. פריטים עם ערך "לא" לא ייכללו בדוח Word ובהדפסה, אלא אם המשתמש יבחר לכלול אותם במסך הזימונים.'));
      this.container.append(this.summonedFieldEl);
    }
  }

  _buildTemplateFieldMapSection() {
    this.container.append(this._heading("מיפוי תגי תבנית Word", true));
    this.container.append(this._para("כאן מקשרים תג שמופיע בתבנית בתוך [ ] לשדה מהרשימה, מרשימת הבסיס, או לשדות מערכת. לתאריכים: קשר 3 תגים נפרדים — יום, חודש, שנה (לועזי / עברי / הדפסה). אם לא מוגדר מיפוי, ההדפסה תמשיך לעבוד לפי שמות כותרות הטבלה."));
    const head = document.createElement("div");
    head.className = "grid-head";
    head.append(this._span("שם התג בוורד"), this._span("מקור השדה"), this._span("שדה למילוי"), this._span(""));
    this.templateFieldMapEl = document.createElement("div");
    const addBtn = this._btn("+ הוסף מיפוי לתבנית");
    addBtn.addEventListener("click", () => this._addTemplateFieldMapRow({}));
    this.container.append(head, this.templateFieldMapEl, addBtn);
  }

  // Tab-driven types: a Choice column whose values become tabs, each with its own templates.
  _buildTypeSection() {
    this.container.append(this._heading(this.labels.typeTitle, true), this._para(this.labels.typeIntro));
    this.typeFieldEl = this._select();
    this.container.append(this._label(this.labels.typeFieldLabel), this.typeFieldEl);

    this.typeInfoEl = this._infoSpan();
    const loadTypesBtn = this._btn("טען אפשרויות מהשדה");
    loadTypesBtn.addEventListener("click", () => this.loadTypeChoices());
    const typesRow = document.createElement("div");
    typesRow.style.marginTop = "12px";
    typesRow.append(loadTypesBtn, this.typeInfoEl);
    this.container.append(typesRow);

    this.typesEl = document.createElement("div");
    this.typesEl.style.marginTop = "10px";
    const addTypeBtn = this._btn("+ הוסף סוג");
    addTypeBtn.addEventListener("click", () => this._addTypeCard({}));
    this.container.append(this.typesEl, addTypeBtn);
  }

  // Flat templates: each is just a print button (no tabs, no Choice column).
  _buildTemplatesSection() {
    this.container.append(this._heading(this.labels.templatesTitle || "תבניות הדפסה (כפתורים)", true));
    this.container.append(this._para(this.labels.templatesIntro || "כל תבנית תופיע ככפתור. הגדר קובץ נפרד להדפסה ולייצוא Word — אפשר להשאיר רק אחד מהם."));
    this.flatTemplatesEl = document.createElement("div");
    this.flatTemplatesEl.className = "type-templates";
    const addFlatBtn = this._btn("+ הוסף תבנית");
    addFlatBtn.addEventListener("click", () => this.flatTemplatesEl.appendChild(this._templateRow({})));
    this.container.append(this.flatTemplatesEl, addFlatBtn);
  }

  // Populates the editor from the raw saved config (prefixed keys).
  load(raw) {
    this.listTitleEl.value = raw[this._key("ListTitle")] || "";
    this.host._populateSelect(this.dateFieldEl, raw[this._key("DateFieldInternal")], this.fields);
    this.host._populateSelect(this.lookupFieldEl, raw[this._key("LookupFieldInternal")], this.fields);
    this.host._populateSelect(this.locationFieldEl, raw[this._key("LocationFieldInternal")], this.fields);
    this.host._populateSelect(this.judgeFieldEl, raw[this._key("JudgeFieldInternal")], this.fields);
    if (this.summonedFieldEl) {
      this.host._populateSelect(this.summonedFieldEl, raw[this._key("SummonedFieldInternal")], this.fields);
    }
    this.columnLayoutPicker.setListFields(this.fields);
    this.columnLayoutPicker.setBaseFields(this.host.getBaseFields());
    this.columnLayoutPicker.setLayout(SharePointConfigStore.resolveColumnLayout(raw, this.prefix));
    this.templateFieldMapEl.innerHTML = "";
    const fieldMap = SharePointConfigStore.normalizeTemplateFieldMap(raw[this._key("TemplateFieldMap")]);
    Object.entries(fieldMap).forEach(([tag, field]) => this._addTemplateFieldMapRow({ tag, ...field }));
    if (!Object.keys(fieldMap).length) this._addTemplateFieldMapRow({});

    if (this.useTypeTabs) {
      this.host._populateSelect(this.typeFieldEl, raw[this._key("TypeFieldInternal")], this.fields);
      this.typesEl.innerHTML = "";
      const types = SharePointConfigStore.normalizeTypes(raw[this._key("Types")]);
      types.forEach((t) => this._addTypeCard(t));
      if (!types.length) this._addTypeCard({});
    } else {
      this.flatTemplatesEl.innerHTML = "";
      const templates = SharePointConfigStore.normalizeTemplates(raw[this._key("Templates")]);
      templates.forEach((t) => this.flatTemplatesEl.appendChild(this._templateRow(t)));
      if (!templates.length) this.flatTemplatesEl.appendChild(this._templateRow({}));
    }
  }

  // Returns this profile's settings as prefixed config keys.
  collect() {
    const columnLayout = this.columnLayoutPicker.getLayout();
    return {
      [this._key("ListTitle")]: this.listTitleEl.value.trim(),
      [this._key("DateFieldInternal")]: this.dateFieldEl.value,
      [this._key("LookupFieldInternal")]: this.lookupFieldEl.value,
      [this._key("ColumnLayout")]: columnLayout,
      [this._key("DisplayFields")]: columnLayout.filter((col) => col.source === "list").map((col) => col.internal),
      [this._key("BaseDisplayFields")]: columnLayout.filter((col) => col.source === "base").map((col) => col.internal),
      [this._key("BasePosition")]: "before",
      [this._key("TypeFieldInternal")]: this.useTypeTabs ? this.typeFieldEl.value : "",
      [this._key("Types")]: this.useTypeTabs ? this._collectTypes() : [],
      [this._key("Templates")]: this.useTypeTabs ? [] : this._collectTemplates(),
      [this._key("LocationFieldInternal")]: this.locationFieldEl.value,
      [this._key("JudgeFieldInternal")]: this.judgeFieldEl.value,
      ...(this.summonedFieldEl ? { [this._key("SummonedFieldInternal")]: this.summonedFieldEl.value } : {}),
      [this._key("TemplateFieldMap")]: this._collectTemplateFieldMap(),
    };
  }

  setBaseFields(fields) {
    this.columnLayoutPicker.setBaseFields(fields);
    this._refreshTemplateFieldRows();
  }

  async loadFields() {
    const siteUrl = this.host.getSiteUrl();
    const listTitle = this.listTitleEl.value.trim();
    if (!siteUrl || !listTitle) return this._setInfo(this.fieldsInfoEl, "יש למלא כתובת אתר וכותרת רשימה");
    this._setInfo(this.fieldsInfoEl, "טוען שדות...");
    try {
      this.fields = await this.host._loadFieldList(siteUrl, listTitle);
      const selects = [this.dateFieldEl, this.lookupFieldEl, this.typeFieldEl, this.locationFieldEl, this.judgeFieldEl];
      if (this.summonedFieldEl) selects.push(this.summonedFieldEl);
      selects.filter(Boolean).forEach((sel) => this.host._populateSelect(sel, sel.value, this.fields));
      this.columnLayoutPicker.setListFields(this.fields);
      this._refreshTemplateFieldRows();
      this._setInfo(this.fieldsInfoEl, `נטענו ${this.fields.length} שדות`);
    } catch (err) {
      console.error(err);
      this._setInfo(this.fieldsInfoEl, `טעינת השדות נכשלה — ${err?.message || "בדוק כתובת והרשאות"}`);
    }
  }

  async loadTypeChoices() {
    const siteUrl = this.host.getSiteUrl();
    const listTitle = this.listTitleEl.value.trim();
    const fieldInternal = this.typeFieldEl.value;
    if (!siteUrl || !listTitle || !fieldInternal) {
      return this._setInfo(this.typeInfoEl, "בחר עמודת סוג וודא שכותרת הרשימה מלאה");
    }
    this._setInfo(this.typeInfoEl, "טוען אפשרויות...");
    try {
      const choices = await this.host._fetchChoices(siteUrl, listTitle, fieldInternal);
      const byName = new Map(this._collectTypes().map((t) => [t.name, t.templates]));
      this.typesEl.innerHTML = "";
      choices.forEach((name) => this._addTypeCard({ name, templates: byName.get(name) }));
      if (!choices.length) this._addTypeCard({});
      this._setInfo(this.typeInfoEl, `נטענו ${choices.length} אפשרויות`);
    } catch (err) {
      console.error(err);
      this._setInfo(this.typeInfoEl, `טעינת האפשרויות נכשלה — ${err?.message || "בדוק כתובת והרשאות"}`);
    }
  }

  // A type card: a Choice value plus one or more Word templates used when printing it.
  _addTypeCard(type = {}) {
    const templates = type.templates?.length
      ? type.templates
      : [{ name: "", printPath: type.templatePath || "", exportPath: type.templatePath || "" }];
    const card = document.createElement("div");
    card.className = "type-card";

    const head = document.createElement("div");
    head.className = "field-row";
    const nameInput = this._input("field-input type-name", "שם הסוג (כפי שמופיע באפשרויות)", type.name || "");
    const removeType = document.createElement("button");
    removeType.type = "button";
    removeType.className = "btn btn--sm btn--danger";
    removeType.textContent = "הסר סוג";
    removeType.addEventListener("click", () => card.remove());
    head.append(nameInput, removeType);

    const templatesEl = document.createElement("div");
    templatesEl.className = "type-templates";
    templates.forEach((t) => templatesEl.appendChild(this._templateRow(t)));

    const addTmpl = this._btn("+ הוסף תבנית");
    addTmpl.addEventListener("click", () => templatesEl.appendChild(this._templateRow({})));

    card.append(head, templatesEl, addTmpl);
    this.typesEl.appendChild(card);
  }

  _templateRow(template = {}) {
    const row = document.createElement("div");
    row.className = "field-row tmpl-row";
    const nameInput = this._input("field-input tmpl-name", "שם התבנית (לכפתור)", template.name || "");
    const printPathInput = this._input("field-input tmpl-print-path", "תבנית להדפסה (.docx)", template.printPath || template.path || "");
    const exportPathInput = this._input("field-input tmpl-export-path", "תבנית לייצוא Word (.docx)", template.exportPath || template.path || "");
    row.append(nameInput, printPathInput, exportPathInput, this.host._removeBtn(row));
    return row;
  }

  _addTemplateFieldMapRow(field = {}) {
    const row = document.createElement("div");
    row.className = "field-row template-map-row";
    const tagInput = this._input("field-input template-tag", "שם התג בוורד", field.tag || "");
    const sourceSelect = this._select();
    sourceSelect.classList.add("template-source");
    sourceSelect.append(new Option("רשימת הזימונים/ישיבות", "list"), new Option("רשימת הבסיס", "base"), new Option("שדות מערכת (אוטומטי)", "system"));
    sourceSelect.value = field.source === "base" ? "base" : field.source === "system" ? "system" : "list";
    const fieldSelect = this._select();
    fieldSelect.classList.add("template-field");
    row.append(tagInput, sourceSelect, fieldSelect, this.host._removeBtn(row));
    sourceSelect.addEventListener("change", () => this._populateTemplateFieldSelect(fieldSelect, sourceSelect.value, fieldSelect.value));
    this.templateFieldMapEl.appendChild(row);
    this._populateTemplateFieldSelect(fieldSelect, sourceSelect.value, field.internal || "");
  }

  _populateTemplateFieldSelect(select, source, selected) {
    if (source === "system") {
      select.innerHTML = "";
      SystemDateService.FIELDS.forEach((field) => select.append(new Option(field.title, field.internal)));
      select.value = selected || SystemDateService.FIELDS[0].internal;
      return;
    }
    const fields = source === "base" ? this.host.getBaseFields() : this.fields;
    this.host._populateSelect(select, selected, fields);
  }

  _refreshTemplateFieldRows() {
    if (!this.templateFieldMapEl) return;
    this.templateFieldMapEl.querySelectorAll(".template-map-row").forEach((row) => {
      const source = row.querySelector(".template-source").value;
      const select = row.querySelector(".template-field");
      this._populateTemplateFieldSelect(select, source, select.value);
    });
  }

  _collectTypes() {
    const types = [];
    this.typesEl.querySelectorAll(".type-card").forEach((card) => {
      const name = card.querySelector(".type-name").value.trim();
      if (!name) return;
      const templates = [];
      card.querySelectorAll(".tmpl-row").forEach((row) => {
        const tName = row.querySelector(".tmpl-name").value.trim();
        const printPath = row.querySelector(".tmpl-print-path").value.trim();
        const exportPath = row.querySelector(".tmpl-export-path").value.trim();
        if (printPath || exportPath) templates.push({ name: tName, printPath, exportPath });
      });
      types.push({ name, templates });
    });
    return types;
  }

  _collectTemplates() {
    const templates = [];
    this.flatTemplatesEl.querySelectorAll(".tmpl-row").forEach((row) => {
      const name = row.querySelector(".tmpl-name").value.trim();
      const printPath = row.querySelector(".tmpl-print-path").value.trim();
      const exportPath = row.querySelector(".tmpl-export-path").value.trim();
      if (printPath || exportPath) templates.push({ name, printPath, exportPath });
    });
    return templates;
  }

  _collectTemplateFieldMap() {
    const map = {};
    this.templateFieldMapEl.querySelectorAll(".template-map-row").forEach((row) => {
      const tag = row.querySelector(".template-tag").value.trim();
      const source = row.querySelector(".template-source").value;
      const internal = row.querySelector(".template-field").value;
      if (!tag || !internal) return;
      map[tag] = source === "system" ? { source: "system", internal } : { source, internal };
    });
    return map;
  }

  // ---- Small DOM builders ----

  _input(className, placeholder, value) { return this.host._input(className, placeholder, value); }

  _select() {
    const select = document.createElement("select");
    select.className = "field-input";
    select.style.width = "100%";
    return select;
  }

  _heading(text, small = false) {
    const h = document.createElement("h2");
    if (small) { h.style.marginTop = "16px"; h.style.fontSize = "13px"; }
    h.textContent = text;
    return h;
  }

  _para(text) {
    const p = document.createElement("p");
    p.className = "mapping-empty";
    p.style.marginBottom = "10px";
    p.textContent = text;
    return p;
  }

  _span(text) {
    const span = document.createElement("span");
    span.textContent = text;
    return span;
  }

  _label(text) {
    const l = document.createElement("label");
    l.className = "field-label";
    l.style.marginTop = "12px";
    l.textContent = text;
    return l;
  }

  _btn(text) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "btn btn--sm";
    btn.style.flex = "0 0 auto";
    btn.textContent = text;
    return btn;
  }

  _infoSpan() {
    const span = document.createElement("span");
    span.style.fontSize = "12px";
    span.style.color = "var(--text-muted)";
    span.style.marginRight = "8px";
    return span;
  }

  _setInfo(el, text) { el.textContent = text; }
}

class ItemPrintTemplatesEditor {
  constructor(container, host) {
    this.container = container;
    this.host = host;
    this.baseFields = [];
    this._build();
  }

  _build() {
    this.container.innerHTML = "";
    const title = document.createElement("h2");
    title.textContent = "תבניות הדפסה לפריט בודד (מסך הדפסה לפריט)";
    const intro = document.createElement("p");
    intro.className = "mapping-empty";
    intro.style.marginBottom = "10px";
    intro.textContent = "הוסף תבניות שיופיעו בתפריט בעת בחירת תיק. לכל תבנית הגדר קובץ Word נפרד להדפסה ולייצוא. קשר תגי Word לשדות מרשימת הבסיס או לשדות מערכת.";
    this.templatesEl = document.createElement("div");
    this.templatesEl.style.marginTop = "10px";
    const addBtn = document.createElement("button");
    addBtn.type = "button";
    addBtn.className = "btn btn--sm";
    addBtn.textContent = "+ הוסף תבנית";
    addBtn.addEventListener("click", () => this._addCard({}));
    this.container.append(title, intro, this.templatesEl, addBtn);
  }

  load(raw) {
    this.templatesEl.innerHTML = "";
    const templates = SharePointConfigStore.normalizeItemPrintTemplates(raw.itemPrintTemplates);
    templates.forEach((template) => this._addCard(template));
    if (!templates.length) this._addCard({});
  }

  collect() {
    return { itemPrintTemplates: this._collectTemplates() };
  }

  setBaseFields(fields) {
    this.baseFields = fields || [];
    this._refreshFieldRows();
  }

  _addCard(template = {}) {
    const card = document.createElement("div");
    card.className = "type-card";
    card.dataset.id = template.id || `tpl_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;

    const nameInput = this.host._input("field-input tpl-menu-name", "שם בתפריט", template.name || "");
    const printPathInput = this.host._input("field-input tpl-print-path", "תבנית להדפסה (.docx)", template.printPath || template.path || "");
    const exportPathInput = this.host._input("field-input tpl-export-path", "תבנית לייצוא Word (.docx)", template.exportPath || template.path || "");
    const topRow = document.createElement("div");
    topRow.className = "field-row tmpl-row";
    topRow.append(nameInput, printPathInput, exportPathInput);

    const mapTitle = document.createElement("p");
    mapTitle.className = "mapping-empty";
    mapTitle.style.margin = "10px 0 6px";
    mapTitle.textContent = "מיפוי תגי Word לשדות מרשימת הבסיס:";

    const mapHead = document.createElement("div");
    mapHead.className = "grid-head";
    ["שם התג בוורד", "מקור השדה", "שדה למילוי", ""].forEach((text) => {
      const span = document.createElement("span");
      span.textContent = text;
      mapHead.appendChild(span);
    });

    const mapEl = document.createElement("div");
    mapEl.className = "template-field-map";
    const fieldMap = template.templateFieldMap || {};
    Object.entries(fieldMap).forEach(([tag, field]) => this._addFieldMapRow(mapEl, { tag, ...field }));
    if (!Object.keys(fieldMap).length) this._addFieldMapRow(mapEl, {});

    const addMapBtn = document.createElement("button");
    addMapBtn.type = "button";
    addMapBtn.className = "btn btn--sm";
    addMapBtn.textContent = "+ הוסף מיפוי";
    addMapBtn.addEventListener("click", () => this._addFieldMapRow(mapEl, {}));

    const removeCard = this.host._removeBtn(card);
    removeCard.textContent = "הסר תבנית";
    card.append(topRow, mapTitle, mapHead, mapEl, addMapBtn, removeCard);
    this.templatesEl.appendChild(card);
  }

  _addFieldMapRow(container, field = {}) {
    const row = document.createElement("div");
    row.className = "field-row template-map-row";
    const tagInput = this.host._input("field-input template-tag", "שם התג בוורד", field.tag || "");
    const sourceSelect = document.createElement("select");
    sourceSelect.className = "field-input template-source";
    sourceSelect.append(new Option("רשימת הבסיס", "base"), new Option("שדות מערכת (אוטומטי)", "system"));
    sourceSelect.value = field.source === "system" ? "system" : "base";
    const fieldSelect = document.createElement("select");
    fieldSelect.className = "field-input template-field";
    sourceSelect.addEventListener("change", () => this._populateFieldSelect(fieldSelect, sourceSelect.value, fieldSelect.value));
    row.append(tagInput, sourceSelect, fieldSelect, this.host._removeBtn(row));
    container.appendChild(row);
    this._populateFieldSelect(fieldSelect, sourceSelect.value, field.internal || "");
  }

  _populateFieldSelect(select, source, selected) {
    if (source === "system") {
      select.innerHTML = "";
      SystemDateService.FIELDS.forEach((field) => select.append(new Option(field.title, field.internal)));
      select.value = selected || SystemDateService.FIELDS[0].internal;
      return;
    }
    this.host._populateSelect(select, selected, this.baseFields);
  }

  _refreshFieldRows() {
    this.templatesEl.querySelectorAll(".template-map-row").forEach((row) => {
      const source = row.querySelector(".template-source").value;
      const select = row.querySelector(".template-field");
      this._populateFieldSelect(select, source, select.value);
    });
  }

  _collectTemplates() {
    const templates = [];
    this.templatesEl.querySelectorAll(".type-card").forEach((card) => {
      const printPath = card.querySelector(".tpl-print-path")?.value.trim() || "";
      const exportPath = card.querySelector(".tpl-export-path")?.value.trim() || "";
      if (!printPath && !exportPath) return;
      const name = card.querySelector(".tpl-menu-name").value.trim();
      const id = card.dataset.id || `tpl_${templates.length}`;
      const templateFieldMap = {};
      card.querySelectorAll(".template-map-row").forEach((row) => {
        const tag = row.querySelector(".template-tag").value.trim();
        const source = row.querySelector(".template-source").value;
        const internal = row.querySelector(".template-field").value;
        if (!tag || !internal) return;
        templateFieldMap[tag] = source === "system" ? { source: "system", internal } : { source: "base", internal };
      });
      templates.push({ id, name, printPath, exportPath, templateFieldMap });
    });
    return templates;
  }
}

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
    this.displayPicker = new OrderedFieldPicker(document.getElementById("spDisplayFields"));
    this.itemPrintEditor = new ItemPrintTemplatesEditor(document.getElementById("itemPrintEditor"), this);

    this.editors = [
      new RelatedListEditor(document.getElementById("summonsEditor"), "summons", {
        title: "רשימת זימונים (רשימה משנית)",
        intro: 'הגדר את רשימת הזימונים, שדה התאריך הלועזי לסינון, ושדה הקישור (Lookup) לרשומת התיק. לאחר מכן ניתן לפתוח את מסך "זימונים לישיבות" ולסנן לפי טווח תאריכים.',
        listLabel: "כותרת רשימת הזימונים",
        listPlaceholder: "זימונים",
        loadFieldsText: "טען שדות מרשימת הזימונים",
        displayTitle: "שדות שיוצגו בטבלת הזימונים",
        baseTitle: "עמודות מרשימת הבסיס שיוצגו ליד הזימון (לפי ה-Lookup)",
        typeTitle: "סוגי זימון ותבניות (טאבים במסך הזימונים)",
        typeIntro: 'בחר את עמודת "סוג זימון" (מסוג אפשרות), טען את האפשרויות, וקשר לכל סוג תבנית אחת או יותר. כל סוג יופיע כטאב, ולכל תבנית יופיע כפתור הדפסה.',
        typeFieldLabel: "עמודת סוג הזימון (Choice)",
        showSummonedField: true,
      }, this),
      new RelatedListEditor(document.getElementById("sessionsEditor"), "sessions", {
        title: "רשימת ישיבות בית דין (רשימה משנית)",
        intro: 'הגדר את רשימת ישיבות בית הדין, שדה התאריך לסינון, ושדה הקישור (Lookup) לרשומת התיק. לאחר מכן ניתן לפתוח את מסך "ישיבות בית דין" ולסנן לפי טווח תאריכים.',
        listLabel: "כותרת רשימת הישיבות",
        listPlaceholder: "ישיבות בת דין",
        loadFieldsText: "טען שדות מרשימת הישיבות",
        displayTitle: "שדות שיוצגו בטבלת הישיבות",
        baseTitle: "עמודות מרשימת הבסיס שיוצגו ליד הישיבה (לפי ה-Lookup)",
        useTypeTabs: false,
        templatesTitle: "סוגי ממ״ח (כפתורי הדפסה)",
        templatesIntro: "הוסף תבנית ממ״ח לכל סוג טופס. כל אחת תופיע ככפתור הדפסה נפרד במסך הישיבות, ותדפיס לכל מי שעבר את סינון התאריכים.",
      }, this),
    ];

    this._bindActions();
    this._init();
  }

  getSiteUrl() { return this.spSiteUrlEl.value.trim(); }
  getBaseFields() { return this.spFields; }

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
    this.config.fields.forEach((field) => this._addRow(field));
    if (!this.config.fields.length) this._addRow({});
  }

  _addRow(field = {}) {
    const tag = typeof field === "string" ? field : field.tag || "";
    const column = typeof field === "string" ? field : field.column || "";
    const source = typeof field === "object" && field.source === "system" ? "system" : "column";
    const internal = typeof field === "object" ? field.internal || "printDay" : "printDay";

    const row = document.createElement("div");
    row.className = "field-row quick-field-row";

    const tagInput = this._input("field-input tag", "שם התג בוורד", tag);
    const sourceSelect = document.createElement("select");
    sourceSelect.className = "field-input quick-source";
    sourceSelect.append(new Option("עמודה בטבלה", "column"), new Option("שדה מערכת (אוטומטי)", "system"));

    const columnInput = this._input("field-input column", "שם העמודה בטבלה", column);
    columnInput.setAttribute("list", "columnOptions");

    const systemSelect = document.createElement("select");
    systemSelect.className = "field-input quick-system";
    SystemDateService.FIELDS.forEach((item) => systemSelect.append(new Option(item.title, item.internal)));
    systemSelect.value = internal;

    const sync = () => {
      const isSystem = sourceSelect.value === "system";
      columnInput.hidden = isSystem;
      systemSelect.hidden = !isSystem;
    };
    sourceSelect.addEventListener("change", sync);
    sourceSelect.value = source;
    sync();

    row.append(tagInput, sourceSelect, columnInput, systemSelect, this._removeBtn(row));
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

    this.displayPicker.setFields(this.spFields);
    this.displayPicker.setOrder(this.spConfig.displayFields || []);

    this.itemPrintEditor.setBaseFields(this.spFields);
    this.itemPrintEditor.load(this.spConfig);
    this.editors.forEach((editor) => editor.load(this.spConfig));
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

  // Fills a select with the given fields (defaults to primary list), keeping the selected internal name.
  _populateSelect(select, selectedInternal, fields = this.spFields) {
    select.innerHTML = "";
    select.appendChild(new Option("— ללא —", ""));

    const known = new Set();
    fields.forEach((f) => {
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
      this.spFields = await this._loadFieldList(siteUrl, listTitle);
      this._refreshSelects();
      this._setFieldsInfo(`נטענו ${this.spFields.length} שדות`);
    } catch (err) {
      console.error(err);
      this._setFieldsInfo(`טעינת השדות נכשלה — ${err?.message || "בדוק כתובת והרשאות"}`);
    }
  }

  // Fetches a list's visible fields as sorted { title, internal } pairs (shared by base + editors).
  async _loadFieldList(siteUrl, listTitle) {
    const raw = await this._fetchFields(siteUrl, listTitle);
    return raw
      .filter((f) => !f.Hidden)
      .map((f) => ({ title: f.Title || f.InternalName, internal: f.InternalName || f.StaticName }))
      .filter((f) => f.internal)
      .sort((a, b) => String(a.title).localeCompare(String(b.title), "he"));
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

  async _fetchChoices(siteUrl, listTitle, fieldInternal) {
    const origin = new URL(siteUrl).origin;
    const tabs = await chrome.tabs.query({ url: `${origin}/*` });
    const tab = tabs.find((t) => t.url?.startsWith(siteUrl)) || tabs[0];
    if (!tab?.id) throw new Error("פתח לשונית SharePoint מחוברת ונסה שוב");

    const [{ result }] = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: async (url, title, field) => {
        const base = url.replace(/\/$/, "");
        const endpoint =
          `${base}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')` +
          `/fields/getbyinternalnameortitle('${encodeURIComponent(field)}')?$select=Choices&$format=json`;
        const res = await fetch(endpoint, { credentials: "include", headers: { Accept: "application/json;odata=verbose" } });
        const text = await res.text();
        if (!res.ok) throw new Error(`SharePoint choices ${res.status}`);
        if (text.trim().startsWith("<")) throw new Error("SharePoint החזיר HTML במקום JSON");
        const body = JSON.parse(text).d || JSON.parse(text);
        return body.Choices?.results || body.Choices || [];
      },
      args: [siteUrl, listTitle, fieldInternal],
    });
    return result || [];
  }

  _refreshSelects() {
    this._populateSelect(this.spMatchFieldEl, this.spMatchFieldEl.value);
    this.spListEl.querySelectorAll(".sp-field").forEach((select) =>
      this._populateSelect(select, select.value)
    );
    this.displayPicker.setFields(this.spFields);
    this.editors.forEach((editor) => editor.setBaseFields(this.spFields));
    this.itemPrintEditor.setBaseFields(this.spFields);
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
    this.listEl.querySelectorAll(".quick-field-row").forEach((row) => {
      if (!row.querySelector(".tag")) return;
      const tag = row.querySelector(".tag").value.trim();
      if (!tag) return;
      const source = row.querySelector(".quick-source")?.value || "column";
      if (source === "system") {
        fields.push({
          tag,
          source: "system",
          internal: row.querySelector(".quick-system")?.value || "printDay",
        });
        return;
      }
      const column = row.querySelector(".column").value.trim();
      fields.push({ tag, column: column || tag, source: "column" });
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
    const raw = {
      siteUrl: this.spSiteUrlEl.value.trim(),
      listTitle: this.spListTitleEl.value.trim(),
      matchTableColumn: this.spMatchColumnEl.value.trim(),
      matchListFieldInternal: this.spMatchFieldEl.value,
      matchListFieldDisplay: matchOption ? matchOption.textContent : "",
      fieldMap,
      displayFields: this.displayPicker.getOrder(),
    };
    this.editors.forEach((editor) => Object.assign(raw, editor.collect()));
    Object.assign(raw, this.itemPrintEditor.collect());
    return raw;
  }

  _bindActions() {
    document.getElementById("addField").addEventListener("click", () => this._addRow({}));
    document.getElementById("addSpField").addEventListener("click", () => this._addSpRow("", ""));
    document.getElementById("loadFields").addEventListener("click", () => this._loadFields());
    document.getElementById("save").addEventListener("click", () => this._save());
    document.getElementById("reset").addEventListener("click", () => this._reset());
    document.getElementById("exportSettings").addEventListener("click", () => this._export());
    document.getElementById("exportQr").addEventListener("click", () => this._exportQr());
    document.getElementById("importSettings").addEventListener("click", () =>
      document.getElementById("importFile").click()
    );
    document.getElementById("importFile").addEventListener("change", (e) => {
      this._import(e.target.files[0]);
      e.target.value = "";
    });
  }

  // Downloads a full JSON snapshot: form edits plus saved custom templates and mappings.
  async _export() {
    const data = await SettingsBackupStore.export({
      quickPrint: this._collectQuickPrint(),
      sharePoint: this._collectSharePoint(),
    });
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "הגדרות-תוסף-זימונים.json";
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    this._status("ההגדרות יוצאו לקובץ", "ok");
  }

  // Splits the same snapshot used by _export() into multiple QR codes for phone transfer.
  async _exportQr() {
    try {
      const data = await SettingsBackupStore.export({
        quickPrint: this._collectQuickPrint(),
        sharePoint: this._collectSharePoint(),
      });
      new QrExportModal().open(JSON.stringify(data));
      this._status("נוצרו קודי QR — סרוק אותם בדף הטלפון", "ok");
    } catch (err) {
      console.error(err);
      this._status(err?.message || "יצירת ה-QR נכשלה", "error");
    }
  }

  // Reads a full settings file, persists every setting, and reloads the form.
  async _import(file) {
    if (!file) return;
    try {
      await SettingsBackupStore.import(JSON.parse(await file.text()));
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
