class StatusBar {
  constructor(el) {
    this.el = el;
  }

  set(text, type = "") {
    this.el.textContent = text;
    this.el.className = type ? `visible ${type}` : "visible";
  }

  clear() {
    this.el.textContent = "";
    this.el.className = "";
  }
}

class MappingPanel {
  constructor(templateId, templateNum, headers, container) {
    this.templateId = templateId;
    this.templateNum = templateNum;
    this.headers = headers;
    this.container = container;
  }

  async render() {
    let placeholders = [];
    try {
      const buffer = await TemplateStore.getBuffer(this.templateId);
      placeholders = await DocxTemplateEngine.extractPlaceholders(buffer);
    } catch {
      placeholders = [];
    }

    const saved = await TemplateStore.getMapping(this.templateId);
    const mapping = {};
    placeholders.forEach((p) => {
      mapping[p] = saved[p] ?? PlaceholderMapper.autoMatch(p, this.headers);
    });
    await TemplateStore.saveMapping(this.templateId, mapping);

    this.container.appendChild(this._buildPanel(placeholders, mapping));
  }

  _buildPanel(placeholders, mapping) {
    const panel = document.createElement("div");
    panel.className = "mapping-panel";

    const title = document.createElement("h3");
    title.textContent = `מיפוי שדות — תבנית ${this.templateNum}`;
    panel.appendChild(title);

    if (!placeholders.length) {
      const empty = document.createElement("p");
      empty.className = "mapping-empty";
      empty.textContent = "לא נמצאו שדות [...] בתבנית.";
      panel.appendChild(empty);
      return panel;
    }

    const grid = document.createElement("div");
    grid.className = "mapping-grid";
    placeholders.forEach((p) => grid.appendChild(this._buildField(p, mapping[p])));
    panel.appendChild(grid);
    return panel;
  }

  _buildField(placeholder, selectedId) {
    const field = document.createElement("div");
    field.className = "mapping-field";

    const label = document.createElement("label");
    label.textContent = `[${placeholder}]`;
    field.appendChild(label);

    const select = document.createElement("select");
    select.appendChild(new Option("— ללא —", ""));
    this.headers.forEach((h) => {
      const opt = new Option(h.label, h.id);
      if (h.id === selectedId) opt.selected = true;
      select.appendChild(opt);
    });
    select.addEventListener("change", () => this._onChange(placeholder, select.value));
    field.appendChild(select);
    return field;
  }

  async _onChange(placeholder, columnId) {
    const mapping = await TemplateStore.getMapping(this.templateId);
    mapping[placeholder] = columnId;
    await TemplateStore.saveMapping(this.templateId, mapping);
  }
}

const TEMPLATE_CONFIGS = [
  { id: "template1", num: 1, label: "תבנית 1 (שלום)" },
  { id: "template2", num: 2, label: "תבנית 2 (תבנית שניה)" },
  { id: "template3", num: 3, label: "תבנית 3" },
];

class SingleExportPage {
  constructor() {
    this.status = new StatusBar(document.getElementById("status"));
    this.headers = [];
    this.rows = [];
    this.service = null;
    this.templates = TEMPLATE_CONFIGS;
    this._buildUploads();
    this._buildToolbar();
    this._load();
  }

  async _load() {
    const data = await chrome.storage.local.get("summonsTableData");
    const payload = data.summonsTableData;
    if (!payload?.headers?.length || !payload?.rows?.length) {
      document.getElementById("summary").textContent = "לא נמצאו נתוני טבלה";
      document.getElementById("emptyState").hidden = false;
      document.getElementById("summonsTable").hidden = true;
      return;
    }
    this.headers = payload.headers;
    this.rows = payload.rows;
    await this._enrichFromSharePoint();
    this.service = new SummonsExportService(this.headers);
    document.getElementById("summary").textContent = `${this.rows.length} מזומנים | ${this.headers.length} עמודות`;
    this._renderTable();
    await this._renderMappingPanels();
    await this._maybeAutoPrint();
  }

  async _enrichFromSharePoint() {
    if (!window.SharePointLookupService) return;
    try {
      this.status.set("שואב נתונים מ-SharePoint...", "info");
      const config = window.SharePointConfigStore
        ? await SharePointConfigStore.get()
        : window.SHAREPOINT_LOOKUP;
      const fieldMap = config?.fieldMap || window.SHAREPOINT_FIELD_MAP;
      const service = new SharePointLookupService(config, fieldMap);
      const result = await service.enrich(this.headers, this.rows);
      this.headers = result.headers;
      this.rows = result.rows;
      this.status.set(result.applied ? "נתוני SharePoint נוספו" : "", result.applied ? "ok" : "");
    } catch {
      this.status.set("שאיבת נתונים מ-SharePoint נכשלה — ממשיך בלי השדות הנוספים", "error");
    }
  }

  async _maybeAutoPrint() {
    const params = new URLSearchParams(location.search);
    if (params.get("quickprint")) return this._runQuickPrint();

    const templateId = params.get("autoprint");
    if (!templateId || !this.rows.length) return;
    const template = this.templates.find((t) => t.id === templateId);
    if (!template) return;
    await this._run(() => this.service.printAll(template.id, this.rows), `מכין הדפסה לתבנית ${template.num}...`);
  }

  async _runQuickPrint() {
    if (!window.QuickPrintService || !this.rows.length) return;
    const config = window.QuickPrintConfigStore
      ? await QuickPrintConfigStore.get()
      : window.QUICK_PRINT_CONFIG;
    const service = new QuickPrintService(config, this.headers);
    await this._run(() => service.printAll(this.rows), "מכין הדפסה מהירה...");
  }

  async _renderMappingPanels() {
    const container = document.getElementById("mappingPanels");
    container.innerHTML = "";
    for (const t of this.templates) {
      await new MappingPanel(t.id, t.num, this.headers, container).render();
    }
  }

  _renderTable() {
    const head = document.getElementById("tableHead");
    const body = document.getElementById("tableBody");
    head.innerHTML = "";
    body.innerHTML = "";

    const headRow = document.createElement("tr");
    this.headers.forEach((h) => {
      const th = document.createElement("th");
      th.textContent = h.label;
      headRow.appendChild(th);
    });
    const actionsTh = document.createElement("th");
    actionsTh.textContent = "פעולות";
    headRow.appendChild(actionsTh);
    head.appendChild(headRow);

    this.rows.forEach((row, index) => {
      const tr = document.createElement("tr");
      this.headers.forEach((h) => {
        const td = document.createElement("td");
        td.textContent = row[h.id] ?? "";
        tr.appendChild(td);
      });
      const tdActions = document.createElement("td");
      tdActions.appendChild(this._buildRowActions(row, index));
      tr.appendChild(tdActions);
      body.appendChild(tr);
    });
  }

  _buildRowActions(row, index) {
    const wrap = document.createElement("div");
    wrap.className = "actions";
    this.templates.forEach((t) => {
      wrap.appendChild(this._button(`הפק ת${t.num}`, "btn btn--sm btn--primary", () => this._run(() => this.service.downloadOne(t.id, row, index, t.num))));
      wrap.appendChild(this._button(`הדפס ת${t.num}`, "btn btn--sm btn--accent", () => this._run(() => this.service.printOne(t.id, row, index))));
    });
    return wrap;
  }

  _button(label, classes, handler) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = classes;
    btn.textContent = label;
    btn.addEventListener("click", handler);
    return btn;
  }

  _buildToolbar() {
    const toolbar = document.getElementById("toolbar");
    toolbar.innerHTML = "";
    this.templates.forEach((t) => {
      const group = document.createElement("div");
      group.className = "toolbar-group";
      const span = document.createElement("span");
      span.textContent = `תבנית ${t.num}:`;
      group.appendChild(span);
      group.appendChild(this._button("הפק את כולם", "btn btn--primary", () =>
        this._run(() => this.service.downloadAll(t.id, this.rows, t.num), `מפיק את כל הקבצים לתבנית ${t.num}...`)
      ));
      group.appendChild(this._button("הדפס את כולם", "btn btn--accent", () =>
        this._run(() => this.service.printAll(t.id, this.rows), `מכין הדפסה לתבנית ${t.num}...`)
      ));
      toolbar.appendChild(group);
    });
  }

  _buildUploads() {
    const row = document.getElementById("uploadRow");
    row.innerHTML = "";
    this.templates.forEach((t) => {
      const item = document.createElement("div");
      item.className = "upload-item";

      const label = document.createElement("label");
      label.textContent = `החלפת ${t.label}`;

      const input = document.createElement("input");
      input.type = "file";
      input.accept = ".docx";
      input.addEventListener("change", (e) => this._handleUpload(t, e.target.files[0]));

      const reset = this._button("איפוס לברירת מחדל", "btn btn--sm", () => this._resetTemplate(t));

      item.append(label, input, reset);
      row.appendChild(item);
    });
  }

  async _handleUpload(template, file) {
    if (!file) return;
    try {
      const buffer = await file.arrayBuffer();
      await TemplateStore.saveCustom(template.id, buffer);
      await TemplateStore.saveMapping(template.id, {});
      await this._renderMappingPanels();
      this.status.set(`תבנית ${template.num} עודכנה`, "ok");
    } catch {
      this.status.set("שגיאה בשמירת התבנית", "error");
    }
  }

  async _resetTemplate(template) {
    await TemplateStore.clearCustom(template.id);
    await TemplateStore.saveMapping(template.id, {});
    await this._renderMappingPanels();
    this.status.set(`תבנית ${template.num} אופסה לברירת מחדל`, "ok");
  }

  async _run(fn, busyText = "מעבד...") {
    try {
      this.status.set(busyText, "info");
      document.querySelectorAll(".btn").forEach((b) => (b.disabled = true));
      await fn();
      this.status.set("בוצע בהצלחה", "ok");
    } catch (err) {
      console.error(err);
      this.status.set(err?.message || "שגיאה בביצוע הפעולה", "error");
    } finally {
      document.querySelectorAll(".btn").forEach((b) => (b.disabled = false));
    }
  }
}

new SingleExportPage();
