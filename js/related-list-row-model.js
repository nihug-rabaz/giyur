class RelatedListRowModel {
  constructor(profile, config = {}) {
    this.profile = profile || {};
    this.config = config;
    this.listService = new RelatedListService(profile);
    this.baseService = null;
    this.baseFields = [];
  }

  get configured() {
    return this.listService.configured;
  }

  async loadFields() {
    await this.listService.loadFields();
    return this.listService.fields;
  }

  previewColumns(listFields, baseFields = []) {
    return this.columnsFromLayout(listFields, baseFields, this.profile.columnLayout || []).columns;
  }

  columnsFromLayout(listFields, baseFields, layout) {
    if (!layout.length) {
      const columns = this.computeColumns(listFields, this.profile.displayFields || []);
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

  computeColumns(fields, displayFields) {
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

  // Builds display rows and enriched print rows for a set of list items.
  async buildForItems(items, { fixedBaseItem } = {}) {
    const listRead = (it, internal) => this.listService.valueOf(it, internal);
    const layout = this.profile.columnLayout || [];

    if (fixedBaseItem) {
      const wantsBaseInLayout = layout.some((col) => col.source === "base");
      const wantsBaseMapping = Object.values(this.profile.templateFieldMap || {})
        .some((field) => field.source === "base");
      if (wantsBaseInLayout || wantsBaseMapping) await this.ensureBaseFields();
      const { columns } = this.columnsFromLayout(
        this.listService.fields,
        wantsBaseInLayout ? this.baseFields : [],
        layout
      );
      const baseRead = wantsBaseInLayout || wantsBaseMapping
        ? (it, internal) => this.baseService.readValue(it, internal)
        : null;
      const rows = items.map((item) => {
        const baseItem = wantsBaseInLayout || wantsBaseMapping ? fixedBaseItem : null;
        const row = this.rowFromPlan(columns, item, { listRead, baseItem, baseRead });
        return { row, printRow: this.withTemplateMappings(row, { listItem: item, baseItem, listRead, baseRead }) };
      });
      return { columns, rows: rows.map((x) => x.row), printRows: rows.map((x) => x.printRow) };
    }

    const wantsBaseInLayout = layout.some((col) => col.source === "base");
    const wantsBaseMapping = Object.values(this.profile.templateFieldMap || {})
      .some((field) => field.source === "base");
    const wantsBase = wantsBaseInLayout || wantsBaseMapping;
    const lookupField = this.profile.lookupField;

    if (!wantsBase) {
      const { columns } = this.columnsFromLayout(this.listService.fields, [], layout);
      const rows = items.map((item) => {
        const row = this.rowFromPlan(columns, item, { listRead });
        return { row, printRow: this.withTemplateMappings(row, { listItem: item, listRead }) };
      });
      return { columns, rows: rows.map((x) => x.row), printRows: rows.map((x) => x.printRow) };
    }

    if (!lookupField) {
      const baseFieldsForLayout = wantsBaseInLayout ? await this._baseFieldsForLayout() : [];
      const { columns } = this.columnsFromLayout(this.listService.fields, baseFieldsForLayout, layout);
      const rows = items.map((item) => {
        const row = this.rowFromPlan(columns, item, { listRead });
        return { row, printRow: this.withTemplateMappings(row, { listItem: item, listRead }) };
      });
      return {
        columns,
        rows: rows.map((x) => x.row),
        printRows: rows.map((x) => x.printRow),
        warning: "להוספת עמודות בסיס הגדר שדה קישור (Lookup).",
      };
    }

    await this.ensureBaseFields();
    const { columns } = this.columnsFromLayout(this.listService.fields, this.baseFields, layout);
    const baseMap = await this.baseMapFor(items, lookupField);
    const baseRead = (it, internal) => this.baseService.readValue(it, internal);
    const rows = items.map((item) => {
      const baseItem = baseMap.get(Number(this.listService.lookupId(item, lookupField)));
      const row = this.rowFromPlan(columns, item, { listRead, baseItem, baseRead });
      return {
        row,
        printRow: this.withTemplateMappings(row, { listItem: item, baseItem, listRead, baseRead }),
      };
    });
    return { columns, rows: rows.map((x) => x.row), printRows: rows.map((x) => x.printRow) };
  }

  async baseMapFor(items, lookupField) {
    const ids = items.map((it) => this.listService.lookupId(it, lookupField)).filter((id) => id != null);
    const people = await this.baseService.fetchItemsByIds(ids);
    return new Map(people.map((p) => [Number(p.ID ?? p.Id), p]));
  }

  async ensureBaseFields() {
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
    await this.baseService.warmLookupCaches(this.baseFields);
  }

  async _baseFieldsForLayout() {
    await this.ensureBaseFields();
    return this.baseFields;
  }

  _formatCell(value, column) {
    return SharePointLookupService.formatValue(value, {
      fieldType: column?.fieldType,
      dateOnly: column?.dateOnly,
    });
  }

  rowFromPlan(columns, item, sources) {
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

  _fieldMeta(source, internal) {
    const fields = source === "base" ? (this.baseFields || []) : (this.listService?.fields || []);
    return fields.find((f) => f.internal === internal) || null;
  }

  withTemplateMappings(row, sources) {
    const printRow = { ...row };
    Object.entries(this.profile.templateFieldMap || {}).forEach(([tag, field]) => {
      if (field.source === "system") return;
      const source = field.source === "base" ? "base" : "list";
      const item = source === "base" ? sources.baseItem : sources.listItem;
      const read = source === "base" ? sources.baseRead : sources.listRead;
      const meta = this._fieldMeta(source, field.internal);
      printRow[tag] = item && read
        ? SharePointLookupService.formatValue(read(item, field.internal), SharePointLookupService.formatOptionsForField(meta))
        : "";
    });
    return printRow;
  }

  templatePrintFields() {
    return Object.entries(this.profile.templateFieldMap || {}).map(([tag, field]) => {
      if (field.source === "system") return { tag, source: "system", internal: field.internal };
      return { tag, column: tag, source: "column" };
    });
  }

  printHeaders(columns, fields) {
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
}

window.RelatedListRowModel = RelatedListRowModel;
