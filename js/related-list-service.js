class RelatedListService {
  constructor(profile) {
    this.profile = profile || {};
    this.service = new SharePointLookupService(
      { siteUrl: this.profile.siteUrl, listUrl: this.profile.listUrl },
      {}
    );
    this.fields = [];
  }

  get configured() {
    return Boolean(this.profile.listTitle && this.profile.listUrl);
  }

  // Loads the related list's visible fields (title/internal/entity/type) for display and filtering.
  async loadFields() {
    const raw = await this.service.fetchFields();
    this.fields = raw
      .filter((f) => !f.Hidden)
      .map((f) => ({
        title: f.Title || f.InternalName,
        internal: f.InternalName || f.StaticName,
        entityProperty: f.EntityPropertyName || f.InternalName || f.StaticName,
        type: f.TypeAsString || "",
        dateOnly: SharePointLookupService.isDateOnlyField(f),
      }))
      .filter((f) => f.internal)
      .sort((a, b) => String(a.title).localeCompare(String(b.title), "he"));
    await this.service.warmLookupCaches(this.fields);
    return this.fields;
  }

  // Returns items in the [from, to] range, optionally filtered by type (Choice) and locations.
  async search({ from, to, type, locations } = {}) {
    const prop = this._entityOf(this.profile.dateField);
    if (!prop) throw new Error("לא הוגדר שדה תאריך לסינון בהגדרות");
    const parts = [];
    if (from) parts.push(`${prop} ge datetime'${from}T00:00:00'`);
    if (to) parts.push(`${prop} le datetime'${to}T23:59:59'`);
    if (type?.field && type?.value) {
      parts.push(`${this._entityOf(type.field)} eq '${this._esc(type.value)}'`);
    }
    if (locations?.values?.length) {
      parts.push(`(${this._locationClause(locations.values)})`);
    }
    if (!parts.length) parts.push(`${prop} ne null`);
    const items = await this.service.queryItems(parts.join(" and "), 5000);
    return this._sortByDateTime(this._dedupeItems(items));
  }

  // Items in this related list whose configured Lookup field points to a given base-list item id.
  async searchByLookup(baseItemId, { type } = {}) {
    const field = this.profile.lookupField;
    if (!field) throw new Error("לא הוגדר שדה קישור (Lookup) בהגדרות");
    const id = Number(baseItemId);
    if (Number.isNaN(id)) return [];
    const typeSuffix = type?.field && type?.value
      ? ` and ${this._entityOf(type.field)} eq '${this._esc(type.value)}'`
      : "";
    const filters = this._lookupIdFilters(field, id).map((part) => `${part}${typeSuffix}`);
    let items;
    try {
      items = await this.service.queryItemsFirst(filters, 5000);
    } catch {
      items = await this._searchByLookupInMemory(field, id, type);
    }
    return this._sortByDateTime(this._dedupeItems(items));
  }

  async _searchByLookupInMemory(lookupField, baseItemId, type) {
    const items = await this.service.fetchItems();
    const targetId = Number(baseItemId);
    return items.filter((item) => {
      const linked = Number(this.lookupId(item, lookupField));
      if (linked !== targetId) return false;
      if (type?.field && type?.value) {
        return String(this.valueOf(item, type.field) ?? "") === type.value;
      }
      return true;
    });
  }

  _lookupIdFilters(internalName, id) {
    const props = this._lookupIdProperties(internalName);
    return props.map((prop) => `${prop} eq ${id}`);
  }

  _lookupIdProperties(internalName) {
    const field = this.fields.find((f) => f.internal === internalName);
    const internal = field?.internal || internalName;
    const entity = field?.entityProperty || internal;
    const props = [];
    const seen = new Set();
    const add = (base) => {
      const key = base ? `${base}Id` : "";
      if (key && !seen.has(key)) { seen.add(key); props.push(key); }
    };
    [internal, entity].forEach((name) => {
      if (/Text$/i.test(name)) add(name.replace(/Text$/i, ""));
    });
    add(internal);
    add(entity);
    return props;
  }

  _dedupeItems(items) {
    const seen = new Set();
    return items.filter((item) => {
      const id = item.ID ?? item.Id;
      if (id == null) return true;
      const key = String(id);
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  itemSortKey(item) {
    const dateVal = this.valueOf(item, this.profile.dateField);
    const timeField = this._sortTimeField();
    const timeVal = timeField ? this.valueOf(item, timeField) : "";
    if (timeField && timeVal) return this._combineDateAndTime(dateVal, timeVal);
    const ms = Date.parse(dateVal);
    return Number.isNaN(ms) ? 0 : ms;
  }

  _sortTimeField() {
    const fields = this.fields || [];
    const summonsTime = fields.find((field) =>
      field.internal !== this.profile.dateField && /שעת\s*הזימון/i.test(field.title)
    );
    if (summonsTime) return summonsTime.internal;
    return fields.find((field) => {
      if (field.internal === this.profile.dateField) return false;
      return /שעת|שעה/i.test(field.title) || /DateTime|Time/i.test(field.type);
    })?.internal || "";
  }

  // Merges a date-only value with a separate time field (e.g. "3:00" inside a longer text).
  _combineDateAndTime(dateVal, timeVal) {
    const base = this._dateBase(dateVal);
    if (!base) return 0;
    const parts = this._parseTimeParts(timeVal);
    if (!parts) return base.getTime();
    base.setHours(parts.h, parts.m, parts.s, 0);
    return base.getTime();
  }

  _dateBase(dateVal) {
    if (!dateVal) return null;
    const parsed = new Date(dateVal);
    if (!Number.isNaN(parsed.getTime())) {
      parsed.setHours(0, 0, 0, 0);
      return parsed;
    }
    const iso = String(dateVal).slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) return null;
    const fallback = new Date(`${iso}T00:00:00`);
    return Number.isNaN(fallback.getTime()) ? null : fallback;
  }

  _parseTimeParts(value) {
    if (value == null || value === "") return null;
    const str = String(value).trim();
    if (str.includes("T")) {
      const iso = new Date(str);
      if (!Number.isNaN(iso.getTime())) {
        return { h: iso.getHours(), m: iso.getMinutes(), s: iso.getSeconds() };
      }
    }
    const colon = str.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
    if (colon) {
      return { h: Number(colon[1]), m: Number(colon[2]), s: Number(colon[3] || 0) };
    }
    const hourOnly = str.match(/^(\d{1,2})$/);
    if (hourOnly) {
      const h = Number(hourOnly[1]);
      if (h >= 0 && h <= 23) return { h, m: 0, s: 0 };
    }
    const embedded = str.match(/(?:^|[^\d:])(\d{1,2})(?:[^\d:]|$)/);
    if (embedded) {
      const h = Number(embedded[1]);
      if (h >= 0 && h <= 23) return { h, m: 0, s: 0 };
    }
    return null;
  }

  _sortByDateTime(items) {
    return items.slice().sort((a, b) => this.itemSortKey(a) - this.itemSortKey(b));
  }

  _sortByDate(items) {
    return this._sortByDateTime(items);
  }

  // Builds the location $filter for a Lookup field (by id) or a Choice field (by text).
  _locationClause(values) {
    const field = this.profile.locationField;
    if (this._isLookup(field)) {
      const props = this._lookupIdProperties(field);
      const prop = props[0] || `${this._entityOf(field)}Id`;
      return values.map((v) => `${prop} eq ${Number(v)}`).join(" or ");
    }
    const entity = this._entityOf(field);
    return values.map((v) => `${entity} eq '${this._esc(v)}'`).join(" or ");
  }

  // Location options as [{ value, label }] — from the lookup target list, or the Choice values.
  async locationOptions() {
    const field = this.profile.locationField;
    if (!field) return [];
    if (this._isLookup(field)) {
      const title = this.fields.find((f) => f.internal === field)?.title || field;
      const opts = await this.service.fetchLookupOptions(title);
      return opts.map((o) => ({ value: String(o.id), label: o.label || String(o.id) }));
    }
    const choices = await this.service.fetchFieldChoices(field);
    return choices.map((c) => ({ value: c, label: c }));
  }

  // The location key of an item (lookup id as string, or the Choice text) for grouping/display.
  locationKeyOf(item) {
    const field = this.profile.locationField;
    if (!field) return "";
    if (this._isLookup(field)) {
      const id = this.lookupId(item, field);
      return id == null ? "" : String(id);
    }
    return String(this.valueOf(item, field) ?? "");
  }

  _isLookup(internal) {
    const type = this.fields.find((f) => f.internal === internal)?.type || "";
    return type === "Lookup" || type === "LookupMulti";
  }

  _esc(value) {
    return String(value).replace(/'/g, "''");
  }

  // Resolves the base-list item id this item points to via its configured Lookup field.
  lookupId(item, internalName) {
    for (const key of this._lookupIdProperties(internalName)) {
      if (item[key] != null) return item[key];
    }
    const entity = this._entityOf(internalName);
    const value = item[entity] ?? item[internalName];
    if (value == null) return null;
    if (Array.isArray(value)) return value[0]?.Id ?? value[0]?.ID ?? null;
    if (typeof value === "object") return value.Id ?? value.ID ?? null;
    return value;
  }

  _entityOf(internal) {
    const field = this.fields.find((f) => f.internal === internal);
    return field ? field.entityProperty : internal;
  }

  valueOf(item, internal) {
    return this.service.readValue(item, internal);
  }
}

window.RelatedListService = RelatedListService;
