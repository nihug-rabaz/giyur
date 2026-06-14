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
      }))
      .filter((f) => f.internal)
      .sort((a, b) => String(a.title).localeCompare(String(b.title), "he"));
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
    return this._sortByDate(items);
  }

  // Builds the location $filter for a Lookup field (by id) or a Choice field (by text).
  _locationClause(values) {
    const field = this.profile.locationField;
    const entity = this._entityOf(field);
    if (this._isLookup(field)) {
      return values.map((v) => `${entity}Id eq ${Number(v)}`).join(" or ");
    }
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

  _sortByDate(items) {
    const internal = this.profile.dateField;
    return items.slice().sort((a, b) => {
      const da = Date.parse(this.valueOf(a, internal)) || 0;
      const db = Date.parse(this.valueOf(b, internal)) || 0;
      return da - db;
    });
  }

  // Resolves the base-list item id this item points to via its configured Lookup field.
  lookupId(item, internalName) {
    const entity = this._entityOf(internalName);
    if (item[entity + "Id"] != null) return item[entity + "Id"];
    const value = item[entity];
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
