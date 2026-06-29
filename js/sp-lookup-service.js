class SharePointItemLookup {
  constructor() {
    this.service = null;
    this.config = null;
    this.fields = [];
  }

  // Loads the saved SharePoint config and prepares the shared lookup service.
  async init() {
    this.config = await SharePointConfigStore.get();
    this.service = new SharePointLookupService(this.config, {});
    return this.config;
  }

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
        readOnly: Boolean(f.ReadOnlyField),
        required: Boolean(f.Required),
      }))
      .filter((f) => f.internal)
      .sort((a, b) => String(a.title).localeCompare(String(b.title), "he"));
    await this.service.warmLookupCaches(this.fields);
    return this.fields;
  }

  // Fields a user can fill when creating an item (writable, excluding system content-type fields).
  creatableFields() {
    const skip = new Set(["ContentType", "Attachments"]);
    return this.fields.filter((f) => !f.readOnly && !skip.has(f.internal));
  }

  // Possible values of a lookup field as [{ id, label }] for dropdown population.
  async lookupOptions(field) {
    return this.service.fetchLookupOptions(field.title);
  }

  // Targeted server-side $filter first (scales to huge lists); falls back to an in-memory scan
  // for columns the server cannot filter (multi-value / non-indexed), so it never hangs.
  async search(searchInternal, value) {
    const target = this._norm(value);
    if (!target || !searchInternal) return [];
    const field = this.fields.find((f) => f.internal === searchInternal);
    const property = field?.entityProperty || searchInternal;
    try {
      return await this.service.searchItems(property, target, field?.type || "");
    } catch (err) {
      console.warn("Server search failed, scanning in memory:", err?.message);
      return this._searchInMemory(searchInternal, target);
    }
  }

  async _searchInMemory(searchInternal, target) {
    const items = await this.service.fetchItems();
    const exact = items.filter((it) => this._norm(this.valueOf(it, searchInternal)) === target);
    if (exact.length) return exact;
    return items.filter((it) => this._norm(this.valueOf(it, searchInternal)).includes(target));
  }

  valueOf(item, internalName) {
    return this.service.readValue(item, internalName);
  }

  async update(item, changes) {
    const id = item?.ID ?? item?.Id;
    if (!id) throw new Error("לא נמצא ID לפריט");
    await this.service.updateItem(id, changes);
    Object.assign(item, changes);
    return true;
  }

  // Creates a new list item in-tab (opens the site itself and reuses the browser session).
  async create(fields) {
    return this.service.createItem(fields);
  }

  _norm(value) {
    return String(value ?? "").trim();
  }
}

window.SharePointItemLookup = SharePointItemLookup;
