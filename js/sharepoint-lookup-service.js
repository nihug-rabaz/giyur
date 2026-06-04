class SharePointLookupService {
  constructor(config, fieldMap) {
    this.config = config || {};
    this.fieldMap = fieldMap || {};
    this._fieldIndex = null;
  }

  // Pulls SP items, matches rows by join display names and injects mapped columns into each row.
  async enrich(headers, rows) {
    const resolved = await this._resolveConfig();
    if (!resolved.matchInternal || !Object.keys(resolved.fieldMap).length) {
      return { headers, rows, applied: false };
    }

    const source = this._findSourceColumn(headers, resolved.tableLabel);
    if (!source) return { headers, rows, applied: false };

    const index = this.buildIndex(await this.fetchItems(), resolved.matchInternal);
    const { headers: newHeaders, targets, created } = this._resolveTargets(headers, resolved.fieldMap);

    rows.forEach((row) => {
      const item = index.get(this._norm(row[source.id]));
      Object.entries(resolved.fieldMap).forEach(([tag, internalName]) => {
        const id = targets[tag];
        if (item) row[id] = this._readField(item, internalName);
        else if (created.has(id)) row[id] = "";
      });
    });

    return { headers: newHeaders, rows, applied: true };
  }

  // Maps each SharePoint tag to a target column id, reusing a same-named column or creating one,
  // so the value always lands under the id the template placeholder reads from.
  _resolveTargets(headers, fieldMap) {
    const newHeaders = [...headers];
    const targets = {};
    const created = new Set();
    Object.keys(fieldMap).forEach((tag) => {
      const existing = headers.find((h) => this._norm(h.label) === this._norm(tag));
      if (existing) {
        targets[tag] = existing.id;
      } else {
        targets[tag] = tag;
        created.add(tag);
        newHeaders.push({ id: tag, label: tag });
      }
    });
    return { headers: newHeaders, targets, created };
  }

  async _resolveConfig() {
    const join = this.config.join || {};
    const tableLabel = join.tableColumnDisplay || this.config.sourceColumnContains || "";

    const hasMatch = Boolean(join.listFieldInternal || this.config.matchField);
    const mapMissing = Object.values(this.fieldMap).some((v) => !v);
    const fieldIndex = hasMatch && !mapMissing ? new Map() : await this._getFieldIndex();

    const matchInternal =
      join.listFieldInternal ||
      this.config.matchField ||
      fieldIndex.get(join.listFieldDisplay) ||
      "";

    const fieldMap = {};
    Object.entries(this.fieldMap).forEach(([documentField, configured]) => {
      const internal =
        configured ||
        fieldIndex.get(documentField) ||
        "";
      if (internal) fieldMap[documentField] = internal;
    });

    return { tableLabel, matchInternal, fieldMap };
  }

  async _getFieldIndex() {
    if (this._fieldIndex) return this._fieldIndex;
    const map = new Map();
    try {
      const fields = await this.fetchFields();
      fields.forEach((f) => {
        const title = this._norm(f.Title);
        const internal = f.InternalName || f.StaticName;
        if (title && internal) map.set(title, internal);
      });
    } catch {
      // No field metadata (offline / blocked) — rely on internal names already in config.
    }
    this._fieldIndex = map;
    return map;
  }

  async fetchFields() {
    if (this._canFetchViaSharePointTab()) return this._fetchFieldsViaSharePointTab();
    const json = await this._fetchJson(this._fieldsUrl(), "SharePoint fields");
    const body = json.d || json;
    return body.results || body.value || [];
  }

  async _fetchFieldsViaSharePointTab() {
    const result = await this._withSharePointTab((tabId) =>
      this._execInTab(tabId, async (endpoint) => {
        const res = await fetch(endpoint, {
          credentials: "include",
          headers: { Accept: "application/json;odata=nometadata" },
        });
        const text = await res.text();
        if (!res.ok) throw new Error(`SharePoint fields ${res.status}`);
        if (text.trim().startsWith("<")) throw new Error("SharePoint החזיר HTML במקום JSON");
        const json = JSON.parse(text);
        return json.value || json.d?.results || [];
      }, [this._fieldsUrl()])
    );
    return result || [];
  }

  async fetchItems() {
    if (this._canFetchViaSharePointTab()) return this._fetchItemsViaSharePointTab();
    return this._fetchItemsDirect();
  }

  async _fetchItemsDirect() {
    const items = [];
    let url = `${this.config.listUrl.replace(/\/$/, "")}?$top=5000`;
    while (url) {
      const json = await this._fetchJson(url, "SharePoint items");
      const body = json.d || json;
      items.push(...(body.results || body.value || []));
      url = body.__next || json["odata.nextLink"] || null;
    }
    return items;
  }

  async _fetchItemsViaSharePointTab() {
    const itemsUrl = `${this.config.listUrl.replace(/\/$/, "")}?$top=5000`;
    const result = await this._withSharePointTab((tabId) =>
      this._execInTab(tabId, async (startUrl) => {
        async function fetchJson(endpoint) {
          const res = await fetch(endpoint, {
            credentials: "include",
            headers: { Accept: "application/json;odata=nometadata" },
          });
          const text = await res.text();
          if (!res.ok) throw new Error(`SharePoint items ${res.status}`);
          if (text.trim().startsWith("<")) throw new Error("SharePoint החזיר HTML במקום JSON");
          return JSON.parse(text);
        }

        const items = [];
        let url = startUrl;
        while (url) {
          const json = await fetchJson(url);
          items.push(...(json.value || json.d?.results || []));
          url = json["odata.nextLink"] || json.d?.__next || null;
        }
        return items;
      }, [itemsUrl])
    );
    return result || [];
  }

  // Matches the table's join column tolerantly (exact, then partial either direction).
  _findSourceColumn(headers, tableLabel) {
    const target = this._norm(tableLabel);
    if (!target) return null;
    return (
      headers.find((h) => this._norm(h.label) === target) ||
      headers.find((h) => {
        const label = this._norm(h.label);
        return label.includes(target) || target.includes(label);
      }) ||
      null
    );
  }

  buildIndex(items, matchInternal) {
    const map = new Map();
    items.forEach((item) => {
      const key = this._norm(this._readField(item, matchInternal));
      if (key) map.set(key, item);
    });
    return map;
  }

  _fieldsUrl() {
    const base = this.config.listUrl.replace(/\/items\/?(\?.*)?$/i, "");
    return `${base}/fields?$select=Title,InternalName,StaticName,Hidden,TypeAsString&$format=json`;
  }

  async _fetchJson(url, label) {
    const res = await fetch(url, {
      credentials: "include",
      headers: { Accept: "application/json;odata=verbose" },
    });
    const text = await res.text();
    if (!res.ok) throw new Error(`${label} ${res.status}`);
    if (text.trim().startsWith("<")) throw new Error("SharePoint החזיר HTML במקום JSON");
    return JSON.parse(text);
  }

  _canFetchViaSharePointTab() {
    return Boolean(globalThis.chrome?.tabs?.query && globalThis.chrome?.scripting?.executeScript);
  }

  // Runs a fetch inside a SharePoint tab (its authenticated session); opens one if none exists
  // and closes only tabs that were opened here, leaving the user's own tabs untouched.
  async _withSharePointTab(run) {
    const { tab, opened } = await this._resolveSharePointTab();
    try {
      return await run(tab.id);
    } finally {
      if (opened) {
        try { await chrome.tabs.remove(tab.id); } catch { /* tab already gone */ }
      }
    }
  }

  async _resolveSharePointTab() {
    const existing = await this._findExistingSharePointTab();
    if (existing?.id) return { tab: existing, opened: false };
    return { tab: await this._openSharePointTab(), opened: true };
  }

  async _findExistingSharePointTab() {
    const listUrl = this.config.listUrl.replace(/\/$/, "");
    const siteUrl = this._siteUrlFromListUrl(listUrl);
    const origin = new URL(listUrl).origin;
    const tabs = await chrome.tabs.query({ url: `${origin}/*` });
    return tabs.find((t) => siteUrl && t.url?.startsWith(siteUrl)) || tabs[0] || null;
  }

  async _openSharePointTab() {
    const listUrl = this.config.listUrl.replace(/\/$/, "");
    const target = this._siteUrlFromListUrl(listUrl) || new URL(listUrl).origin;
    const tab = await chrome.tabs.create({ url: target, active: true });
    await this._waitForTabComplete(tab.id);
    return tab;
  }

  _waitForTabComplete(tabId, timeoutMs = 20000) {
    return new Promise((resolve) => {
      let done = false;
      const finish = () => {
        if (done) return;
        done = true;
        chrome.tabs.onUpdated.removeListener(listener);
        clearTimeout(timer);
        resolve();
      };
      const listener = (id, info) => { if (id === tabId && info.status === "complete") finish(); };
      chrome.tabs.onUpdated.addListener(listener);
      const timer = setTimeout(finish, timeoutMs);
    });
  }

  async _execInTab(tabId, func, args) {
    const [{ result }] = await chrome.scripting.executeScript({ target: { tabId }, func, args });
    return result;
  }

  _siteUrlFromListUrl(listUrl) {
    try {
      const url = new URL(listUrl);
      const apiIndex = url.pathname.toLowerCase().indexOf("/_api/");
      if (apiIndex < 0) return url.origin;
      return `${url.origin}${url.pathname.slice(0, apiIndex)}`;
    } catch {
      return "";
    }
  }

  // Resolves a field value, falling back to a fuzzy key match when SharePoint renamed the internal name.
  _readField(item, internalName) {
    return this._normalizeValue(this._resolveItemValue(item, internalName));
  }

  _resolveItemValue(item, internalName) {
    if (!item || !internalName) return undefined;
    if (item[internalName] !== undefined) return item[internalName];
    const lower = internalName.toLowerCase();
    const altKey = Object.keys(item).find(
      (k) => k.toLowerCase().includes(lower) || lower.includes(k.toLowerCase())
    );
    return altKey ? item[altKey] : undefined;
  }

  // Flattens lookup / person / multi-value fields to plain text like the SharePoint console snippet.
  _normalizeValue(value) {
    if (value === null || value === undefined) return "";
    if (Array.isArray(value)) {
      return value.map((v) => this._scalarValue(v)).filter((v) => v !== "").join(", ");
    }
    if (typeof value === "object") return this._scalarValue(value);
    return value;
  }

  _scalarValue(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "object") {
      return v.Title ?? v.LookupValue ?? v.Name ?? v.Email ?? v.Value ?? v.Label ?? "";
    }
    return v;
  }

  _norm(value) {
    return String(value ?? "").trim();
  }
}

window.SharePointLookupService = SharePointLookupService;
