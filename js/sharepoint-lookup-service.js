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

  // Tries the richer query (with Required/ReadOnlyField) first, falling back to the minimal
  // proven query if the server rejects the extra fields, so field loading never breaks.
  async fetchFields() {
    if (this._canFetchViaSharePointTab()) {
      const rich = await this._fetchArrayViaTab(this._fieldsUrl(true), "SharePoint fields").catch(() => null);
      if (rich && rich.length) return rich;
      return this._fetchArrayViaTab(this._fieldsUrl(false), "SharePoint fields");
    }
    const json = await this._fetchJson(this._fieldsUrl(false), "SharePoint fields");
    const body = json.d || json;
    return body.results || body.value || [];
  }

  async _fetchArrayViaTab(url, label) {
    const result = await this._withSharePointTab((tabId) =>
      this._execInTab(tabId, async (endpoint, errLabel) => {
        const res = await fetch(endpoint, {
          credentials: "include",
          headers: { Accept: "application/json;odata=nometadata" },
        });
        const text = await res.text();
        if (!res.ok) throw new Error(`${errLabel} ${res.status}`);
        if (text.trim().startsWith("<")) throw new Error("SharePoint החזיר HTML במקום JSON");
        const json = JSON.parse(text);
        return json.value || json.d?.results || [];
      }, [url, label])
    );
    return result || [];
  }

  async fetchItems() {
    if (this._canFetchViaSharePointTab()) return this._fetchItemsViaSharePointTab();
    return this._fetchItemsDirect();
  }

  // Targeted server-side lookup by a single column; returns only matching items (scales to huge lists).
  async searchItems(propertyName, value, type) {
    if (!propertyName) return [];
    const candidates = this._filterCandidates(propertyName, value, type);
    if (this._canFetchViaSharePointTab()) return this._searchItemsViaTab(candidates);
    return this._searchItemsDirect(candidates);
  }

  // Runs an arbitrary OData $filter server-side (e.g. a date range) and returns up to `top` items.
  async queryItems(filter, top = 200) {
    if (!filter) return [];
    if (this._canFetchViaSharePointTab()) return this._searchItemsViaTab([filter], top);
    return this._searchItemsDirect([filter], top);
  }

  // Fetches list items by their ids, chunked into OR-filters so big id sets stay within URL limits.
  async fetchItemsByIds(ids, chunkSize = 30) {
    const unique = [...new Set((ids || []).map(Number).filter((n) => !Number.isNaN(n)))];
    if (!unique.length) return [];
    const out = [];
    for (let i = 0; i < unique.length; i += chunkSize) {
      const filter = unique.slice(i, i + chunkSize).map((id) => `Id eq ${id}`).join(" or ");
      out.push(...(await this.queryItems(filter, chunkSize)));
    }
    return out;
  }

  // Produces $filter variants to try; numeric columns drop the quotes, text columns keep them.
  _filterCandidates(prop, value, type) {
    const raw = String(value).trim();
    const quoted = `${prop} eq '${raw.replace(/'/g, "''")}'`;
    const unquoted = `${prop} eq ${raw}`;
    const numericTypes = ["Number", "Counter", "Integer", "Currency"];
    const looksNumeric = /^-?\d+(\.\d+)?$/.test(raw);
    if (!looksNumeric) return [quoted];
    return numericTypes.includes(type) ? [unquoted, quoted] : [quoted, unquoted];
  }

  async _searchItemsViaTab(candidates, top = 50) {
    const base = this.config.listUrl.replace(/\/$/, "");
    const outcome = await this._withSharePointTab((tabId) =>
      this._execInTab(tabId, async (listBase, filters, topN) => {
        async function readJson(url) {
          const controller = new AbortController();
          const timer = setTimeout(() => controller.abort(), 20000);
          try {
            const res = await fetch(url, {
              credentials: "include",
              headers: { Accept: "application/json;odata=nometadata" },
              signal: controller.signal,
            });
            const text = await res.text();
            if (!res.ok) throw new Error(text ? text.slice(0, 300) : `SharePoint ${res.status}`);
            if (text.trim().startsWith("<")) throw new Error("SharePoint החזיר HTML במקום JSON");
            return JSON.parse(text);
          } catch (e) {
            if (e.name === "AbortError") throw new Error("timeout");
            throw e;
          } finally {
            clearTimeout(timer);
          }
        }
        let lastError = "";
        for (const filter of filters) {
          try {
            const url = `${listBase}?$top=${topN}&$filter=${encodeURIComponent(filter)}`;
            const json = await readJson(url);
            return { items: json.value || json.d?.results || [] };
          } catch (e) {
            lastError = e && e.message ? e.message : String(e);
          }
        }
        return { error: lastError };
      }, [base, candidates, top])
    );
    if (outcome?.error) {
      const hint = outcome.error === "timeout"
        ? "החיפוש איטי מדי (כנראה עמודה לא מאונדקסת ברשימה גדולה, או עמודה מרובת-ערכים)."
        : outcome.error;
      throw new Error(`חיפוש נכשל — ${hint} ודא שעמודת החיפוש היא ערך יחיד ומסומנת כ-Indexed.`);
    }
    return outcome?.items || [];
  }

  async _searchItemsDirect(candidates, top = 50) {
    let lastError;
    for (const filter of candidates) {
      try {
        const url = `${this.config.listUrl.replace(/\/$/, "")}?$top=${top}&$filter=${encodeURIComponent(filter)}`;
        const json = await this._fetchJson(url, "SharePoint search");
        const body = json.d || json;
        return body.results || body.value || [];
      } catch (e) {
        lastError = e;
      }
    }
    if (lastError) throw lastError;
    return [];
  }

  // Returns [{ id, label }] of every item in a lookup field's target list, for dropdown population.
  async fetchLookupOptions(fieldTitle) {
    if (!fieldTitle || !this._canFetchViaSharePointTab()) return [];
    const baseUrl = this._listBaseUrl();
    const webUrl = this._siteUrlFromListUrl(this.config.listUrl);
    const result = await this._withSharePointTab((tabId) =>
      this._execInTab(tabId, async (base, web, title) => {
        const flat = { Accept: "application/json;odata=nometadata" };
        async function readJson(url) {
          const res = await fetch(url, { credentials: "include", headers: flat });
          const text = await res.text();
          if (!res.ok) throw new Error(text || `SharePoint ${res.status}`);
          return text ? JSON.parse(text) : {};
        }
        const metaJson = await readJson(`${base}/fields/getbyinternalnameortitle('${encodeURIComponent(title)}')?$select=LookupList,LookupField&$format=json`);
        const meta = metaJson.d || metaJson;
        const listGuid = String(meta.LookupList || "").replace(/[{}]/g, "");
        const col = meta.LookupField || "Title";
        if (!listGuid) return [];
        const options = [];
        let url = `${web}/_api/web/lists(guid'${listGuid}')/items?$select=Id,${encodeURIComponent(col)}&$top=5000`;
        while (url) {
          const json = await readJson(url);
          (json.value || json.d?.results || []).forEach((it) => {
            options.push({ id: it.Id ?? it.ID, label: String(it[col] ?? "") });
          });
          url = json["odata.nextLink"] || json.d?.__next || null;
        }
        return options;
      }, [baseUrl, webUrl, fieldTitle])
    );
    return (result || []).filter((o) => o.id != null);
  }

  // Returns the allowed values of a Choice field (for building filter checkboxes / tabs).
  async fetchFieldChoices(fieldNameOrTitle) {
    if (!fieldNameOrTitle || !this._canFetchViaSharePointTab()) return [];
    const base = this._listBaseUrl();
    const result = await this._withSharePointTab((tabId) =>
      this._execInTab(tabId, async (baseUrl, field) => {
        const url = `${baseUrl}/fields/getbyinternalnameortitle('${encodeURIComponent(field)}')?$select=Choices&$format=json`;
        const res = await fetch(url, { credentials: "include", headers: { Accept: "application/json;odata=nometadata" } });
        const text = await res.text();
        if (!res.ok) throw new Error(text || `SharePoint ${res.status}`);
        const body = text ? JSON.parse(text) : {};
        return body.Choices?.results || body.Choices || [];
      }, [base, fieldNameOrTitle])
    );
    return result || [];
  }

  async updateItem(itemId, changes) {
    if (!itemId || !changes || !Object.keys(changes).length) return false;
    if (!this._canFetchViaSharePointTab()) throw new Error("עריכה דורשת לשונית SharePoint מאומתת");
    return this._updateItemViaSharePointTab(itemId, changes);
  }

  // Creates a new list item from { internalFieldName: value } using the authenticated tab's session.
  async createItem(fields) {
    if (!fields || !Object.keys(fields).length) throw new Error("אין שדות ליצירת פריט");
    if (!this._canFetchViaSharePointTab()) throw new Error("יצירה דורשת לשונית SharePoint מאומתת");
    return this._createItemViaSharePointTab(fields);
  }

  async _createItemViaSharePointTab(fields) {
    const itemsUrl = this.config.listUrl.replace(/\/$/, "");
    const listBaseUrl = this._listBaseUrl();
    const siteUrl = this._siteUrlFromListUrl(this.config.listUrl);
    const digestUrl = `${siteUrl}/_api/contextinfo`;
    const outcome = await this._withSharePointTab((tabId) =>
      this._execInTab(tabId, async (createUrl, baseUrl, contextUrl, webUrl, body) => {
        const flat = { Accept: "application/json;odata=nometadata" };
        const out = { step: "digest", createUrl, skipped: [] };
        async function readJson(url, headers = flat) {
          const res = await fetch(url, { credentials: "include", headers });
          const text = await res.text();
          if (!res.ok) throw new Error(text || `SharePoint ${res.status}`);
          return text ? JSON.parse(text) : {};
        }
        // Builds name -> { entity, type } for every field, accepting internal/static/title/entity keys.
        async function loadFieldInfo() {
          const urls = [
            `${baseUrl}/fields?$select=Title,InternalName,StaticName,EntityPropertyName,TypeAsString&$format=json`,
            `${baseUrl}/fields?$format=json`,
          ];
          for (const url of urls) {
            try {
              const json = await readJson(url);
              const fields = json.value || json.d?.results || [];
              const map = {};
              fields.forEach((field) => {
                const entity = field.EntityPropertyName;
                if (!entity) return;
                const info = { entity, type: field.TypeAsString || "", title: field.Title || entity };
                [field.InternalName, field.StaticName, field.Title, field.EntityPropertyName]
                  .filter(Boolean)
                  .forEach((name) => { map[name] = info; });
              });
              if (Object.keys(map).length) return map;
            } catch {}
          }
          return {};
        }
        // Reads the lookup field's target list GUID and display column.
        async function lookupMeta(info) {
          const url = `${baseUrl}/fields/getbyinternalnameortitle('${encodeURIComponent(info.title)}')?$select=LookupList,LookupField&$format=json`;
          const d = (await readJson(url)).d || (await readJson(url));
          return { list: String(d.LookupList || "").replace(/[{}]/g, ""), field: d.LookupField || "Title" };
        }
        // Resolves a lookup value to the target item id (numeric input is treated as the id directly).
        async function resolveLookupId(meta, value) {
          const raw = String(value).trim();
          if (/^\d+$/.test(raw)) return Number(raw);
          if (!meta.list) return null;
          const col = meta.field || "Title";
          const q = `${webUrl}/_api/web/lists(guid'${meta.list}')/items?$select=Id,${encodeURIComponent(col)}&$filter=${encodeURIComponent(col)} eq '${raw.replace(/'/g, "''")}'&$top=1`;
          const items = (await readJson(q)).value || [];
          return items.length ? (items[0].Id ?? items[0].ID) : null;
        }
        async function buildPayload(input, info) {
          const payload = {};
          for (const [key, value] of Object.entries(input || {})) {
            const fi = info[key];
            const entity = fi ? fi.entity : key;
            const type = fi ? fi.type : "";
            if (type === "Lookup") {
              const id = await resolveLookupId(await lookupMeta(fi), value);
              if (id == null) { out.skipped.push(fi ? fi.title : key); continue; }
              payload[entity + "Id"] = id;
            } else if (type === "LookupMulti") {
              const meta = await lookupMeta(fi);
              const ids = [];
              for (const part of String(value).split(",").map((s) => s.trim()).filter(Boolean)) {
                const id = await resolveLookupId(meta, part);
                if (id != null) ids.push(id);
              }
              if (!ids.length) { out.skipped.push(fi ? fi.title : key); continue; }
              payload[entity + "Id"] = { __metadata: { type: "Collection(Edm.Int32)" }, results: ids };
            } else if (type === "User") {
              const raw = String(value).trim();
              if (!/^\d+$/.test(raw)) { out.skipped.push(fi ? fi.title : key); continue; }
              payload[entity + "Id"] = Number(raw);
            } else {
              payload[entity] = value;
            }
          }
          return payload;
        }
        try {
          const digestRes = await fetch(contextUrl, { method: "POST", credentials: "include", headers: flat });
          const digestText = await digestRes.text();
          out.digestStatus = digestRes.status;
          if (!digestRes.ok) { out.error = digestText.slice(0, 600); return out; }
          const digestJson = JSON.parse(digestText);
          const digest =
            digestJson.FormDigestValue ||
            digestJson.GetContextWebInformation?.FormDigestValue ||
            digestJson.d?.GetContextWebInformation?.FormDigestValue;
          if (!digest) { out.error = "לא התקבל X-RequestDigest"; return out; }

          out.step = "type";
          let typeName = "";
          try {
            const tRes = await fetch(`${baseUrl}?$select=ListItemEntityTypeFullName`, { credentials: "include", headers: flat });
            const tJson = JSON.parse(await tRes.text());
            typeName = tJson.ListItemEntityTypeFullName || tJson.d?.ListItemEntityTypeFullName || "";
          } catch { /* proceed without explicit entity type */ }
          out.typeName = typeName;

          out.step = "fields";
          const info = await loadFieldInfo();
          const normalizedBody = await buildPayload(body, info);
          out.payloadKeys = Object.keys(normalizedBody);

          out.step = "create";
          const payload = typeName ? { __metadata: { type: typeName }, ...normalizedBody } : { ...normalizedBody };
          const res = await fetch(createUrl, {
            method: "POST",
            credentials: "include",
            headers: {
              Accept: "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose",
              "X-RequestDigest": digest,
            },
            body: JSON.stringify(payload),
          });
          const text = await res.text();
          out.createStatus = res.status;
          out.ok = res.ok;
          if (!res.ok) { out.error = text.slice(0, 1000); return out; }
          let json = {};
          try { json = text ? JSON.parse(text) : {}; } catch { /* empty body */ }
          out.item = json.d || json;
          return out;
        } catch (e) {
          out.error = e && e.message ? e.message : String(e);
          return out;
        }
      }, [itemsUrl, listBaseUrl, digestUrl, siteUrl, fields])
    );

    if (!outcome || !outcome.ok) {
      console.error("SharePoint create failed:", outcome);
      const where = outcome?.step || "unknown";
      const status = outcome?.createStatus ?? outcome?.digestStatus ?? "";
      const detail = outcome?.error || "ללא פירוט";
      throw new Error(`כשל בשלב ${where}${status ? ` (${status})` : ""}: ${detail}`);
    }
    const item = outcome.item && typeof outcome.item === "object" ? outcome.item : {};
    if (outcome.skipped?.length) item._skipped = outcome.skipped;
    return item;
  }

  async _updateItemViaSharePointTab(itemId, changes) {
    const itemUrl = this._itemUrl(itemId);
    const digestUrl = `${this._siteUrlFromListUrl(this.config.listUrl)}/_api/contextinfo`;
    return this._withSharePointTab((tabId) =>
      this._execInTab(tabId, async (url, contextUrl, body) => {
        const digestRes = await fetch(contextUrl, {
          method: "POST",
          credentials: "include",
          headers: { Accept: "application/json;odata=nometadata" },
        });
        const digestText = await digestRes.text();
        if (!digestRes.ok) throw new Error(`SharePoint digest ${digestRes.status}`);
        const digestJson = JSON.parse(digestText);
        const digest =
          digestJson.FormDigestValue ||
          digestJson.GetContextWebInformation?.FormDigestValue ||
          digestJson.d?.GetContextWebInformation?.FormDigestValue;
        if (!digest) throw new Error("לא התקבל X-RequestDigest מ-SharePoint");

        const res = await fetch(url, {
          method: "POST",
          credentials: "include",
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": digest,
            "X-HTTP-Method": "MERGE",
            "IF-MATCH": "*",
          },
          body: JSON.stringify(body),
        });
        const text = await res.text();
        if (!res.ok) throw new Error(text || `SharePoint update ${res.status}`);
        return true;
      }, [itemUrl, digestUrl, changes])
    );
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

  _listBaseUrl() {
    return this.config.listUrl.replace(/\/items\/?(\?.*)?$/i, "");
  }

  _fieldsUrl(rich) {
    const select = rich
      ? "Title,InternalName,StaticName,EntityPropertyName,Hidden,TypeAsString,Required,ReadOnlyField"
      : "Title,InternalName,StaticName,Hidden,TypeAsString";
    return `${this._listBaseUrl()}/fields?$select=${select}&$format=json`;
  }

  _itemUrl(itemId) {
    return `${this.config.listUrl.replace(/\/$/, "")}(${encodeURIComponent(itemId)})`;
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

  // Public accessor for reading a normalized field value from a fetched item.
  readValue(item, internalName) {
    return this._readField(item, internalName);
  }

  // Turns SharePoint ISO date/datetime strings into a readable local date for display.
  static formatValue(value) {
    if (typeof value !== "string") return value;
    if (!/^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2}(\.\d+)?(Z|[+-]\d{2}:\d{2})?)?$/.test(value)) return value;
    const ms = Date.parse(value);
    return Number.isNaN(ms) ? value : new Date(ms).toLocaleDateString("he-IL");
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
