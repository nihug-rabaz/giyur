class QuickPrintPage {
  constructor() {
    this.status = document.getElementById("status");
    this._run();
  }

  async _run() {
    try {
      const payload = await this._loadPayload();
      if (!payload.headers.length || !payload.rows.length) {
        this._setStatus("לא נמצאו נתוני טבלה להדפסה");
        return;
      }

      const enriched = await this._enrich(payload.headers, payload.rows);
      if (this._noteIsProblem) alert(`שליפת SharePoint:\n${this._note}`);
      const config = await QuickPrintConfigStore.get();
      const templateOverride = new URLSearchParams(location.search).get("template");
      if (templateOverride) config.templatePath = templateOverride;
      const service = new QuickPrintService(config, enriched.headers);
      await service.printAll(enriched.rows);
      this._setStatus(`${this._note ? this._note + " · " : ""}חלון ההדפסה נפתח. אפשר לסגור את הלשונית לאחר ההדפסה.`);
    } catch (err) {
      console.error(err);
      this._setStatus(err?.message || "שגיאה בהכנת ההדפסה");
    }
  }

  async _loadPayload() {
    const data = await chrome.storage.local.get("summonsTableData");
    return {
      headers: data.summonsTableData?.headers || [],
      rows: data.summonsTableData?.rows || [],
    };
  }

  // Enriches rows from SharePoint and records a human-readable note explaining the outcome.
  async _enrich(headers, rows) {
    this._noteIsProblem = true;
    if (!window.SharePointLookupService || !window.SharePointConfigStore) {
      this._note = "שליפת SharePoint לא זמינה";
      return { headers, rows };
    }
    try {
      const config = await SharePointConfigStore.get();
      const fieldMap = config?.fieldMap || {};
      const missing = this._missingConfig(config, fieldMap);
      if (missing) {
        this._note = `דילוג על SharePoint — ${missing}`;
        console.warn("SharePoint enrich skipped:", missing, config);
        return { headers, rows };
      }

      this._setStatus("פותח SharePoint לשליפת נתונים...");
      const result = await new SharePointLookupService(config, fieldMap).enrich(headers, rows);
      this._noteIsProblem = !result.applied;
      this._note = result.applied
        ? "נתוני SharePoint נוספו"
        : "לא נמצאה התאמה ל-SharePoint — בדוק את עמודת הקישור והערכים";
      return { headers: result.headers, rows: result.rows };
    } catch (err) {
      console.error("SharePoint enrich failed", err);
      this._note = `שאיבת SharePoint נכשלה — ${err?.message || "בדוק הגדרות והרשאות"}`;
      return { headers, rows };
    }
  }

  _missingConfig(config, fieldMap) {
    if (!config.listTitle && !config.listUrl) return "לא הוגדרה רשימה";
    if (!config.join?.listFieldInternal) return "לא נבחר שדה התאמה ב-SharePoint";
    if (!config.join?.tableColumnDisplay) return "לא הוגדרה עמודת קישור בטבלה";
    if (!Object.keys(fieldMap).length) return "לא הוגדר שיוך שדות לתגים";
    return "";
  }

  _setStatus(text) {
    this.status.textContent = text;
  }
}

new QuickPrintPage();
