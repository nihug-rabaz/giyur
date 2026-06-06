class BackendClient {
  constructor(baseUrl) {
    this.baseUrl = (baseUrl || "http://localhost:3000").replace(/\/$/, "");
  }

  // Reads the site session cookies (incl. HttpOnly) and asks the backend to create an item
  // on the given site/list, so the backend targets whatever the extension is configured for.
  async createItem(siteUrl, listName, fields) {
    const cookie = await this._readCookie(siteUrl);
    let res;
    try {
      res = await fetch(`${this.baseUrl}/create-item`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ siteUrl, listName, fields, cookie }),
      });
    } catch {
      throw new Error("ה-backend לא זמין — ודא ש-'npm start' רץ בתיקיית backend");
    }
    const json = await res.json().catch(() => ({}));
    if (!res.ok || !json.success) throw new Error(json.error || `שגיאת שרת (${res.status})`);
    return json.data || {};
  }

  // Collects every cookie scoped to the site into a single Cookie header string.
  _readCookie(siteUrl) {
    return new Promise((resolve, reject) => {
      if (!chrome.cookies?.getAll) return reject(new Error("הרשאת cookies חסרה בתוסף"));
      chrome.cookies.getAll({ url: siteUrl }, (cookies) => {
        const err = chrome.runtime.lastError;
        if (err) return reject(new Error(err.message));
        const header = (cookies || []).map((c) => `${c.name}=${c.value}`).join("; ");
        if (!header) return reject(new Error("לא נמצאו Cookies — התחבר ל-SharePoint בדפדפן ונסה שוב"));
        resolve(header);
      });
    });
  }
}

window.BackendClient = BackendClient;
