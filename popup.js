class StatusBar {
  constructor(el) { this.el = el; }

  set(text, type = "") {
    this.el.textContent = text;
    this.el.className = type;
  }
}

class TabService {
  static async getActive() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    return tab;
  }

  static async exec(tabId, func, args = []) {
    const results = await chrome.scripting.executeScript({ target: { tabId }, func, args });
    return results?.[0]?.result;
  }

  static waitForComplete(tabId, timeoutMs = 15000) {
    return new Promise((resolve) => {
      let settled = false;
      const finish = () => {
        if (settled) return;
        settled = true;
        chrome.tabs.onUpdated.removeListener(listener);
        clearTimeout(timer);
        resolve();
      };
      const listener = (id, info) => { if (id === tabId && info.status === "complete") finish(); };
      chrome.tabs.onUpdated.addListener(listener);
      const timer = setTimeout(finish, timeoutMs);
    });
  }
}

class ActionButton {
  constructor({ id, status, busyText }) {
    this.btn = document.getElementById(id);
    this.status = status;
    this.busyText = busyText;
    this.btn?.addEventListener("click", () => this._run());
  }

  async _run() {
    if (!this.btn) return;
    this.btn.disabled = true;
    this.btn.classList.add("is-loading");
    this.status.set(this.busyText, "info");
    try {
      await this.handle();
    } catch (err) {
      console.error(err);
      this.status.set("שגיאה לא צפויה — נסה שוב", "error");
    } finally {
      this.btn.disabled = false;
      this.btn.classList.remove("is-loading");
    }
  }

  async handle() { throw new Error("handle() must be implemented"); }
}

class CreateSummonsButton extends ActionButton {
  async handle() {
    const tab = await TabService.getActive();
    if (!tab?.id) return this.status.set("לא נמצאה לשונית פעילה", "error");

    const res = await TabService.exec(tab.id, () => {
      const btn = document.querySelector('button[type="submit"]');
      if (!btn) return { ok: false, error: "לא נמצא כפתור שליחה בדף" };
      btn.click();
      return { ok: true };
    });

    if (res?.ok) this.status.set("המזומנים נוצרו בהצלחה", "ok");
    else this.status.set(res?.error || "שגיאה ביצירת המזומנים", "error");
  }
}

class OpenSummonsPageButton extends ActionButton {
  static TARGET_URL = "https://rabaz.army.idf/molecule/client.html#/org/molecule/app/app__max-app/pages/51";
  static BUTTON_SELECTOR = 'button[aria-label="יצירה מרובה"]';

  async handle() {
    const tab = await TabService.getActive();
    if (!tab?.id) return this.status.set("לא נמצאה לשונית פעילה", "error");

    if (tab.url !== OpenSummonsPageButton.TARGET_URL) {
      this.status.set("טוען את דף הזימונים...", "info");
      await chrome.tabs.update(tab.id, { url: OpenSummonsPageButton.TARGET_URL });
      await TabService.waitForComplete(tab.id);
    }

    this.status.set("לוחץ על יצירה מרובה...", "info");
    const res = await this._clickWhenReady(tab.id);
    if (res?.ok) this.status.set("נפתח חלון יצירה מרובה", "ok");
    else this.status.set(res?.error || 'לא נמצא הכפתור "יצירה מרובה"', "error");
  }

  _clickWhenReady(tabId) {
    return TabService.exec(tabId, async (selector) => {
      const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
      const start = Date.now();
      while (Date.now() - start < 15000) {
        const btn = document.querySelector(selector);
        if (btn) { btn.click(); return { ok: true }; }
        await sleep(200);
      }
      return { ok: false, error: 'לא נמצא הכפתור "יצירה מרובה" בדף' };
    }, [OpenSummonsPageButton.BUTTON_SELECTOR]);
  }
}

class DownloadReportButton extends ActionButton {
  async handle() {
    const tab = await TabService.getActive();
    if (!tab?.id) return this.status.set("לא נמצאה לשונית פעילה", "error");
    try {
      const response = await chrome.tabs.sendMessage(tab.id, { action: "exportWord" });
      if (response?.ok) this.status.set("ההורדה התחילה", "ok");
      else this.status.set(response?.error || "שגיאה בייצוא", "error");
    } catch {
      this.status.set("פתח את דף הדוח ונסה שוב", "error");
    }
  }
}

class AdminUnlock {
  static CLICKS_REQUIRED = 10;
  static SECRET = "שלום";
  static CLICK_RESET_MS = 3000;
  static TYPE_WINDOW_MS = 8000;

  constructor(logoId) {
    this.logo = document.getElementById(logoId);
    this.clicks = 0;
    this.lastClick = 0;
    this.armed = false;
    this.armedUntil = 0;
    this.buffer = "";
    this._bind();
  }

  _bind() {
    if (!this.logo) return;
    this.logo.addEventListener("click", () => this._onClick());
    document.addEventListener("keydown", (e) => this._onKey(e));
  }

  // Counts rapid logo clicks; reaching the threshold arms the secret-word listener.
  _onClick() {
    const now = Date.now();
    this.clicks = now - this.lastClick > AdminUnlock.CLICK_RESET_MS ? 1 : this.clicks + 1;
    this.lastClick = now;
    if (this.clicks >= AdminUnlock.CLICKS_REQUIRED) {
      this.armed = true;
      this.armedUntil = now + AdminUnlock.TYPE_WINDOW_MS;
      this.buffer = "";
      this.clicks = 0;
    }
  }

  // While armed, opens the admin window once the secret word is typed.
  _onKey(e) {
    if (!this.armed) return;
    if (Date.now() > this.armedUntil) return this._reset();
    if (e.key && e.key.length === 1) this.buffer += e.key;
    if (this.buffer.includes(AdminUnlock.SECRET)) {
      this._reset();
      chrome.tabs.create({ url: chrome.runtime.getURL("admin.html") });
    }
  }

  _reset() {
    this.armed = false;
    this.buffer = "";
  }
}

class AdminGate {
  static PASSWORD = "rabaz123";

  constructor({ toggleId, gateId, inputId, enterId, status }) {
    this.status = status;
    this.gate = document.getElementById(gateId);
    this.input = document.getElementById(inputId);
    document.getElementById(toggleId)?.addEventListener("click", () => this._toggle());
    document.getElementById(enterId)?.addEventListener("click", () => this._submit());
    this.input?.addEventListener("keydown", (e) => { if (e.key === "Enter") this._submit(); });
  }

  // Reveals the inline password field and focuses it.
  _toggle() {
    if (!this.gate) return;
    this.gate.hidden = !this.gate.hidden;
    if (!this.gate.hidden) this.input?.focus();
  }

  // Opens the admin window only when the entered password matches.
  _submit() {
    if (this.input?.value === AdminGate.PASSWORD) {
      chrome.tabs.create({ url: chrome.runtime.getURL("admin.html") });
      this.status.set("נפתח מסך הגדרות", "ok");
      this.gate.hidden = true;
    } else {
      this.status.set("סיסמה שגויה", "error");
    }
    if (this.input) this.input.value = "";
  }
}

class OpenPageButton {
  constructor(id, page) {
    document.getElementById(id)?.addEventListener("click", () =>
      chrome.tabs.create({ url: chrome.runtime.getURL(page) })
    );
  }
}

const status = new StatusBar(document.getElementById("status"));
new CreateSummonsButton({ id: "createSummonsBtn", status, busyText: "יוצר מזומנים..." });
new OpenSummonsPageButton({ id: "openSummonsPageBtn", status, busyText: "פותח דף זימונים..." });
new DownloadReportButton({ id: "downloadBtn", status, busyText: "מייצא..." });
new AdminUnlock("appLogo");
new AdminGate({ toggleId: "openAdminBtn", gateId: "adminGate", inputId: "adminPass", enterId: "adminEnter", status });
new OpenPageButton("printHubBtn", "hub.html");
new OpenPageButton("spLookupBtn", "sp-lookup.html");
