chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg.action === "exportWord") {
    stageSharePointConfig()
      .then(() => injectPageScript("export-page.js"))
      .then(() => sendResponse({ ok: true }))
      .catch(() => sendResponse({ ok: false, error: "לא ניתן לטעון את סקריפט הייצוא" }));
    return true;
  }

  if (msg.action === "scrapeTable") {
    scrapeSummonsTable(sendResponse);
    return true;
  }
});

function injectPageScript(file) {
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = chrome.runtime.getURL(file);
    script.onload = () => {
      script.remove();
      resolve();
    };
    script.onerror = reject;
    (document.head || document.documentElement).appendChild(script);
  });
}

// Hands the saved SharePoint config to the page-context export script via a shared DOM attribute.
async function stageSharePointConfig() {
  try {
    const config = window.SharePointConfigStore
      ? await SharePointConfigStore.get()
      : window.SHAREPOINT_LOOKUP;
    const payload = { lookup: config, fieldMap: config?.fieldMap || window.SHAREPOINT_FIELD_MAP };
    document.documentElement.setAttribute("data-sp-config", JSON.stringify(payload));
  } catch {
    document.documentElement.removeAttribute("data-sp-config");
  }
}

// Returns the raw scraped table; SharePoint enrichment runs later in extension pages
// (quick-print / single-export), which can reach SharePoint via its authenticated tab.
async function scrapeSummonsTable(sendResponse) {
  try {
    const dialog = document.querySelector('[role="dialog"]');
    if (!dialog) {
      sendResponse({ ok: false, error: "לא נמצא דיאלוג עם טבלת זימונים" });
      return;
    }
    const data = new window.SummonsTableReader(dialog).read();
    sendResponse({ ok: true, data });
  } catch (err) {
    sendResponse({ ok: false, error: err?.message || "שגיאה בקריאת הטבלה" });
  }
}
