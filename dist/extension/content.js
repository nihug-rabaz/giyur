chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg.action !== "exportWord") return;

  const script = document.createElement("script");
  script.src = chrome.runtime.getURL("export-page.js");
  script.onload = () => script.remove();
  script.onerror = () => {
    sendResponse({ ok: false, error: "לא ניתן לטעון את סקריפט הייצוא" });
  };

  (document.head || document.documentElement).appendChild(script);
  sendResponse({ ok: true });
});
