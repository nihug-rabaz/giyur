// Renders a set of QR codes (one per chunk) in a modal, with the total count shown clearly.
// Relies on the bundled qrcode-generator (window.qrcode) and QrChunker.
class QrExportModal {
  static CELL_SIZE = 4;
  static MARGIN = 10;
  static TYPE_FALLBACKS = [0, 15, 20, 25, 30, 35, 40];

  constructor() {
    this.overlay = null;
  }

  // Splits the text into chunks and shows one QR per chunk.
  open(text, chunkSize) {
    const payloads = QrChunker.encode(text, chunkSize);
    this._ensure();
    this.countEl.textContent = `נוצרו ${payloads.length} קודים — בדף הטלפון הזן ${payloads.length} וסרוק את כולם.`;
    this.gridEl.innerHTML = "";
    payloads.forEach((payload, i) => this.gridEl.appendChild(this._qrCard(payload, i + 1, payloads.length)));
    this.overlay.style.display = "flex";
  }

  _qrCard(payload, index, total) {
    const card = document.createElement("div");
    card.className = "qr-card";
    const img = document.createElement("img");
    img.alt = `QR ${index}/${total}`;
    img.src = this._dataUrl(payload);
    const label = document.createElement("span");
    label.className = "qr-card-label";
    label.textContent = `${index} / ${total}`;
    card.append(img, label);
    return card;
  }

  _dataUrl(data) {
    for (const type of QrExportModal.TYPE_FALLBACKS) {
      try {
        const qr = qrcode(type, "M");
        qr.addData(data);
        qr.make();
        return qr.createDataURL(QrExportModal.CELL_SIZE, QrExportModal.MARGIN);
      } catch (err) {
        // Data too large for this version — try a bigger one.
      }
    }
    throw new Error("הנתונים גדולים מדי ל-QR — הקטן את גודל הקטע");
  }

  _ensure() {
    if (this.overlay) return;
    const overlay = document.createElement("div");
    overlay.className = "modal-overlay";
    overlay.style.display = "none";

    const box = document.createElement("div");
    box.className = "modal qr-modal";

    const head = document.createElement("div");
    head.className = "modal-header";
    const title = document.createElement("h3");
    title.className = "modal-title";
    title.textContent = "ייצוא הגדרות לקודי QR";
    const close = document.createElement("button");
    close.type = "button";
    close.className = "modal-close";
    close.textContent = "\u2715";
    close.addEventListener("click", () => this._close());
    head.append(title, close);

    this.countEl = document.createElement("div");
    this.countEl.className = "modal-status info";
    this.countEl.style.padding = "12px 20px 0";

    const body = document.createElement("div");
    body.className = "modal-body";
    this.gridEl = document.createElement("div");
    this.gridEl.className = "qr-grid";
    body.appendChild(this.gridEl);

    box.append(head, this.countEl, body);
    overlay.appendChild(box);
    overlay.addEventListener("click", (e) => { if (e.target === overlay) this._close(); });
    document.addEventListener("keydown", (e) => { if (e.key === "Escape") this._close(); });
    document.body.appendChild(overlay);
    this.overlay = overlay;
  }

  _close() {
    if (this.overlay) this.overlay.style.display = "none";
  }
}

window.QrExportModal = QrExportModal;
