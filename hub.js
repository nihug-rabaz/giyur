// Shell that aggregates the existing summons + sessions screens (loaded as-is in iframes).
// A hidden top bar slides open on demand; frames load lazily on first activation.
class PrintHub {
  constructor() {
    this.body = document.body;
    this.toggle = document.getElementById("hubToggle");
    this.currentEl = document.getElementById("hubCurrent");
    this.scrim = document.getElementById("hubScrim");
    this.items = [...document.querySelectorAll(".hub-nav-item")];

    this.toggle.addEventListener("click", () => this._setOpen(!this._isOpen()));
    this.scrim.addEventListener("click", () => this._setOpen(false));
    document.addEventListener("keydown", (e) => { if (e.key === "Escape") this._setOpen(false); });
    this.items.forEach((item) => item.addEventListener("click", () => this._activate(item)));

    const first = this.items.find((i) => i.classList.contains("is-active")) || this.items[0];
    if (first) this._activate(first, false);
  }

  _isOpen() { return this.body.classList.contains("hub-open"); }

  _setOpen(open) {
    this.body.classList.toggle("hub-open", open);
    this.toggle.setAttribute("aria-expanded", String(open));
  }

  _activate(item, close = true) {
    const target = item.dataset.target;
    const frame = document.getElementById(`frame-${target}`);
    if (frame && !frame.src) frame.src = item.dataset.src;

    this.items.forEach((i) => i.classList.toggle("is-active", i === item));
    document.querySelectorAll(".hub-frame").forEach((f) =>
      f.classList.toggle("is-active", f === frame)
    );
    this.currentEl.textContent = item.dataset.name || "";
    if (close) this._setOpen(false);
  }
}

new PrintHub();
