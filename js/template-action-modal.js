class TemplateActionModal {
  static PRINT_ITEMS = [
    { label: "הדפסה", value: "browserPrint", variant: "" },
    { label: "ייצוא ל-Word", value: "download", variant: "export" },
  ];

  constructor() {
    this.onChoose = null;
    this._build();
  }

  open({ title, onChoose, hint, anchor, items }) {
    this.titleEl.textContent = title || "תבנית";
    this.hintEl.textContent = hint || "";
    this.hintEl.hidden = !hint;
    this.onChoose = onChoose;
    this._renderItems(items || TemplateActionModal.PRINT_ITEMS);
    this.overlay.style.display = "block";
    this.menu.hidden = false;
    requestAnimationFrame(() => this._position(anchor));
  }

  _renderItems(items) {
    this.actionsEl.innerHTML = "";
    items.forEach((item) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "template-action-menu__item"
        + (item.variant === "export" ? " template-action-menu__item--export" : "");
      btn.textContent = item.label;
      btn.addEventListener("click", () => this._choose(item.value));
      this.actionsEl.appendChild(btn);
    });
  }

  close() {
    this.overlay.style.display = "none";
    this.menu.hidden = true;
    this.onChoose = null;
  }

  _choose(mode) {
    const handler = this.onChoose;
    this.close();
    handler?.(mode);
  }

  _position(anchor) {
    const margin = 8;
    const gap = 6;
    let x;
    let y;
    if (anchor instanceof MouseEvent) {
      x = anchor.clientX;
      y = anchor.clientY + gap;
    } else if (anchor instanceof Element) {
      const rect = anchor.getBoundingClientRect();
      x = rect.left;
      y = rect.bottom + gap;
    } else if (anchor?.x != null && anchor?.y != null) {
      x = anchor.x;
      y = anchor.y + gap;
    } else {
      x = window.innerWidth / 2 - 90;
      y = window.innerHeight / 2 - 60;
    }

    this.menu.style.visibility = "hidden";
    const width = this.menu.offsetWidth;
    const height = this.menu.offsetHeight;
    const rtl = document.documentElement.dir === "rtl";
    let left = rtl ? x - width : x;
    let top = y;

    if (top + height > window.innerHeight - margin) {
      const above = (anchor instanceof MouseEvent ? anchor.clientY : y) - height - gap * 2;
      if (above >= margin) top = above;
    }
    left = Math.max(margin, Math.min(left, window.innerWidth - width - margin));
    top = Math.max(margin, Math.min(top, window.innerHeight - height - margin));

    this.menu.style.left = `${left}px`;
    this.menu.style.top = `${top}px`;
    this.menu.style.visibility = "visible";
  }

  _build() {
    this.overlay = document.createElement("div");
    this.overlay.className = "template-action-overlay";
    this.overlay.style.display = "none";
    this.overlay.addEventListener("click", (e) => { if (e.target === this.overlay) this.close(); });

    this.menu = document.createElement("div");
    this.menu.className = "template-action-menu";
    this.menu.hidden = true;
    this.menu.addEventListener("click", (e) => e.stopPropagation());

    this.titleEl = document.createElement("div");
    this.titleEl.className = "template-action-menu__title";

    this.hintEl = document.createElement("div");
    this.hintEl.className = "template-action-menu__hint";

    const head = document.createElement("div");
    head.className = "template-action-menu__head";
    head.append(this.titleEl, this.hintEl);

    this.actionsEl = document.createElement("div");
    this.actionsEl.className = "template-action-menu__actions";

    this.menu.append(head, this.actionsEl);
    this.overlay.appendChild(this.menu);
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && this.overlay.style.display === "block") this.close();
    });
    document.body.appendChild(this.overlay);
  }
}

window.TemplateActionModal = TemplateActionModal;
