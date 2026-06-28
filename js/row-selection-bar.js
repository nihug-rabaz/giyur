class RowSelectionBar {
  constructor({ onTemplate, onClear }) {
    this.onTemplate = onTemplate;
    this.onClear = onClear;
    this._build();
  }

  update({ count, templates }) {
    if (!count) {
      this.root.hidden = true;
      return;
    }
    this.countEl.textContent = String(count);
    this.templatesEl.innerHTML = "";
    const fragment = document.createDocumentFragment();
    (templates || []).forEach((template, index) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "btn btn--accent btn--sm";
      btn.textContent = template.name || `תבנית ${index + 1}`;
      btn.addEventListener("click", (e) => this.onTemplate?.(template, e));
      fragment.appendChild(btn);
    });
    this.templatesEl.appendChild(fragment);
    this.root.hidden = false;
  }

  hide() {
    this.root.hidden = true;
  }

  _build() {
    this.root = document.createElement("div");
    this.root.className = "selection-bar";
    this.root.hidden = true;

    const info = document.createElement("div");
    info.className = "selection-bar__info";
    this.countEl = document.createElement("span");
    this.countEl.className = "selection-bar__count";
    const label = document.createElement("span");
    label.className = "selection-bar__label";
    label.textContent = "נבחרו";
    info.append(this.countEl, label);

    this.templatesEl = document.createElement("div");
    this.templatesEl.className = "selection-bar__templates";

    const clearBtn = document.createElement("button");
    clearBtn.type = "button";
    clearBtn.className = "btn btn--sm selection-bar__clear";
    clearBtn.textContent = "בטל בחירה";
    clearBtn.addEventListener("click", () => this.onClear?.());

    this.root.append(info, this.templatesEl, clearBtn);
    document.body.appendChild(this.root);
  }
}

window.RowSelectionBar = RowSelectionBar;
