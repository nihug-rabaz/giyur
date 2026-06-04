(function () {
if (window.SummonsTableReader) return;

class SummonsTableReader {
  constructor(dialog) {
    this.dialog = dialog;
  }

  read() {
    const headers = this._getHeaders();
    const rows = this._getRows(headers);
    return { headers, rows };
  }

  _clean(text) {
    return (text || "").replace(/\s+/g, " ").trim();
  }

  _isPlaceholder(text) {
    const t = (text || "").toLowerCase().trim();
    return (
      !t ||
      t === "-" ||
      t.includes("select") ||
      t.includes("בחר") ||
      t.includes("choose") ||
      t.includes("search") ||
      t.includes("type") ||
      t.includes("הקלד") ||
      t.includes("לחץ") ||
      t.includes("click") ||
      t.includes("אין נתונים")
    );
  }

  _getHeaders() {
    return Array.from(this.dialog.querySelectorAll(".ag-header-cell[col-id]"))
      .map((h) => ({
        id: h.getAttribute("col-id"),
        label: this._clean(h.innerText),
      }))
      .filter((h) => h.id && h.id !== "status")
      .reverse();
  }

  _extractCellValue(cell) {
    if (!cell) return "";

    const checkbox = cell.querySelector('input[type="checkbox"]');
    if (checkbox) return checkbox.checked ? "כן" : "לא";

    const input = cell.querySelector('input:not([type="checkbox"]), textarea');
    if (input) {
      const val = input.value?.trim();
      if (val) return this._clean(val);
      return "";
    }

    const nodes = cell.querySelectorAll("[aria-label]");
    for (const n of nodes) {
      const v = n.getAttribute("aria-label");
      if (v && !this._isPlaceholder(v)) return this._clean(v);
    }

    const chips = cell.querySelectorAll(".MuiChip-root");
    if (chips.length) {
      const val = [...chips].map((c) => c.innerText).join(", ");
      if (!this._isPlaceholder(val)) return this._clean(val);
    }

    const btn = cell.querySelector('[role="button"]');
    if (btn?.innerText && !this._isPlaceholder(btn.innerText)) {
      return this._clean(btn.innerText);
    }

    const text = this._clean(cell.innerText || cell.textContent);
    if (!this._isPlaceholder(text)) return text;

    return "";
  }

  _normalizeValue(v) {
    const val = String(v).trim().toLowerCase();
    if (["true", "on", "1", "yes", "כן"].includes(val)) return "כן";
    if (["false", "off", "0", "no", "לא"].includes(val)) return "לא";
    return v;
  }

  _getRows(headers) {
    const rows = [];
    const seen = new Set();

    const rowEls = Array.from(
      this.dialog.querySelectorAll(".ag-center-cols-container .ag-row")
    ).filter((row) => {
      const id = row.getAttribute("row-id") || row.getAttribute("aria-rowindex");
      if (!id || seen.has(id)) return false;
      seen.add(id);
      return true;
    });

    rowEls.forEach((row) => {
      const obj = {};
      headers.forEach((h) => {
        const cell = row.querySelector(`[col-id="${h.id}"]`);
        obj[h.id] = this._normalizeValue(this._extractCellValue(cell));
      });
      rows.push(obj);
    });

    return rows;
  }
}

window.SummonsTableReader = SummonsTableReader;
})();
