(async function () {

  if (!window.docx) {
    await new Promise((res, rej) => {
      const s = document.createElement("script");
      s.src = document.currentScript.src.replace(/export-page\.js(\?.*)?$/, "index.umd.js");
      s.onload = res;
      s.onerror = rej;
      document.head.appendChild(s);
    });
  }

  const {
    Document,
    Packer,
    Paragraph,
    TextRun,
    Table,
    TableRow,
    TableCell,
    WidthType,
    AlignmentType,
    PageOrientation,
    ImageRun,
    VerticalAlign
  } = window.docx;

  const dialog = document.querySelector('[role="dialog"]');
  if (!dialog) return alert("לא נמצא dialog");

  const clean = (t) =>
    (t || "").replace(/\s+/g, " ").trim();

  // ================= PLACEHOLDER FILTER =================
  function isPlaceholder(text) {
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

  // ================= HEADERS =================
  function getHeaders() {
    return Array.from(dialog.querySelectorAll('.ag-header-cell[col-id]'))
      .map(h => ({
        id: h.getAttribute('col-id'),
        label: clean(h.innerText)
      }))
      .filter(h => h.id && h.id !== "status")
      .reverse();
  }

  // ================= CELL EXTRACT =================
  function extractCellValue(cell) {
    if (!cell) return "";

    // ===== SWITCH / CHECKBOX (הכי חשוב!) =====
    const checkbox = cell.querySelector('input[type="checkbox"]');
    if (checkbox) {
      return checkbox.checked ? "כן" : "לא";
    }

    // ===== INPUT / TEXTAREA =====
    const input = cell.querySelector('input:not([type="checkbox"]), textarea');
    if (input) {
      const val = input.value?.trim();
      if (val) return clean(val);
      return "";
    }

    // ===== MULTISELECT / ARIA =====
    const nodes = cell.querySelectorAll('[aria-label]');
    for (const n of nodes) {
      const v = n.getAttribute('aria-label');
      if (v && !isPlaceholder(v)) {
        return clean(v);
      }
    }

    // ===== CHIPS =====
    const chips = cell.querySelectorAll('.MuiChip-root');
    if (chips.length) {
      const val = [...chips].map(c => c.innerText).join(', ');
      if (!isPlaceholder(val)) return clean(val);
    }

    // ===== BUTTON =====
    const btn = cell.querySelector('[role="button"]');
    if (btn?.innerText && !isPlaceholder(btn.innerText)) {
      return clean(btn.innerText);
    }

    // ===== TEXT =====
    const text = clean(cell.innerText || cell.textContent);
    if (!isPlaceholder(text)) return text;

    return "";
  }

  // ================= BOOL NORMALIZE =================
  function normalizeValue(v) {
    const val = String(v).trim().toLowerCase();

    if (["true", "on", "1", "yes", "כן"].includes(val)) return "כן";
    if (["false", "off", "0", "no", "לא"].includes(val)) return "לא";

    return v;
  }

  // ================= ROWS =================
  function getRows(headers) {
    const rows = [];
    const seen = new Set();

    const rowEls = Array.from(
      dialog.querySelectorAll('.ag-center-cols-container .ag-row')
    ).filter(row => {
      const id = row.getAttribute('row-id') || row.getAttribute('aria-rowindex');
      if (!id || seen.has(id)) return false;
      seen.add(id);
      return true;
    });

    rowEls.forEach(row => {
      const obj = {};

      headers.forEach(h => {
        const cell = row.querySelector(`[col-id="${h.id}"]`);
        let val = extractCellValue(cell);
        obj[h.id] = normalizeValue(val);
      });

      rows.push(obj);
    });

    return rows;
  }

  const headers = getHeaders();
  const rows = getRows(headers);

  // ================= LOGO =================
  const logoBlob = await fetch("https://times.rabaz.co.il/logo1.png").then(r => r.blob());
  const logoBuffer = await logoBlob.arrayBuffer();

  const logo = new ImageRun({
    data: new Uint8Array(logoBuffer),
    transformation: { width: 94, height: 94 }
  });

  // ================= TABLE =================
  const table = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },

    rows: [

      // HEADER
      new TableRow({
        children: headers.map(h =>
          new TableCell({
            verticalAlign: VerticalAlign.CENTER,
            shading: { fill: "D9D9D9" },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                bidirectional: true,
                children: [
                  new TextRun({ text: h.label, bold: true })
                ]
              })
            ]
          })
        )
      }),

      // DATA
      ...rows.map((r) =>
        new TableRow({
          children: headers.map(h =>
            new TableCell({
              verticalAlign: VerticalAlign.CENTER,
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  bidirectional: true,
                  children: [
                    new TextRun(String(r[h.id] ?? ""))
                  ]
                })
              ]
            })
          )
        })
      )
    ]
  });

  // ================= DOC =================
  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { orientation: PageOrientation.LANDSCAPE },
          bidi: true
        }
      },
      children: [

        new Paragraph({
          alignment: AlignmentType.LEFT,
          children: [logo]
        }),

        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "דוח דימונים", bold: true, size: 32 })
          ]
        }),

        table
      ]
    }]
  });

  const blob = await Packer.toBlob(doc);

  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "דוח דימונים.docx";
  a.click();

  console.log("✔ Export completed");

})();
