(async function () {
  const scriptUrl = document.currentScript?.src || "";

  function readStagedSharePointConfig() {
    try {
      const raw = document.documentElement.getAttribute("data-sp-config");
      return raw ? JSON.parse(raw) : null;
    } catch {
      return null;
    }
  }

  if (!window.docx) {
    await new Promise((res, rej) => {
      const s = document.createElement("script");
      s.src = scriptUrl.replace(/export-page\.js(\?.*)?$/, "index.umd.js");
      s.onload = res;
      s.onerror = rej;
      document.head.appendChild(s);
    });
  }

  const base = scriptUrl.replace(/export-page\.js(\?.*)?$/, "");
  const deps = ["js/table-reader.js", "js/sharepoint-field-map.js", "js/sharepoint-lookup-service.js"];
  for (const file of deps) {
    const globalKey = file.includes("table-reader")
      ? "SummonsTableReader"
      : file.includes("lookup-service")
        ? "SharePointLookupService"
        : "SHAREPOINT_LOOKUP";
    if (globalKey === "SHAREPOINT_LOOKUP" ? window.SHAREPOINT_LOOKUP : window[globalKey]) continue;
    await new Promise((res, rej) => {
      const s = document.createElement("script");
      s.src = base + file;
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

  let { headers, rows } = new window.SummonsTableReader(dialog).read();
  const staged = readStagedSharePointConfig();
  const lookup = staged?.lookup || window.SHAREPOINT_LOOKUP;
  const fieldMap = staged?.fieldMap || window.SHAREPOINT_FIELD_MAP;
  if (window.SharePointLookupService && lookup) {
    try {
      const enriched = await new SharePointLookupService(lookup, fieldMap).enrich(headers, rows);
      headers = enriched.headers;
      rows = enriched.rows;
    } catch (err) {
      console.warn("SharePoint enrich failed", err);
    }
  }

  const logoBlob = await fetch("https://times.rabaz.co.il/logo1.png").then(r => r.blob());
  const logoBuffer = await logoBlob.arrayBuffer();

  const logo = new ImageRun({
    data: new Uint8Array(logoBuffer),
    transformation: { width: 94, height: 94 }
  });

  const table = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: headers.map(h =>
          new TableCell({
            verticalAlign: VerticalAlign.CENTER,
            shading: { fill: "D9D9D9" },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                bidirectional: true,
                children: [new TextRun({ text: h.label, bold: true })]
              })
            ]
          })
        )
      }),
      ...rows.map((r) =>
        new TableRow({
          children: headers.map(h =>
            new TableCell({
              verticalAlign: VerticalAlign.CENTER,
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  bidirectional: true,
                  children: [new TextRun(String(r[h.id] ?? ""))]
                })
              ]
            })
          )
        })
      )
    ]
  });

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

})();
