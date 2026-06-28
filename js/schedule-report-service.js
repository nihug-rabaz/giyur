// Builds a formal "לוח מוזמנים" Word report: header, metadata blocks, zebra table, footer with page numbers.
class ScheduleReportService {
  constructor(options = {}) {
    this.fileName = options.fileName || "לוח-מוזמנים.docx";
    this.logoUrls = options.logoUrls || ["https://times.rabaz.co.il/logo1.png"];
  }

  async download(headers, groups) {
    const blob = await this.build(headers, groups);
    FileDownloadHelper.download(blob, this.fileName);
  }

  async build(headers, groups) {
    const d = window.docxBuilder || window.docx;
    if (!d) throw new Error("ספריית יצירת ה-Word לא נטענה");
    const sections = Array.isArray(groups) && groups.length
      ? groups.filter((group) => Array.isArray(group.rows))
      : [{ title: null, rows: groups || [], meta: {} }];
    const children = [];
    const logos = await this._loadLogos(d);

    sections.forEach((group, index) => {
      if (index > 0) children.push(this._pageBreak(d));
      const meta = { ...this._defaultMeta(group.rows), ...(group.meta || {}) };
      if (group.title && !meta.location) meta.location = group.title;
      children.push(...this._headerBlock(d, logos, meta));
      children.push(this._metaTable(d, meta));
      children.push(this._spacer(d));
      children.push(this._dataTable(d, headers, group.rows));
    });

    const printedAt = sections[0]?.meta?.printedAt || new Date().toLocaleString("he-IL");
    const doc = new d.Document({
      sections: [{
        properties: {
          page: {
            size: { orientation: d.PageOrientation.LANDSCAPE },
            margin: { top: 720, right: 720, bottom: 900, left: 720 },
          },
        },
        footers: { default: this._footer(d, printedAt) },
        children,
      }],
    });
    return d.Packer.toBlob(doc);
  }

  _defaultMeta(rows) {
    return {
      title: "לוח מוזמנים",
      judge: "",
      gregorianDate: "",
      hebrewDate: "",
      location: "",
      summonsType: "",
      summonedCount: rows?.length || 0,
    };
  }

  _headerBlock(d, logos, meta) {
    const judgeCell = this._cell(d, meta.judge ? `אב״ד: ${meta.judge}` : "", {
      align: d.AlignmentType.LEFT,
      bold: false,
      size: 22,
      borders: this._noBorders(d),
      width: { size: 25, type: d.WidthType.PERCENTAGE },
    });
    const logoCell = this._cell(d, "", {
      align: d.AlignmentType.CENTER,
      borders: this._noBorders(d),
      width: { size: 50, type: d.WidthType.PERCENTAGE },
      children: logos.length
        ? [new d.Paragraph({ alignment: d.AlignmentType.CENTER, children: logos })]
        : [this._para(d, "בית הדין הצבאי", { align: d.AlignmentType.CENTER, size: 20 })],
    });
    const spacerCell = this._cell(d, "", { borders: this._noBorders(d), width: { size: 25, type: d.WidthType.PERCENTAGE } });

    return [
      new d.Table({
        width: { size: 100, type: d.WidthType.PERCENTAGE },
        visuallyRightToLeft: true,
        borders: d.TableBorders.NONE,
        rows: [new d.TableRow({ children: [logoCell, spacerCell, judgeCell] })],
      }),
      this._para(d, meta.title || "לוח מוזמנים", { align: d.AlignmentType.CENTER, bold: true, size: 56, spacing: { after: 200 } }),
    ];
  }

  _metaTable(d, meta) {
    const block = (lines) => lines.map((t) => this._para(d, t, { align: d.AlignmentType.CENTER, size: 20, spacing: { after: 60 } }));
    const col1 = this._cell(d, "", { borders: this._lightBorders(d), children: block([
      meta.hebrewDate ? `תאריך עברי: ${meta.hebrewDate}` : "תאריך עברי:",
      meta.gregorianDate ? `תאריך לועזי: ${meta.gregorianDate}` : "תאריך לועזי:",
    ]) });
    const col2 = this._cell(d, "", { borders: this._lightBorders(d), children: block([
      meta.location ? `מקום הזימון: ${meta.location}` : "מקום הזימון:",
      meta.summonsType ? `סוג הזימון: ${meta.summonsType}` : "סוג הזימון:",
    ]) });
    const col3 = this._cell(d, "", { borders: this._lightBorders(d), children: block([
      `חיילים מוזמנים: ${meta.summonedCount ?? ""}`,
    ]) });

    return new d.Table({
      width: { size: 100, type: d.WidthType.PERCENTAGE },
      visuallyRightToLeft: true,
      borders: this._lightBorders(d),
      rows: [new d.TableRow({ children: [col1, col2, col3] })],
    });
  }

  _dataTable(d, headers, rows) {
    const borders = this._gridBorders(d);
    const headerRow = new d.TableRow({
      tableHeader: true,
      children: headers.map((h) => this._cell(d, h.label, {
        bold: true,
        size: 18,
        shading: { fill: "D9D9D9" },
        borders,
      })),
    });
    const bodyRows = (rows || []).map((row, i) =>
      new d.TableRow({
        children: headers.map((h) => this._cell(d, String(row[h.id] ?? ""), {
          size: 16,
          shading: { fill: i % 2 ? "F2F2F2" : "FFFFFF" },
          borders,
        })),
      })
    );

    return new d.Table({
      width: { size: 100, type: d.WidthType.PERCENTAGE },
      visuallyRightToLeft: true,
      borders,
      rows: [headerRow, ...bodyRows],
    });
  }

  _footer(d, printedAt) {
    const pagePara = new d.Paragraph({
      alignment: d.AlignmentType.LEFT,
      bidirectional: true,
      children: [
        new d.TextRun({ text: "עמוד ", size: 18 }),
        new d.TextRun({ children: [d.PageNumber.CURRENT], size: 18 }),
        new d.TextRun({ text: " מתוך ", size: 18 }),
        new d.TextRun({ children: [d.PageNumber.TOTAL_PAGES], size: 18 }),
      ],
    });
    const datePara = new d.Paragraph({
      alignment: d.AlignmentType.RIGHT,
      bidirectional: true,
      children: [new d.TextRun({ text: `תאריך ושעת הדפסה: ${printedAt}`, size: 18 })],
    });
    return new d.Footer({
      children: [
        new d.Table({
          width: { size: 100, type: d.WidthType.PERCENTAGE },
          visuallyRightToLeft: true,
          borders: d.TableBorders.NONE,
          rows: [new d.TableRow({
            children: [
              this._cell(d, "", { borders: this._noBorders(d), children: [pagePara] }),
              this._cell(d, "", { borders: this._noBorders(d), children: [datePara] }),
            ],
          })],
        }),
      ],
    });
  }

  _cell(d, text, opts = {}) {
    const children = opts.children || [this._para(d, text, opts)];
    return new d.TableCell({
      verticalAlign: d.VerticalAlign.CENTER,
      width: opts.width,
      shading: opts.shading,
      borders: opts.borders,
      margins: { top: 60, bottom: 60, left: 80, right: 80 },
      children,
    });
  }

  _para(d, text, { align, bold, size, spacing } = {}) {
    return new d.Paragraph({
      alignment: align || d.AlignmentType.CENTER,
      bidirectional: true,
      spacing,
      children: [new d.TextRun({ text: text || "", bold: !!bold, size: size || 20 })],
    });
  }

  _spacer(d) {
    return new d.Paragraph({ spacing: { after: 160 }, children: [] });
  }

  _pageBreak(d) {
    return new d.Paragraph({ children: [new d.PageBreak()] });
  }

  _gridBorders(d) {
    const side = { style: d.BorderStyle.SINGLE, size: 4, color: "A6A6A6" };
    return { top: side, bottom: side, left: side, right: side, insideHorizontal: side, insideVertical: side };
  }

  _lightBorders(d) {
    const side = { style: d.BorderStyle.SINGLE, size: 2, color: "CFCFCF" };
    return { top: side, bottom: side, left: side, right: side, insideHorizontal: side, insideVertical: side };
  }

  _noBorders(d) {
    return d.TableBorders.NONE;
  }

  async _loadLogos(d) {
    const runs = [];
    for (const url of this.logoUrls.slice(0, 3)) {
      try {
        const buffer = await (await fetch(url)).arrayBuffer();
        runs.push(new d.ImageRun({ data: new Uint8Array(buffer), transformation: { width: 72, height: 72 } }));
        runs.push(new d.TextRun({ text: "  " }));
      } catch {
        // Skip logos that fail to load.
      }
    }
    return runs;
  }
}

window.ScheduleReportService = ScheduleReportService;
