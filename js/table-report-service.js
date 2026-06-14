class TableReportService {
  constructor(options = {}) {
    this.title = options.title || "דוח";
    this.fileName = options.fileName || "דוח.docx";
    this.logoUrl = options.logoUrl || "https://times.rabaz.co.il/logo1.png";
  }

  // Builds a landscape Word report from one or more groups (each = a titled table) and downloads it.
  // `groups` is [{ title, rows }]; passing a plain rows array renders a single untitled table.
  async download(headers, groups) {
    const blob = await this.build(headers, groups);
    FileDownloadHelper.download(blob, this.fileName);
  }

  async build(headers, groups) {
    const d = window.docxBuilder || window.docx;
    if (!d) throw new Error("ספריית יצירת ה-Word לא נטענה");
    const sections = Array.isArray(groups) && groups[0]?.rows ? groups : [{ title: null, rows: groups || [] }];
    const children = [];
    const logo = await this._logo(d);
    if (logo) children.push(new d.Paragraph({ alignment: d.AlignmentType.LEFT, children: [logo] }));
    children.push(new d.Paragraph({
      alignment: d.AlignmentType.CENTER,
      children: [new d.TextRun({ text: this.title, bold: true, size: 32 })],
    }));
    sections.forEach((group, i) => {
      if (i > 0) children.push(new d.Paragraph({ children: [] }));
      if (group.title) {
        children.push(new d.Paragraph({
          alignment: d.AlignmentType.CENTER,
          bidirectional: true,
          children: [new d.TextRun({ text: group.title, bold: true, size: 26 })],
        }));
      }
      children.push(this._table(d, headers, group.rows));
    });

    const doc = new d.Document({
      sections: [{
        properties: { page: { size: { orientation: d.PageOrientation.LANDSCAPE }, bidi: true } },
        children,
      }],
    });
    return d.Packer.toBlob(doc);
  }

  _table(d, headers, rows) {
    return new d.Table({
      width: { size: 100, type: d.WidthType.PERCENTAGE },
      visuallyRightToLeft: true,
      rows: [
        this._row(d, headers.map((h) => h.label), { header: true }),
        ...rows.map((r) => this._row(d, headers.map((h) => String(r[h.id] ?? "")))),
      ],
    });
  }

  _row(d, values, { header = false } = {}) {
    return new d.TableRow({
      children: values.map((text) =>
        new d.TableCell({
          verticalAlign: d.VerticalAlign.CENTER,
          shading: header ? { fill: "D9D9D9" } : undefined,
          children: [new d.Paragraph({
            alignment: d.AlignmentType.CENTER,
            bidirectional: true,
            children: [new d.TextRun({ text, bold: header })],
          })],
        })
      ),
    });
  }

  async _logo(d) {
    try {
      const buffer = await (await fetch(this.logoUrl)).arrayBuffer();
      return new d.ImageRun({ data: new Uint8Array(buffer), transformation: { width: 94, height: 94 } });
    } catch {
      return null;
    }
  }
}

window.TableReportService = TableReportService;
