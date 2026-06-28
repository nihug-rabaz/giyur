class DocxTemplateEngine {
  static XML_PARTS = [
    "word/document.xml",
    "word/header1.xml",
    "word/header2.xml",
    "word/header3.xml",
    "word/footer1.xml",
    "word/footer2.xml",
    "word/footer3.xml",
  ];

  constructor(templateBuffer) {
    this.templateBuffer = templateBuffer;
  }

  async fill(rowData, headers, mapping = {}) {
    const zip = await JSZip.loadAsync(this.templateBuffer);
    const replacements = this._buildReplacements(rowData, headers, mapping);

    for (const key of DocxTemplateEngine.getXmlPartKeys(zip)) {
      let xml = await zip.file(key).async("string");
      xml = DocxTemplateEngine.collapsePlaceholderRuns(xml);
      xml = this._applyReplacements(xml, replacements);
      zip.file(key, xml);
    }

    return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  }

  static getXmlPartKeys(zip) {
    return Object.keys(zip.files).filter(
      (k) => !zip.files[k].dir && DocxTemplateEngine.XML_PARTS.includes(k.replace(/\\/g, "/"))
    );
  }

  static getZipPart(zip, partPath) {
    const normalized = partPath.replace(/\\/g, "/");
    const key = Object.keys(zip.files).find(
      (k) => k.replace(/\\/g, "/") === normalized && !zip.files[k].dir
    );
    return key ? zip.file(key) : null;
  }

  static collapsePlaceholderRuns(xml) {
    return xml.replace(/\[(?:<[^>]+>|[^\[\]<>])*?\]/g, (match) =>
      match.replace(/<[^>]+>/g, "")
    );
  }

  static async extractPlaceholders(buffer) {
    const zip = await JSZip.loadAsync(buffer);
    const found = new Set();
    for (const part of DocxTemplateEngine.XML_PARTS) {
      const file = DocxTemplateEngine.getZipPart(zip, part);
      if (!file) continue;
      const xml = DocxTemplateEngine.collapsePlaceholderRuns(await file.async("string"));
      for (const m of xml.matchAll(/\[([^\[\]]+)\]/g)) {
        const name = m[1].trim();
        if (name) found.add(name);
      }
    }
    return [...found];
  }

  _buildReplacements(rowData, headers, mapping) {
    const map = new Map();

    Object.entries(mapping).forEach(([placeholder, columnId]) => {
      if (!columnId) return;
      const value = String(rowData[columnId] ?? "");
      map.set(`[${placeholder}]`, this._escapeXml(value));
    });

    headers.forEach((h) => {
      const key = `[${h.label}]`;
      if (!map.has(key)) {
        map.set(key, this._escapeXml(String(rowData[h.id] ?? "")));
      }
    });

    return map;
  }

  _escapeXml(text) {
    return text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  _applyReplacements(xml, replacements) {
    let result = xml;
    [...replacements.entries()]
      .sort((a, b) => b[0].length - a[0].length)
      .forEach(([key, value]) => {
        result = result.split(key).join(value);
      });
    return result;
  }
}

class DocxPrintRenderer {
  static A4_WIDTH_PX = 794;
  static A4_HEIGHT_PX = 1123;

  static async extractPlainText(docxBlob) {
    const buffer = await docxBlob.arrayBuffer();
    const zip = await JSZip.loadAsync(buffer);
    const docFile = DocxTemplateEngine.getZipPart(zip, "word/document.xml");
    if (!docFile) return "";
    const xml = await docFile.async("string");
    return this._xmlToPlainText(xml);
  }

  static _xmlToPlainText(xml) {
    const withBreaks = xml.replace(/<w:br[^/]*\/>/g, "\n");
    const stripped = withBreaks.replace(/<[^>]+>/g, "");
    return stripped
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&amp;/g, "&")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  }

  static buildPrintHtml(sections) {
    const body = sections
      .map(
        (text, i) =>
          `<section class="summons-page${i < sections.length - 1 ? " page-break" : ""}"><pre>${this._escapeHtml(text)}</pre></section>`
      )
      .join("");
    return `<!DOCTYPE html><html lang="he" dir="rtl"><head><meta charset="UTF-8"/><title></title><style>
      @page { size: auto; margin: 0; }
      html, body { margin: 0; padding: 0; }
      body { font-family: "Segoe UI", Tahoma, Arial, sans-serif; direction: rtl; }
      .summons-page { padding: 2cm; box-sizing: border-box; }
      pre { white-space: pre-wrap; font-size: 18px; line-height: 1.6; margin: 0; }
      .page-break { page-break-after: always; }
    </style></head><body>${body}</body></html>`;
  }

  static _escapeHtml(text) {
    return text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  // Wraps pre-rendered (styled) docx HTML pages with page breaks for a single print job.
  static buildStyledHtml(pages) {
    const body = pages
      .map(
        (html, i) =>
          `<div class="qp-doc${i < pages.length - 1 ? " qp-break" : ""}">${html}</div>`
      )
      .join("");
    return `<!DOCTYPE html><html lang="he" dir="rtl"><head><meta charset="UTF-8"/><title></title><style>
      @page { size: 210mm 297mm; margin: 0; }
      html, body { margin: 0; padding: 0; width: 210mm; }
      body { direction: rtl; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .qp-break { page-break-after: always; break-after: page; }
      .qp-doc { width: ${this.A4_WIDTH_PX}px; margin: 0 auto; }
      .docx-wrapper { background: transparent !important; padding: 0 !important; margin: 0 !important; }
      .docx { box-shadow: none !important; margin: 0 auto !important; width: ${this.A4_WIDTH_PX}px !important; max-width: ${this.A4_WIDTH_PX}px !important; }
      section.docx { width: ${this.A4_WIDTH_PX}px !important; max-width: ${this.A4_WIDTH_PX}px !important; box-sizing: border-box; }
      @media print {
        html, body { width: 210mm; margin: 0; padding: 0; }
        .qp-doc, .docx-wrapper, .docx, section.docx {
          width: 210mm !important;
          max-width: 210mm !important;
          transform: none !important;
          zoom: 1 !important;
        }
        .docx-wrapper { padding: 0 !important; background: none !important; }
      }
    </style></head><body>${body}</body></html>`;
  }

  // Shared print path: renders filled DOCX blobs to styled HTML (keeping the template design),
  // falling back to plain text only if the local renderer is unavailable.
  static async printBlobs(blobs) {
    if (window.docx?.renderAsync) {
      try {
        return await this._printStyled(blobs);
      } catch (err) {
        console.warn("Styled DOCX print failed, falling back to plain text", err);
      }
    }
    return this._printPlain(blobs);
  }

  static async _printStyled(blobs) {
    const host = document.createElement("div");
    host.style.cssText = `position:fixed;left:-12000px;top:0;width:${this.A4_WIDTH_PX}px;min-height:${this.A4_HEIGHT_PX}px;overflow:hidden;visibility:hidden;pointer-events:none`;
    document.body.appendChild(host);
    try {
      const pages = [];
      for (const blob of blobs) {
        const container = document.createElement("div");
        container.style.width = `${this.A4_WIDTH_PX}px`;
        host.appendChild(container);
        await window.docx.renderAsync(blob, container, container, {
          className: "docx",
          inWrapper: false,
          ignoreWidth: false,
          ignoreHeight: false,
          breakPages: true,
          useBase64URL: true,
          hideWrapperOnPrint: true,
        });
        pages.push(container.innerHTML);
      }
      await this.printHtml(this.buildStyledHtml(pages));
    } finally {
      host.remove();
    }
  }

  static async _printPlain(blobs) {
    const sections = [];
    for (const blob of blobs) sections.push(await this.extractPlainText(blob));
    await this.printHtml(this.buildPrintHtml(sections));
  }

  static _waitForPrintAssets(doc, timeoutMs = 12000) {
    const images = [...(doc.images || [])];
    if (!images.length) return Promise.resolve();
    return Promise.all(
      images.map(
        (img) =>
          img.complete
            ? Promise.resolve()
            : new Promise((resolve) => {
                const done = () => resolve();
                img.addEventListener("load", done, { once: true });
                img.addEventListener("error", done, { once: true });
                setTimeout(done, timeoutMs);
              })
      )
    );
  }

  static printHtml(html) {
    return new Promise((resolve) => {
      const frame = document.createElement("iframe");
      frame.style.cssText = `position:fixed;left:-12000px;top:0;width:${this.A4_WIDTH_PX}px;height:${this.A4_HEIGHT_PX}px;border:0;visibility:hidden`;
      document.body.appendChild(frame);
      const win = frame.contentWindow;
      const doc = frame.contentDocument || win.document;
      doc.open();
      doc.write(html);
      doc.close();
      let printed = false;
      const finish = () => {
        if (printed) return;
        printed = true;
        try {
          win.focus();
          win.print();
        } catch (err) {
          console.warn("Browser print failed", err);
        }
        setTimeout(() => {
          frame.remove();
          resolve();
        }, 2000);
      };
      const schedulePrint = () => this._waitForPrintAssets(doc).then(finish);
      frame.onload = () => schedulePrint();
      setTimeout(() => schedulePrint(), 1200);
    });
  }
}

class FileDownloadHelper {
  static download(blob, filename) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  }

  static sanitizeFilename(name) {
    return (name || "זימון")
      .replace(/[\\/:*?"<>|]/g, "_")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 80) || "זימון";
  }
}

window.DocxTemplateEngine = DocxTemplateEngine;
window.DocxPrintRenderer = DocxPrintRenderer;
window.FileDownloadHelper = FileDownloadHelper;
