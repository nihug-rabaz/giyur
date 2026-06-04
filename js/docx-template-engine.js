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
    replacements.forEach((value, key) => {
      result = result.split(key).join(value);
    });
    return result;
  }
}

class DocxPrintRenderer {
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
      @page { size: auto; margin: 0; }
      html, body { margin: 0; padding: 0; }
      .qp-break { page-break-after: always; }
      .docx-wrapper { background: transparent !important; padding: 0 !important; }
      .docx { box-shadow: none !important; margin: 0 auto !important; }
    </style></head><body>${body}</body></html>`;
  }

  // Shared print path: renders filled DOCX blobs to styled HTML (keeping the template design),
  // falling back to plain text only if the local renderer is unavailable.
  static async printBlobs(blobs) {
    if (window.docx?.renderAsync) return this._printStyled(blobs);
    return this._printPlain(blobs);
  }

  static async _printStyled(blobs) {
    const host = document.createElement("div");
    host.style.cssText = "position:fixed;left:-99999px;top:0;width:0;height:0;overflow:hidden";
    document.body.appendChild(host);
    try {
      const pages = [];
      for (const blob of blobs) {
        const container = document.createElement("div");
        host.appendChild(container);
        await window.docx.renderAsync(blob, container, container, {
          className: "docx",
          inWrapper: true,
          ignoreWidth: false,
          ignoreHeight: false,
          breakPages: true,
          useBase64URL: true,
        });
        pages.push(container.innerHTML);
      }
      this.printHtml(this.buildStyledHtml(pages));
    } finally {
      host.remove();
    }
  }

  static async _printPlain(blobs) {
    const sections = [];
    for (const blob of blobs) sections.push(await this.extractPlainText(blob));
    this.printHtml(this.buildPrintHtml(sections));
  }

  static printHtml(html) {
    const frame = document.createElement("iframe");
    frame.style.cssText = "position:fixed;right:0;bottom:0;width:0;height:0;border:0";
    document.body.appendChild(frame);
    const doc = frame.contentDocument || frame.contentWindow.document;
    doc.open();
    doc.write(html);
    doc.close();
    frame.onload = () => {
      frame.contentWindow.focus();
      frame.contentWindow.print();
      setTimeout(() => frame.remove(), 1000);
    };
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
