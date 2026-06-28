// Splits a (UTF-8) string into ordered QR-sized chunks and reassembles them back.
// Wire format per chunk: "GQR1|<index>|<total>|<base64piece>" (index is 1-based).
class QrChunker {
  static PREFIX = "GQR1";
  static DEFAULT_CHUNK = 600;

  // Returns an array of payload strings, one per QR code.
  static encode(text, chunkSize = QrChunker.DEFAULT_CHUNK) {
    const b64 = QrChunker._toBase64(text || "");
    const pieces = QrChunker._slice(b64, chunkSize);
    const total = pieces.length || 1;
    return pieces.length
      ? pieces.map((piece, i) => `${QrChunker.PREFIX}|${i + 1}|${total}|${piece}`)
      : [`${QrChunker.PREFIX}|1|1|`];
  }

  // Parses one scanned payload into { index, total, piece } or null if it isn't ours.
  static parse(payload) {
    const text = String(payload || "");
    if (!text.startsWith(`${QrChunker.PREFIX}|`)) return null;
    const parts = text.split("|");
    if (parts.length < 4) return null;
    const index = Number(parts[1]);
    const total = Number(parts[2]);
    const piece = parts.slice(3).join("|");
    if (!Number.isInteger(index) || !Number.isInteger(total) || index < 1 || total < 1) return null;
    return { index, total, piece };
  }

  // Reassembles ordered pieces (array indexed by chunk order) into the original string.
  static join(pieces) {
    return QrChunker._fromBase64(pieces.join(""));
  }

  static _slice(str, size) {
    const out = [];
    for (let i = 0; i < str.length; i += size) out.push(str.slice(i, i + size));
    return out;
  }

  static _toBase64(str) {
    const bytes = new TextEncoder().encode(str);
    let binary = "";
    bytes.forEach((b) => { binary += String.fromCharCode(b); });
    return btoa(binary);
  }

  static _fromBase64(b64) {
    const binary = atob(b64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return new TextDecoder().decode(bytes);
  }
}

if (typeof window !== "undefined") window.QrChunker = QrChunker;
