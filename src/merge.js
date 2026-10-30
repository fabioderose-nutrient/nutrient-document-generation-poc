// ═══════════════════════════════════════════════════════════════════
//  MERGE ENGINE — replaces {{Field}} placeholders inside a DOCX.
//  Works at the OOXML level: for each paragraph whose combined text
//  contains a placeholder, runs are collapsed into a single run
//  (keeping the formatting of the first text run) so that placeholders
//  split across runs are still resolved.
// ═══════════════════════════════════════════════════════════════════

function escapeXml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

const PLACEHOLDER_RE = /\{\{\s*([^{}]+?)\s*\}\}/g;

function mergeXml(xml, record) {
  return xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g, (para) => {
    const texts = [...para.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map((m) => m[1]);
    const joined = texts.join("");
    PLACEHOLDER_RE.lastIndex = 0;
    if (!PLACEHOLDER_RE.test(joined)) return para;

    const mergedText = joined.replace(PLACEHOLDER_RE, (m, key) => {
      const val = record[key.trim()];
      return val !== undefined && val !== null ? escapeXml(val) : m;
    });

    // Paragraph properties (may include the paragraph-mark run props)
    const pPrMatch = para.match(/<w:pPr(?:\s[^>]*)?>[\s\S]*?<\/w:pPr>|<w:pPr\/>/);
    const pPr = pPrMatch ? pPrMatch[0] : "";

    // First run properties outside of pPr
    const bodyWithoutPPr = pPr ? para.replace(pPr, "") : para;
    const rPrMatch = bodyWithoutPPr.match(/<w:rPr(?:\s[^>]*)?>[\s\S]*?<\/w:rPr>/);
    const rPr = rPrMatch ? rPrMatch[0] : "";

    const lines = mergedText.split(/\n/);
    const textXml = lines
      .map((line) => `<w:t xml:space="preserve">${line}</w:t>`)
      .join("<w:br/>");

    return `<w:p>${pPr}<w:r>${rPr}${textXml}</w:r></w:p>`;
  });
}

// Merge a record into a DOCX (ArrayBuffer or Blob) → returns a Blob.
export async function mergeDocx(docxBufferOrBlob, record) {
  const JSZip = (await import("jszip")).default;
  const zip = await JSZip.loadAsync(docxBufferOrBlob);
  const targets = Object.keys(zip.files).filter((p) =>
    /^word\/(document|header\d*|footer\d*)\.xml$/.test(p)
  );
  for (const path of targets) {
    const xml = await zip.file(path).async("string");
    zip.file(path, mergeXml(xml, record));
  }
  return zip.generateAsync({ type: "blob" });
}

// List distinct placeholder names present in a DOCX.
export async function listDocxFields(docxBufferOrBlob) {
  const JSZip = (await import("jszip")).default;
  const zip = await JSZip.loadAsync(docxBufferOrBlob);
  const targets = Object.keys(zip.files).filter((p) =>
    /^word\/(document|header\d*|footer\d*)\.xml$/.test(p)
  );
  const found = new Set();
  for (const path of targets) {
    const xml = await zip.file(path).async("string");
    for (const para of xml.matchAll(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g)) {
      const joined = [...para[0].matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map((m) => m[1]).join("");
      for (const m of joined.matchAll(PLACEHOLDER_RE)) found.add(m[1].trim());
    }
  }
  return [...found];
}
