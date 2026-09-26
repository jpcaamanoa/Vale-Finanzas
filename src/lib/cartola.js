import * as pdfjsLib from "pdfjs-dist";
import pdfjsWorkerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorkerUrl;

const DATE_RE = /(\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4})/;
const AMOUNT_RE = /-?\$?\s?\d{1,3}(?:[.,]\d{3})+(?:[.,]\d{2})?|-?\$?\s?\d{4,}/g;

function normalizeAmount(str) {
  const neg = /^-/.test(str.trim());
  const digits = str.replace(/[^\d]/g, "");
  if (!digits) return null;
  const n = parseInt(digits, 10);
  if (!n) return null;
  return neg ? -n : n;
}

function normalizeDate(str) {
  const m = str.match(/(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})/);
  if (!m) return null;
  let [, d, mo, y] = m;
  if (y.length === 2) y = "20" + y;
  if (Number(mo) > 12) [d, mo] = [mo, d]; // tolerate MM/DD stragglers
  d = d.padStart(2, "0");
  mo = mo.padStart(2, "0");
  if (Number(mo) < 1 || Number(mo) > 12) return null;
  return `${y}-${mo}-${d}`;
}

/** Reads a PDF file and returns its text content as one string per visual line. */
export async function extractLinesFromPdf(file) {
  const buf = await file.arrayBuffer();
  const doc = await pdfjsLib.getDocument({ data: buf }).promise;
  const lines = [];
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const content = await page.getTextContent();
    const byY = new Map();
    content.items.forEach((item) => {
      const y = Math.round(item.transform[5]);
      if (!byY.has(y)) byY.set(y, []);
      byY.get(y).push(item);
    });
    const ys = Array.from(byY.keys()).sort((a, b) => b - a);
    ys.forEach((y) => {
      const items = byY.get(y).sort((a, b) => a.transform[4] - b.transform[4]);
      const text = items
        .map((i) => i.str)
        .join(" ")
        .replace(/\s+/g, " ")
        .trim();
      if (text) lines.push(text);
    });
  }
  return lines;
}

/**
 * Best-effort extraction of transaction-looking rows from cartola text lines.
 * Bank statement layouts vary a lot, so this is a heuristic (date + an amount
 * on the same line) — always meant to be reviewed/edited before importing,
 * never trusted blindly.
 */
export function parseTransactions(lines) {
  const rows = [];
  lines.forEach((line) => {
    const dateMatch = line.match(DATE_RE);
    if (!dateMatch) return;
    const date = normalizeDate(dateMatch[1]);
    if (!date) return;

    // Only look for amounts AFTER the date ends, so the date's own digits
    // (e.g. the year) can never be mistaken for the transaction amount.
    const restOfLine = line.slice(dateMatch.index + dateMatch[0].length);
    const amounts = (restOfLine.match(AMOUNT_RE) || []).filter((a) => a.replace(/[^\d]/g, "").length >= 3);
    if (!amounts.length) return;
    // Common layout: description ... monto (posiblemente saldo al final) — tomamos el primer monto.
    const amountStr = amounts[0];
    const amount = normalizeAmount(amountStr);
    if (!amount) return;

    let desc = restOfLine;
    amounts.forEach((a) => {
      desc = desc.replace(a, "");
    });
    desc = desc.replace(/\s+/g, " ").trim();

    rows.push({ date, description: desc || "—", amount: Math.abs(amount) });
  });
  return rows;
}

/** Guesses a category id for a transaction description by keyword match. Returns null if none match. */
export function guessCategoryId(text, categories) {
  const t = (text || "").toLowerCase();
  for (const cat of categories || []) {
    if ((cat.keywords || []).some((kw) => kw && t.includes(String(kw).toLowerCase()))) return cat.id;
  }
  return null;
}
