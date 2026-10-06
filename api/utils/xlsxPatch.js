// IRONLOG/api/utils/xlsxPatch.js — write values into chosen cells of an existing
// .xlsx/.xlsm without rebuilding the workbook.
//
// Workbook libraries rewrite the whole file and drop what they do not
// understand (macros, pivot caches, add-in sheets, images, data validation
// lists). Here only the sheet XML of the cells being written changes: every
// other part of the package is copied as it is. Strings go in as inline
// strings (no shared-string table edits); the cell's existing style is kept;
// Excel is told to recalculate every formula when the file is opened.

import JSZip from "jszip";

const XML_ESC = { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" };
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => XML_ESC[c]);
const unesc = (s) => String(s)
  .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");

/** "AB12" → { col: 28, row: 12 } */
export function parseRef(ref) {
  const m = /^([A-Z]+)(\d+)$/.exec(String(ref).toUpperCase());
  if (!m) throw new Error(`Bad cell reference: ${ref}`);
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { col, row: Number(m[2]), letters: m[1] };
}

export function colLetters(n) {
  let s = "";
  let x = Number(n);
  while (x > 0) {
    const r = (x - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

/** Excel serial date (1900 system) for YYYY-MM-DD. */
export function excelDate(ymd) {
  const t = Date.parse(`${String(ymd).slice(0, 10)}T00:00:00Z`);
  return Math.round(t / 86400000) + 25569;
}

export class XlsxPatcher {
  static async open(buffer) {
    const p = new XlsxPatcher();
    p.zip = await JSZip.loadAsync(buffer, { createFolders: false });
    const wb = await p.zip.file("xl/workbook.xml").async("string");
    const rels = await p.zip.file("xl/_rels/workbook.xml.rels").async("string");
    const target = {};
    for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
      const id = /Id="([^"]+)"/.exec(m[0])?.[1];
      const t = /Target="([^"]+)"/.exec(m[0])?.[1];
      if (id && t) target[id] = t.startsWith("/") ? t.slice(1) : `xl/${t}`;
    }
    p.sheets = {};
    for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
      const name = unesc(/name="([^"]+)"/.exec(m[0])?.[1] || "");
      const rid = /r:id="([^"]+)"/.exec(m[0])?.[1];
      if (name && rid && target[rid]) p.sheets[name] = target[rid];
    }
    p.workbookXml = wb;
    p.xml = new Map();
    p.shared = null;
    return p;
  }

  hasSheet(name) {
    return Boolean(this.sheets[name]);
  }

  async sheetXml(name) {
    const path = this.sheets[name];
    if (!path) throw new Error(`Sheet not found in the template: ${name}`);
    if (!this.xml.has(path)) this.xml.set(path, await this.zip.file(path).async("string"));
    return this.xml.get(path);
  }

  async sharedStrings() {
    if (this.shared) return this.shared;
    const f = this.zip.file("xl/sharedStrings.xml");
    const xml = f ? await f.async("string") : "";
    this.shared = [...xml.matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => unesc([...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((t) => t[1]).join("")));
    return this.shared;
  }

  /** Values of the given columns for every row of a sheet: Map(row → { D: value, … }). */
  async readColumns(name, letters) {
    const xml = await this.sheetXml(name);
    const shared = await this.sharedStrings();
    const want = new Set(letters);
    const out = new Map();
    for (const m of xml.matchAll(/<c r="([A-Z]+)(\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      if (!want.has(m[1])) continue;
      const attrs = m[3] || "";
      const body = m[4] || "";
      const type = /\bt="([^"]+)"/.exec(attrs)?.[1];
      let v = /<v>([\s\S]*?)<\/v>/.exec(body)?.[1];
      if (type === "inlineStr") v = [...body.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((t) => t[1]).join("");
      if (v == null || v === "") continue;
      let value = type === "s" ? shared[Number(v)] : unesc(v);
      if (!type || type === "n") value = Number(v);
      const row = Number(m[2]);
      if (!out.has(row)) out.set(row, {});
      out.get(row)[m[1]] = value;
    }
    return out;
  }

  /**
   * cells: { "D6": "E502AM", "F6": 40, "C3": null (clear) }. Numbers are written
   * as numbers, everything else as text; null/"" clears the value. A cell
   * holding a formula is never overwritten.
   */
  async setCells(name, cells) {
    const path = this.sheets[name];
    let xml = await this.sheetXml(name);
    const byRow = new Map();
    for (const [ref, value] of Object.entries(cells)) {
      const p = parseRef(ref);
      if (!byRow.has(p.row)) byRow.set(p.row, []);
      byRow.get(p.row).push({ ...p, ref: `${p.letters}${p.row}`, value });
    }
    const sdStart = xml.indexOf("<sheetData");
    const sdOpenEnd = xml.indexOf(">", sdStart) + 1;
    const selfClosed = xml[sdOpenEnd - 2] === "/";
    let sheetData = selfClosed ? "" : xml.slice(sdOpenEnd, xml.indexOf("</sheetData>"));
    const head = selfClosed ? xml.slice(0, sdStart) + "<sheetData>" : xml.slice(0, sdOpenEnd);
    const tail = selfClosed ? "</sheetData>" + xml.slice(sdOpenEnd) : xml.slice(xml.indexOf("</sheetData>"));

    for (const [rowNo, list] of byRow) {
      const rowRe = new RegExp(`<row r="${rowNo}"(?=[\\s>/])[^>]*?(?:/>|>[\\s\\S]*?</row>)`);
      let m = rowRe.exec(sheetData);
      if (!m) {
        // Insert a new row in order.
        const newRow = `<row r="${rowNo}"></row>`;
        let at = sheetData.length;
        for (const r of sheetData.matchAll(/<row r="(\d+)"/g)) {
          if (Number(r[1]) > rowNo) { at = r.index; break; }
        }
        sheetData = sheetData.slice(0, at) + newRow + sheetData.slice(at);
        m = rowRe.exec(sheetData);
      }
      let rowXml = m[0];
      if (rowXml.endsWith("/>")) rowXml = `${rowXml.slice(0, -2)}></row>`;
      const open = rowXml.slice(0, rowXml.indexOf(">") + 1);
      let body = rowXml.slice(open.length, rowXml.length - "</row>".length);
      for (const c of list.sort((a, b) => a.col - b.col)) {
        const cellRe = new RegExp(`<c r="${c.ref}"(?=[\\s>/])([^>]*?)(?:/>|>([\\s\\S]*?)</c>)`);
        const cm = cellRe.exec(body);
        if (cm && /<f[\s>/]/.test(cm[2] || "")) continue; // never replace a formula
        const style = cm ? (/\bs="(\d+)"/.exec(cm[1])?.[1] ?? null) : null;
        const sAttr = style != null ? ` s="${style}"` : "";
        let cell;
        if (c.value == null || c.value === "") cell = `<c r="${c.ref}"${sAttr}/>`;
        else if (typeof c.value === "number" && Number.isFinite(c.value)) cell = `<c r="${c.ref}"${sAttr}><v>${c.value}</v></c>`;
        else cell = `<c r="${c.ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${esc(c.value)}</t></is></c>`;
        if (cm) {
          body = body.slice(0, cm.index) + cell + body.slice(cm.index + cm[0].length);
        } else {
          let at = body.length;
          for (const x of body.matchAll(/<c r="([A-Z]+)\d+"/g)) {
            if (parseRef(`${x[1]}1`).col > c.col) { at = x.index; break; }
          }
          body = body.slice(0, at) + cell + body.slice(at);
        }
      }
      sheetData = sheetData.slice(0, m.index) + open + body + "</row>" + sheetData.slice(m.index + m[0].length);
    }
    xml = head + sheetData + tail;
    this.xml.set(path, xml);
  }

  /** Clears every cell of the given columns between two rows (inclusive). */
  async clearRange(name, letters, fromRow, toRow) {
    const cells = {};
    for (let r = fromRow; r <= toRow; r += 1) for (const l of letters) cells[`${l}${r}`] = null;
    await this.setCells(name, cells);
  }

  async save() {
    for (const [path, xml] of this.xml) this.zip.file(path, xml, { createFolders: false });
    // Excel recalculates every formula when the file is opened.
    let wb = this.workbookXml;
    if (/<calcPr\b/.test(wb)) {
      wb = wb.replace(/<calcPr\b([^>]*?)(\/?)>/, (all, attrs, slash) => `<calcPr${attrs.replace(/\s*fullCalcOnLoad="[^"]*"/, "")} fullCalcOnLoad="1"${slash}>`);
    } else {
      wb = wb.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>');
    }
    this.zip.file("xl/workbook.xml", wb, { createFolders: false });
    return this.zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } });
  }
}
