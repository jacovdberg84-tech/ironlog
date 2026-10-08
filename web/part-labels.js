/**
 * IRONLOG part labels: IronLog's own label for parts that come without a
 * maker's barcode. Each label carries the part code as a Code 128 barcode
 * (read by any scanner) and as a QR (phones), with the name and bin.
 *
 * Needs vendor/jsbarcode-code128-3.11.6.min.js and vendor/qrcode-generator-1.4.4.js.
 * Used by Stock Control (app/stock-labels.js) and the stores terminal.
 *
 *   IronlogPartLabels.print([{ part_code, part_name, bin, copies }], { layout, skip })
 */
(function (global) {
  const LAYOUT_KEY = "ironlog-label-layout";
  const LAYOUTS = {
    "a4-21": { name: "A4 sheet · 21 labels, 63.5 × 38.1 mm (Avery L7160 / J8160)", w: 63.5, h: 38.1, cols: 3, rows: 7, top: 15.15, left: 7.21, gapX: 2.54, gapY: 0, qr: true },
    "a4-24": { name: "A4 sheet · 24 labels, 70 × 37 mm", w: 70, h: 37, cols: 3, rows: 8, top: 0.5, left: 0, gapX: 0, gapY: 0, qr: true },
    "a4-14": { name: "A4 sheet · 14 labels, 99.1 × 38.1 mm (Avery L7163)", w: 99.1, h: 38.1, cols: 2, rows: 7, top: 15.15, left: 4.65, gapX: 2.54, gapY: 0, qr: true },
    "roll-62x29": { name: "Label printer · 62 × 29 mm (Brother DK-11209)", w: 62, h: 29, roll: true, qr: true },
    "roll-50x25": { name: "Label printer · 50 × 25 mm", w: 50, h: 25, roll: true, qr: false },
    "roll-100x50": { name: "Label printer · 100 × 50 mm (Zebra)", w: 100, h: 50, roll: true, qr: true },
  };

  const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

  function savedLayout() {
    try {
      const k = localStorage.getItem(LAYOUT_KEY);
      return LAYOUTS[k] ? k : "a4-21";
    } catch { return "a4-21"; }
  }
  function saveLayout(k) {
    try { if (LAYOUTS[k]) localStorage.setItem(LAYOUT_KEY, k); } catch { /* storage off */ }
  }

  function barcodeSvg(code) {
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    // 10-module quiet zone either side: scanners need the white margin.
    global.JsBarcode(svg, code, { format: "CODE128", displayValue: false, margin: 0, marginLeft: 20, marginRight: 20, height: 50, width: 2 });
    // Stretch to the label width; bars keep their proportions across the width.
    svg.setAttribute("preserveAspectRatio", "none");
    svg.removeAttribute("width");
    svg.removeAttribute("height");
    svg.removeAttribute("style");
    return svg.outerHTML;
  }

  function qrSvg(text) {
    const qr = global.qrcode(0, "M");
    qr.addData(text);
    qr.make();
    return qr.createSvgTag({ cellSize: 2, margin: 0, scalable: true });
  }

  /** The text inside a label's QR: the part's store link, which every IronLog scan reads as the part. */
  function qrText(partCode, site) {
    const origin = /^https?:$/.test(global.location?.protocol || "") ? global.location.origin : "";
    return `${origin}/web/store-mobile.html?site=${encodeURIComponent(site || "main")}&part_code=${encodeURIComponent(partCode)}`;
  }

  /** Code text size (mm) so the whole part code fits on one line. */
  function codeFont(code, widthMm, maxMm) {
    const perChar = 0.74; // bold Arial capitals and digits: about 0.72 × font size wide
    return Math.max(2, Math.min(maxMm, widthMm / (Math.max(String(code).length, 1) * perChar)));
  }

  function labelHtml(item, L, site, m) {
    const textWidth = L.w - 2 * m.pad - (L.qr ? m.qr + m.pad : 0);
    return `
      <div class="lbl">
        <div class="top">
          <div class="txt">
            <div class="code" style="font-size:${codeFont(item.part_code, textWidth, m.code).toFixed(2)}mm">${esc(item.part_code)}</div>
            <div class="name">${esc(item.part_name || "")}</div>
            ${item.bin ? `<div class="bin">BIN ${esc(item.bin)}</div>` : ""}
          </div>
          ${L.qr ? `<div class="qr">${qrSvg(qrText(item.part_code, site))}</div>` : ""}
        </div>
        <div class="bc">${barcodeSvg(item.part_code)}</div>
      </div>`;
  }

  /** The printable document for these labels. */
  function buildDocument(items, { layout = savedLayout(), skip = 0, site = "main" } = {}) {
    const L = LAYOUTS[layout] || LAYOUTS["a4-21"];
    const labels = [];
    for (const it of items) {
      const n = Math.max(1, Math.min(100, Math.round(Number(it.copies || 1))));
      for (let i = 0; i < n; i += 1) labels.push(it);
    }
    if (!labels.length) throw new Error("No labels to print");
    const pad = Math.min(L.h * 0.07, 2.5);
    const codeSize = Math.max(3.2, Math.min(L.h * 0.17, 7));
    const bcH = L.h * 0.34; // the barcode: full width along the bottom
    const qr = L.h - 2 * pad - bcH - pad; // the QR: top right, beside the text
    const m = { pad, code: codeSize, qr };
    const small = Math.max(2.1, codeSize * 0.45);
    const css = `
      * { box-sizing: border-box; margin: 0; padding: 0; }
      html, body { background: #fff; color: #000; font-family: Arial, Helvetica, sans-serif; }
      .lbl { width: ${L.w}mm; height: ${L.h}mm; padding: ${pad}mm; display: flex; flex-direction: column; gap: ${pad}mm; overflow: hidden; }
      .top { flex: 1; min-height: 0; display: flex; gap: ${pad}mm; }
      .qr { flex: 0 0 ${qr}mm; height: ${qr}mm; }
      .qr svg { width: 100%; height: 100%; display: block; }
      .txt { flex: 1; min-width: 0; display: flex; flex-direction: column; }
      .code { font-weight: 700; line-height: 1.05; white-space: nowrap; overflow: hidden; }
      .name { font-size: ${small}mm; line-height: 1.15; max-height: 2.3em; overflow: hidden; margin-top: 0.4mm; }
      .bin { font-size: ${small}mm; font-weight: 700; margin-top: 0.4mm; }
      .bc { flex: 0 0 ${bcH}mm; height: ${bcH}mm; }
      .bc svg { width: 100%; height: 100%; display: block; }
      ${L.roll
        ? `@page { size: ${L.w}mm ${L.h}mm; margin: 0; } .lbl { page-break-after: always; break-after: page; } .lbl:last-child { page-break-after: auto; break-after: auto; }`
        : `@page { size: A4; margin: 0; }
           .page { width: 210mm; height: 297mm; padding: ${L.top}mm 0 0 ${L.left}mm; display: grid; align-content: start;
             grid-template-columns: repeat(${L.cols}, ${L.w}mm); grid-auto-rows: ${L.h}mm; column-gap: ${L.gapX}mm; row-gap: ${L.gapY}mm;
             page-break-after: always; break-after: page; overflow: hidden; }
           .page:last-child { page-break-after: auto; break-after: auto; }`}
    `;
    let body = "";
    if (L.roll) {
      body = labels.map((it) => labelHtml(it, L, site, m)).join("");
    } else {
      const perPage = L.cols * L.rows;
      const start = Math.max(0, Math.min(perPage - 1, Math.round(Number(skip) || 0)));
      const cells = [...Array(start).fill(null), ...labels];
      for (let i = 0; i < cells.length; i += perPage) {
        body += `<div class="page">${cells.slice(i, i + perPage).map((it) => (it ? labelHtml(it, L, site, m) : `<div></div>`)).join("")}</div>`;
      }
    }
    return { html: `<!doctype html><html><head><meta charset="utf-8"><title>IronLog part labels</title><style>${css}</style></head><body>${body}</body></html>`, count: labels.length };
  }

  /** Print through a hidden frame (works where pop-ups are blocked). */
  function print(items, opts = {}) {
    const { html, count } = buildDocument(items, opts);
    if (opts.layout) saveLayout(opts.layout);
    const old = document.getElementById("ironlogLabelFrame");
    if (old) old.remove();
    const frame = document.createElement("iframe");
    frame.id = "ironlogLabelFrame";
    frame.setAttribute("aria-hidden", "true");
    frame.style.cssText = "position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden";
    document.body.appendChild(frame);
    const doc = frame.contentDocument;
    doc.open();
    doc.write(html);
    doc.close();
    setTimeout(() => {
      try {
        frame.contentWindow.focus();
        frame.contentWindow.print();
      } catch (e) {
        console.error(e);
      }
    }, 250);
    return count;
  }

  global.IronlogPartLabels = { LAYOUTS, buildDocument, print, savedLayout, saveLayout };
})(window);
