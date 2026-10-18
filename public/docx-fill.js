/* Fills Bible_Layout_Template.docx with a translation. Only text is replaced; every style, section, column,
   drop-cap, header and field in the template is left untouched. Works in the browser and in Node. */
(function (root) {
  const CODES = "GEN EXO LEV NUM DEU JOS JDG RUT 1SA 2SA 1KI 2KI 1CH 2CH EZR NEH EST JOB PSA PRO ECC SNG ISA JER LAM EZK DAN HOS JOL AMO OBA JON MIC NAM HAB ZEP HAG ZEC MAL MAT MRK LUK JHN ACT ROM 1CO 2CO GAL EPH PHP COL 1TH 2TH 1TI 2TI TIT PHM HEB JAS 1PE 2PE 1JN 2JN 3JN JUD REV".split(" ");
  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const T = (t) => `<w:r><w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;
  // the NBSP lives inside the small number run, so a line can never end on a number with its verse on the next line
  const NUM = (n) => `<w:r><w:rPr><w:rStyle w:val="VerseNum"/></w:rPr><w:t xml:space="preserve">${n}\u00A0</w:t></w:r>`;
  const P = (style, runs) => `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr>${runs}</w:p>`;
  const DROP = (n) => `<w:p><w:pPr><w:pStyle w:val="ChapterNumber"/><w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/></w:pPr>${T(n)}</w:p>`;
  const setText = (xml, from, to) => xml.split(`>${from}<`).join(`>${esc(to)}<`);

  // rows: [{ch, v, text}] -> chapter drop-cap + running-on verse paragraphs.
  // A verse whose text starts with "¶" begins a new paragraph.
  function chapters(rows) {
    const by = new Map();
    for (const r of rows) {
      if (!String(r.text || "").trim()) continue;
      if (!by.has(+r.ch)) by.set(+r.ch, []);
      by.get(+r.ch).push(r);
    }
    let out = "";
    for (const ch of [...by.keys()].sort((a, b) => a - b)) {
      const vs = by.get(ch).sort((a, b) => a.v - b.v);
      out += DROP(ch);
      let style = "VerseFirst", runs = "", first = true;
      for (const r of vs) {
        let txt = String(r.text).replace(/\s+/g, " ").trim(), brk = false;
        if (txt[0] === "¶") { brk = true; txt = txt.slice(1).trim(); }
        if (brk && !first) { out += P(style, runs); style = "Verse"; runs = ""; }
        runs += (first && +r.v === 1 ? "" : NUM(r.v)) + T(txt + " ");
        first = false;
      }
      out += P(style, runs);
    }
    return out;
  }

  async function fillBibleDocx(JSZip, templateBuf, o) {
    const zip = await JSZip.loadAsync(templateBuf);
    let d = await zip.file("word/document.xml").async("string");
    const headRe = /<w:p><w:pPr><w:pStyle w:val="Heading1"\/><\/w:pPr><w:bookmarkStart w:name="bk_([^"]+)"[^>]*\/><w:r><w:t xml:space="preserve">([^<]*)<\/w:t><\/w:r><w:bookmarkEnd[^>]*\/><\/w:p>/g;
    const heads = [...d.matchAll(headRe)];
    if (heads.length !== CODES.length) throw new Error("Templat tidak dijangka (" + heads.length + " kitab, perlu 66)");

    // book bodies, last to first so earlier offsets stay valid
    for (let i = heads.length - 1; i >= 0; i--) {
      const m = heads[i], code = CODES[i];
      const hEnd = m.index + m[0].length;
      const e1 = d.indexOf("</w:p>", hEnd) + 6;                       // end of the 1-column section break
      const bound = i < heads.length - 1 ? heads[i + 1].index : d.length;
      let endPos = d.lastIndexOf("<w:p><w:pPr><w:sectPr>", bound);    // the 2-column section break closing the book
      if (endPos < e1) endPos = d.lastIndexOf("<w:sectPr>");           // last book: closed by the body-level sectPr
      const name = (o.names && o.names[code]) || m[2];
      const head = m[0].replace(/(<w:t xml:space="preserve">)[^<]*(<\/w:t>)/, (_, a, b) => a + esc(name) + b);
      const body = chapters((o.verses && o.verses[code]) || []);
      if (!body) { d = d.slice(0, m.index) + head + d.slice(hEnd); continue; }   // no verses: keep the template's dummy text as is
      const keepTo = o.keepSectionHeading ? d.indexOf("</w:p>", e1) + 6 : e1;     // optional: keep the "Section Heading" line
      d = d.slice(0, m.index) + head + d.slice(hEnd, keepTo) + body + d.slice(endPos);
    }

    // contents page entries
    heads.forEach((m, i) => {
      const name = (o.names && o.names[CODES[i]]) || m[2];
      const re = new RegExp('(<w:t xml:space="preserve">)[^<]*(</w:t></w:r><w:r><w:t xml:space="preserve">\\t</w:t></w:r><w:fldSimple w:instr="PAGEREF bk_' + m[1] + ' )');
      d = d.replace(re, (_, a, b) => a + esc(name) + b);
    });

    // front page + testament titles
    if (o.title) d = setText(d, "BIBLE TITLE", o.title);
    if (o.line2) d = setText(d, "Translation · Edition · Year", o.line2);
    if (o.notes) d = setText(d, "Copyright, publisher, ISBN and edition notes go here.", o.notes);
    if (o.otTitle) d = setText(d, "The Old Testament", o.otTitle);
    if (o.ntTitle) d = setText(d, "The New Testament", o.ntTitle);

    // running heads: "page" -> "Muka Surat" (the PAGE / STYLEREF fields themselves are untouched)
    const pageWord = o.pageWord || "Muka Surat";
    for (const n of Object.keys(zip.files)) {
      if (!/^word\/header\d+\.xml$/.test(n)) continue;
      const h = await zip.file(n).async("string");
      zip.file(n, h.split('<w:t xml:space="preserve">page </w:t>').join('<w:t xml:space="preserve">' + esc(pageWord) + ' </w:t>'));
    }

    // optional typeface / body size, changed in styles.xml only (every style keeps its own size, spacing and layout)
    if (o.font || o.bodySize) {
      let st = await zip.file("word/styles.xml").async("string");
      const q = (f) => String(f).replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
      if (o.font) {
        const num = o.numFont || (/^georgia$/i.test(o.font) ? "Calibri" : o.font);   // Georgia digits are old-style, so keep lining digits for verse numbers
        st = st.split('"Garamond"').join('"' + q(o.font) + '"').split('"Calibri"').join('"' + q(num) + '"');
      }
      if (o.bodySize) {
        const sz = Math.round(o.bodySize * 2), line = Math.round(sz * 11.667);        // 9 pt -> sz 18, line 210 as in the template
        st = st.replace(/(<w:rPrDefault>[\s\S]*?<w:sz w:val=")\d+("\/><w:szCs w:val=")\d+/, "$1" + sz + "$2" + sz)
               .replace(/(<w:pPrDefault>[\s\S]*?w:line=")\d+/, "$1" + line)
               .replace(/(w:styleId="ChapterNumber"[\s\S]*?w:line=")\d+/, "$1" + line * 3);
      }
      zip.file("word/styles.xml", st);
    }

    zip.file("word/document.xml", d);
    return zip.generateAsync({ type: o.type || "blob", compression: "DEFLATE", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  }

  root.fillBibleDocx = fillBibleDocx;
  root.FILL_CODES = CODES;
  if (typeof module !== "undefined") module.exports = { fillBibleDocx, CODES };
})(typeof window !== "undefined" ? window : globalThis);
