/* Fills Bible_Layout_Template.docx with a translation. Only text is replaced; every style, section, column,
   drop-cap, header and field in the template is left untouched. Works in the browser and in Node. */
(function (root) {
  const CODES = "GEN EXO LEV NUM DEU JOS JDG RUT 1SA 2SA 1KI 2KI 1CH 2CH EZR NEH EST JOB PSA PRO ECC SNG ISA JER LAM EZK DAN HOS JOL AMO OBA JON MIC NAM HAB ZEP HAG ZEC MAL MAT MRK LUK JHN ACT ROM 1CO 2CO GAL EPH PHP COL 1TH 2TH 1TI 2TI TIT PHM HEB JAS 1PE 2PE 1JN 2JN 3JN JUD REV".split(" ");
  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const T = (t) => `<w:r><w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;
  // the NBSP lives inside the small number run, so a line can never end on a number with its verse on the next line
  const NUM = (n) => `<w:r><w:rPr><w:rStyle w:val="VerseNum"/></w:rPr><w:t xml:space="preserve">${n}\u00A0</w:t></w:r>`;
  const P = (style, runs, gap) => `<w:p><w:pPr><w:pStyle w:val="${style}"/>${gap ? `<w:spacing w:after="${gap}"/>` : ""}</w:pPr>${runs}</w:p>`;
  const DROP = (n) => `<w:p><w:pPr><w:pStyle w:val="ChapterNumber"/><w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/></w:pPr>${T(n)}</w:p>`;
  const setText = (xml, from, to) => xml.split(`>${from}<`).join(`>${esc(to)}<`);

  // ---- Malay hyphenation (Pedoman Umum Ejaan, MABBIM 1991): optional (soft) hyphens; Word shows one only at a line end ----
  // 1) base words by syllable: V-V split (ai/au/oi never), V-C-V -> V-CV, V-CC-V -> VC-CV, 3+ consonants -> after the first;
  //    ng ny sy kh gh count as one consonant; a single-vowel syllable is never left at the end or start of a line.
  // 2) derived words may also break at prefix/suffix boundaries (makan-an, me-rasa-kan, mem-bantu, pergi-lah); the root is kept
  //    whole when short, and "-i" is never split. An affix is only recognised if its root also occurs as a word in the
  //    same translation, so "perang" or "terang" are not cut as per-ang / ter-ang.
  const isV = (c) => c !== undefined && "aeiou".includes(c.toLowerCase());
  const dip = (a, b) => ["ai", "au", "oi"].includes((a + b).toLowerCase());
  const DIG = ["ng", "ny", "sy", "kh", "gh"], DIG_OLD = DIG.concat(["dj", "tj", "sj", "nj", "ch"]);   // old spelling: dj tj sj nj ch
  function units(w, old) {
    const D = old ? DIG_OLD : DIG, u = [];
    for (let i = 0; i < w.length;) {
      const two = w.slice(i, i + 2).toLowerCase();
      if (old && two === "oe") { u.push({ v: true, s: i }); i += 2; }
      else if (isV(w[i])) { u.push({ v: true, s: i }); i++; }
      else { u.push({ v: false, s: i }); i += D.includes(two) ? 2 : 1; }
    }
    return u;
  }
  function syllBreaks(w, old) {
    const u = units(w, old), pts = [];
    for (let k = 0; k < u.length; k++) {
      if (!u[k].v) continue;
      if (k + 1 < u.length && u[k + 1].v) { if (!dip(w[u[k].s], w[u[k + 1].s])) pts.push(u[k + 1].s); continue; }
      let j = k + 1;
      while (j < u.length && !u[j].v) j++;
      const n = j - k - 1;
      if (j >= u.length || n === 0) continue;
      pts.push(n === 1 ? u[k + 1].s : u[k + 2].s);
    }
    return pts;
  }
  const loneStart = (w, i) => isV(w[i]) && (i + 1 >= w.length || (isV(w[i + 1]) && !dip(w[i], w[i + 1])));
  const loneEnd = (w, i) => isV(w[i - 1]) && (i < 2 || (isV(w[i - 2]) && !dip(w[i - 2], w[i - 1])));
  const okBreak = (w, i, after) => i >= 2 && w.length - i >= after && !loneStart(w, i) && !loneEnd(w, i);

  const PRE = ["memper", "diper", "terper", "meng", "meny", "mem", "men", "peng", "peny", "pem", "pen", "ber", "ter", "per", "me", "pe", "di", "ke", "se"];
  const RESTORE = { meng: ["", "k"], meny: ["s"], mem: ["", "p"], men: ["", "t"], peng: ["", "k"], peny: ["s"], pem: ["", "p"], pen: ["", "t"] };
  function splitWord(w, h) {
    const SUF = h.old ? ["kan", "lah", "kah", "pun", "nya", "nja", "an", ""] : ["kan", "lah", "kah", "pun", "nya", "an", ""];
    const known = (r) => r.length >= 4 && ((h.V.get(r) || 0) >= 2 || (r.length >= 5 && r.endsWith("i") && (h.V.get(r.slice(0, -1)) || 0) >= 2));
    let best = null;
    for (const s of SUF) for (const p of [""].concat(PRE)) {
      if (!p && !s) continue;
      const mid = w.slice(p.length, w.length - s.length);
      if (w.slice(0, p.length) !== p || (s && w.slice(w.length - s.length) !== s) || mid.length < 3) continue;
      for (const r0 of RESTORE[p] || [""]) {
        if (!known(r0 + mid)) continue;
        const c = { p, s, dropped: r0 !== "", score: p.length + s.length - (r0 ? 0.5 : 0) };
        if (!best || c.score > best.score) best = c;
      }
    }
    return best;
  }
  function pointsFor(w, h) {
    const sp = splitWord(w, h);
    let pts;
    if (!sp) pts = syllBreaks(w, h.old).filter((i) => okBreak(w, i, 3));
    else {
      const pe = sp.p.length, ss = w.length - sp.s.length;
      pts = [];
      if (sp.dropped) pts = syllBreaks(w, h.old).filter((i) => i < ss && okBreak(w, i, 3));      // meN-/peN- that swallowed the root's first letter: me-nu-lis
      else {
        if (sp.p && okBreak(w, pe, 3)) pts.push(pe);
        if (w.slice(pe, ss).length >= h.rootSplit) pts = pts.concat(syllBreaks(w.slice(pe, ss), h.old).map((i) => i + pe).filter((i) => okBreak(w, i, 3)));
      }
      if (sp.s && okBreak(w, ss, 2)) pts.push(ss);
    }
    pts = [...new Set(pts)].sort((a, b) => a - b);
    for (let i = pts.length - 2; i >= 0; i--) if (pts[i + 1] - pts[i] < 2) pts.splice(i, 1);   // never leave a one-letter piece (suffix boundary wins)
    return pts;
  }
  function makeHyph(o) {
    const V = new Map();
    let oldHits = 0;
    for (const c in o.verses || {}) for (const r of o.verses[c]) for (const m of String(r.text || "").toLowerCase().matchAll(/[a-z]+/g)) {
      V.set(m[0], (V.get(m[0]) || 0) + 1);
      if (/dj|tj|sj/.test(m[0])) oldHits++;
    }
    const ex = {};
    for (const [k, v] of Object.entries(o.hyphenExceptions || {})) {                       // {"Yerusalem": "Ye-ru-sa-lem"}
      let n = 0;
      ex[k.toLowerCase()] = v.split("-").slice(0, -1).map((part) => (n += part.length));
    }
    return { V, ex, cache: new Map(), rootSplit: o.rootSplit == null ? 7 : o.rootSplit,
             old: o.oldSpelling != null ? !!o.oldSpelling : oldHits >= 10 };                // old spelling is auto-detected (dj/tj/sj words)
  }
  const SOFT = '</w:t><w:softHyphen/><w:t xml:space="preserve">';
  function TH(text, h) {
    if (!h) return T(text);
    let out = "", last = 0, m;
    const re = /[A-Za-z]{7,}/g;
    while ((m = re.exec(text))) {
      out += esc(text.slice(last, m.index));
      const w = m[0], k = w.toLowerCase();
      if (!h.cache.has(k)) h.cache.set(k, h.ex[k] || pointsFor(k, h));
      let a = 0;
      for (const i of h.cache.get(k)) { out += esc(w.slice(a, i)) + SOFT; a = i; }
      out += esc(w.slice(a));
      last = m.index + w.length;
    }
    return `<w:r><w:t xml:space="preserve">${out}${esc(text.slice(last))}</w:t></w:r>`;
  }

  // rows: [{ch, v, text}] -> chapter drop-cap + verse paragraphs.
  // A verse starting with "¶" always starts a paragraph. If a chapter has no ¶ marks, paragraphs hold paraMin..paraMax verses
  // (a new one starts at the first verse after a sentence end once paraMin is reached, forced at paraMax), and a lone
  // verse is never left on its own: it joins the previous paragraph.
  // Paragraphs have no gap between them; only the end of a chapter gets one blank line (never on the drop-cap paragraph itself).
  const SPACER = '<w:p><w:pPr><w:pStyle w:val="Verse"/><w:ind w:firstLine="0"/></w:pPr></w:p>';
  function chapters(rows, opt) {
    const by = new Map();
    for (const r of rows) {
      if (!String(r.text || "").trim()) continue;
      if (!by.has(+r.ch)) by.set(+r.ch, []);
      by.get(+r.ch).push(r);
    }
    let out = "";
    for (const ch of [...by.keys()].sort((a, b) => a - b)) {
      const vs = by.get(ch).sort((a, b) => a.v - b.v);
      const manual = vs.some((r) => String(r.text).trim()[0] === "¶");
      const groups = [[]];
      let prev = "";
      for (const r of vs) {
        let txt = String(r.text).replace(/\s+/g, " ").trim(), brk = false;
        const cur = groups[groups.length - 1];
        if (txt[0] === "¶") { brk = true; txt = txt.slice(1).trim(); }
        else if (!manual && opt.max && cur.length >= opt.min && (cur.length >= opt.max || /[.!?]["”’»)]*$/.test(prev))) brk = true;
        if (brk && cur.length) groups.push([]);
        groups[groups.length - 1].push({ v: r.v, txt });
        prev = txt;
      }
      if (!manual && groups.length > 1 && groups[groups.length - 1].length === 1) groups[groups.length - 2].push(groups.pop()[0]);
      out += DROP(ch);
      groups.forEach((g, gi) => {
        const style = gi === 0 ? "VerseFirst" : "Verse", last = gi === groups.length - 1;
        let runs = "";
        g.forEach((x, k) => { runs += (gi === 0 && k === 0 && +x.v === 1 ? "" : NUM(x.v)) + TH(x.txt + " ", opt.hyph); });
        out += P(style, runs, last && style === "Verse" ? opt.gap : 0);
        if (last && style === "VerseFirst" && opt.gap) out += SPACER;
      });
    }
    return out;
  }

  async function fillBibleDocx(JSZip, templateBuf, o) {
    const zip = await JSZip.loadAsync(templateBuf);
    let d = await zip.file("word/document.xml").async("string");
    const headRe = /<w:p><w:pPr><w:pStyle w:val="Heading1"\/><\/w:pPr><w:bookmarkStart w:name="bk_([^"]+)"[^>]*\/><w:r><w:t xml:space="preserve">([^<]*)<\/w:t><\/w:r><w:bookmarkEnd[^>]*\/><\/w:p>/g;
    const heads = [...d.matchAll(headRe)];
    if (heads.length !== CODES.length) throw new Error("Templat tidak dijangka (" + heads.length + " kitab, perlu 66)");

    const line = o.bodySize ? Math.round(Math.round(o.bodySize * 2) * 11.667) : 210;          // template: 9 pt, exact 10.5 pt leading
    const popt = { min: o.paraMin == null ? 4 : o.paraMin, max: o.paraMax == null ? 5 : o.paraMax,
                   gap: Math.round(line * (o.chapterGapLines == null ? 1 : o.chapterGapLines)),
                   hyph: o.hyphenate !== false ? makeHyph(o) : null };                                             // Malay hyphenation on by default   // one blank line after each paragraph

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
      const body = chapters((o.verses && o.verses[code]) || [], popt);
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

    // style tweaks, all in styles.xml: paragraph indent, tiny tracking, optional typeface / size / hyphenation language
    {
      let st = await zip.file("word/styles.xml").async("string");
      const q = (f) => String(f).replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
      const indent = o.indent == null ? 240 : o.indent;                 // first-line indent of 2nd+ paragraphs (twips; 240 = 12 pt, about 5 spaces)
      const tracking = o.tracking == null ? -2 : o.tracking;            // character spacing in 1/20 pt; -2 = -0.1 pt, invisible but fits more per line
      st = st.replace(/(w:styleId="Verse"[\s\S]*?<w:ind w:firstLine=")\d+/, "$1" + indent);
      if (tracking) st = st.replace(/(w:styleId="Verse"[\s\S]*?<\/w:pPr>)(<\/w:style>)/, "$1<w:rPr><w:spacing w:val=\"" + tracking + "\"/></w:rPr>$2");
      if (o.font) {
        const num = o.numFont || (/^georgia$/i.test(o.font) ? "Calibri" : o.font);   // Georgia digits are old-style, so keep lining digits for verse numbers
        st = st.split('"Garamond"').join('"' + q(o.font) + '"').split('"Calibri"').join('"' + q(num) + '"');
      }
      if (o.bodySize) {
        const sz = Math.round(o.bodySize * 2);
        st = st.replace(/(<w:rPrDefault>[\s\S]*?<w:sz w:val=")\d+("\/><w:szCs w:val=")\d+/, "$1" + sz + "$2" + sz)
               .replace(/(<w:pPrDefault>[\s\S]*?w:line=")\d+/, "$1" + line)
      }
      st = st.replace(/(w:styleId="ChapterNumber"[\s\S]*?w:line=")\d+/, "$1" + (line * 3 - 20));   // drop cap: just under 3 body lines
      if (o.hyphenLang) st = st.replace(/(<w:rPrDefault>[\s\S]*?<w:szCs w:val="\d+"\/>)(<\/w:rPr>)/, "$1<w:lang w:val=\"" + q(o.hyphenLang) + "\"/>$2");
      zip.file("word/styles.xml", st);
    }

    zip.file("word/document.xml", d);
    return zip.generateAsync({ type: o.type || "blob", compression: "DEFLATE", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  }

  root.fillBibleDocx = fillBibleDocx;
  root.FILL_CODES = CODES;
  if (typeof module !== "undefined") module.exports = { fillBibleDocx, CODES };
})(typeof window !== "undefined" ? window : globalThis);
