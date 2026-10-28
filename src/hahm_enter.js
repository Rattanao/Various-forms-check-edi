  // ───────────── 2) ENTER.pdf ชุดรวม HAHM — พอร์ตจาก scripts/hahm_edi.py ─────────────
  // ENTER ของ HAHM เป็น PDF ชุดเดียวที่รวมฟอร์มหลาย forwarder: แยกชนิดทีละหน้า แล้วอ่านด้วยตัวอ่านของแต่ละชนิด
  // ฟอร์ม MAC-NELS / THE M CORP (CMS) / DSV ส่งเฉพาะหน้านั้นให้ตัวอ่านของเว็บพี่น้อง (macnels.js / cms.js / dsv.js — สำเนาไม่แก้ไข)
  const CONT2_RE = /(?<![A-Z])[A-Z]{3}[UJZ]\d{7}/g;          // ตัวที่ 4 = U/J/Z (ISO 6346) — กัน SHIPMENT ID / ท้ายเลข B/L
  const BL_FULL_RE = /^HASL[A-Z]\d{8,}[A-Z]?$/;
  const BL_ANY_RE = /HASL[A-Z]\d{8,}[A-Z]?/;
  const BL_IN_TXT_RE = /HASL[A-Z]\d{8,11}[A-Z]?(?![A-Za-z0-9])/g;
  const STATUS_KW = /STATUS|ลากตู้|ลากตต|เปิดตู้|ขนส่ง|\bCFS\b/i;
  const PKG_WORD_RE = /^(?:WOODEN\s+)?(PALLETS?|PLTS?|CARTONS?|CTNS?|CASES?|BOX(?:ES)?|BOX\(ES\)|PACKAGES?|PKGS?|SETS?|ROLLS?|REELS?|DRUMS?|BAGS?|BALES?|CRATES?|UNITS?|PIECES?|BUNDLES?|SKIDS?)\b/i;
  const PKG_WORD_FULL = new RegExp(PKG_WORD_RE.source.replace(/\\b$/, "") + "$", "i");
  const MINUS_TEMP_RE = /(?:MINUS|-)\s*(\d{1,2}(?:\.\d)?)\s*(?:DEG|DEGREES|°|'\s*C)/i;

  const nospace = s => String(s || "").replace(/\s+/g, "").toUpperCase();
  const topOf = l => (l.top !== undefined ? l.top : l.y);
  const inAnnot = (l, annots) => annots.some(a => a.x0 - 3 <= l.x && l.x <= a.x1 && a.y0 - 3 <= topOf(l) && topOf(l) <= a.y1 + 1);
  function contsIn(...txts) {
    const seen = [];
    for (const t of txts) for (const c of String(t || "").toUpperCase().match(CONT2_RE) || []) if (!seen.includes(c)) seen.push(c);
    return seen;
  }
  function splitPkg(s) {
    const m = String(s || "").match(/^\s*(\d[\d,]*)\s*([A-Za-z][\s\S]*)?$/);
    return m ? [parseInt(m[1].replace(/,/g, ""), 10), (m[2] || "").trim()] : [null, ""];
  }
  function cleanStatus(parts) {
    const out = [];
    for (let p of parts) {
      // ฟอนต์เพี้ยน: PyMuPDF ได้ 'ลากตต' / pdf.js ได้ 'ลากตตต'
      p = String(p).replace(/ลากตต+/g, "ลากตู้").replace(/^\s*(?:STATUS|REMARKS?)\s*[:=]?\s*/i, "").replace(/^[\s*"“”:]+|[\s*"“”:]+$/g, "");
      if (p && !out.includes(p)) out.push(p);
    }
    return out.join(" ");
  }
  const isStatusLine = t => STATUS_KW.test(t);
  function findReefer(...txts) {
    for (const t of txts) { const m = String(t || "").match(MINUS_TEMP_RE); if (m) return `-${m[1]}°C`; }
    return findTemp(...txts);
  }
  function companyName(lines) {
    lines = lines.filter(l => l.trim());
    if (!lines.length) return "";
    let out = [], hit = false;
    for (const l of lines.slice(0, 3)) {
      out.push(l.trim());
      if (/\b(LTD|LIMITED|CORP(?:ORATION)?|INC|PLC|LLC|CO\.?,?\s*LTD)\b\.?\s*\*?$/i.test(l.trim()) || /\b(LTD|LIMITED)\b/i.test(l)) { hit = true; break; }
    }
    if (!hit) out = [lines[0].trim()];
    let s = out.join(" ").replace(/ D\./g, "D.").replace(/\s+/g, " ").replace(/^[\s*]+|[\s*]+$/g, "");
    return s.replace(/LT D\./g, "LTD.");
  }
  const blank = (bl, page) => ({ bl, cons: "", cont: "", pkgs: null, pkgtype: "", gw: null, meas: null, marks: "", desc: "",
    status_raw: "", transit_raw: "", dg: "", reefer: "", movement_raw: "", low_confidence: false, pages: [page] });
  function finish(rec, ...extra) {
    const blob = [rec.marks, rec.desc, ...extra.filter(Boolean)].join(" ");
    rec.dg = rec.dg || findDg(blob);
    rec.reefer = rec.reefer || findReefer(blob);
    if (!rec.transit_raw && TRANSIT_RE.test(blob)) rec.transit_raw = blob;
    return rec;
  }
  const uniqSort = a => [...new Set(a)].sort((x, y) => x - y);
  const byYX = (a, b) => a.y - b.y || a.x - b.x;

  // ── ชนิดหน้า
  function pageKind(lines) {
    const ns = nospace(lines.map(l => l.t).join(" "));
    if (!ns) return "SCAN";
    if (ns.includes("MAC-NELS") && ns.includes("ATTACHEDSHEETS")) return "SHARED_MACNELS";
    if (ns.includes("INBOUNDCARGOMANIFEST")) return "SHARED_CMS";
    if (ns.includes("AMENDMENTOFMANIFEST") && ns.includes("B/LCHANGED")) return "SHARED_DSV";
    if (ns.includes("ATTACHEDRIDER") && !ns.includes("SHIPPER'SLOAD")) return "RIDER";
    if (ns.includes("MARK&NO.") && ns.includes("CNTR:") && ns.includes("B/LNO.:")) return "MARKNO";
    if (ns.includes("SHIPPER'SLOAD")) return "BLFORM";
    const up = lines.map(l => l.t.toUpperCase());
    if (up.includes("CONSIGNEE") && up.some(u => u.startsWith("QUANTITY")) && up.some(u => u.startsWith("DESCRIPTION")) && up.some(u => u.startsWith("GROSS WEIGHT")))
      return "TABLE";
    if (lines.length <= 3 && !BL_ANY_RE.test(ns)) return "SCAN";
    return "LETTER";
  }

  // ── ARRIVAL NOTICE / สำเนา B/L ของ Heung-A
  const lab = (lines, key) => lines.find(l => nospace(l.t).startsWith(key)) || null;
  function parseBlform(lines, annots, page) {
    const blLine = lines.find(l => BL_FULL_RE.test(nospace(l.t)) && l.x > 330 && l.y < 140);
    if (!blLine) return {};
    const rec = blank(nospace(blLine.t), page);
    rec.src = "BLFORM";
    const body = lines.filter(l => !inAnnot(l, annots));
    const sl = lab(body, "SHIPPER'SLOAD");
    const ySl = sl.y, xSl = sl.x;
    const tots = [lab(body, "(*)TOTALNO"), lab(body, "SAY:")].filter(Boolean).map(l => l.y);
    const yT = tots.length ? Math.min(...tots) : ySl + 230;
    const kgs = body.find(l => Math.abs(l.y - ySl) < 4 && /KGS/.test(l.t));
    const cbm = body.find(l => Math.abs(l.y - ySl) < 4 && /CBM/.test(l.t));
    const pkl = body.find(l => Math.abs(l.y - ySl) < 4 && l.x < xSl - 2 && /^\d/.test(l.t));
    const xGw = (kgs ? kgs.x : 420) - 6;
    rec.gw = kgs ? num(kgs.t) : null;
    rec.meas = cbm ? num(cbm.t) : null;
    if (pkl) [rec.pkgs, rec.pkgtype] = splitPkg(pkl.t);

    const lc = lab(body, "CONSIGNEE"), ln = lab(body, "NOTIFYPARTY"), lp = lab(body, "PRECARRIAGE");
    const left = body.filter(l => l.x < 200 && l.y < ySl - 30 && !BL_ANY_RE.test(nospace(l.t))).sort((a, b) => a.y - b.y);
    let consLines, notiLines;
    if (lc && ln) {
      const yEnd = lp ? lp.y : ln.y + 100;
      consLines = left.filter(l => lc.y + 3 < l.y && l.y < ln.y - 1).map(l => l.t);
      notiLines = left.filter(l => ln.y + 3 < l.y && l.y < yEnd - 1).map(l => l.t);
    } else {
      const blocks = []; let cur = [];
      for (const l of left) { if (cur.length && l.y - cur[cur.length - 1].y > 16) { blocks.push(cur); cur = []; } cur.push(l); }
      if (cur.length) blocks.push(cur);
      consLines = blocks.length > 1 ? blocks[1].map(l => l.t) : [];
      notiLines = blocks.length > 2 ? blocks[2].map(l => l.t) : [];
    }
    let cons = companyName(consLines);
    const notify = companyName(notiLines);
    if (/^TO\s+ORDER/i.test(cons) && notify && !/^SAME\s+AS/i.test(notify)) { rec.cons_note = `ENTER Consignee = ${cons} → ใช้ชื่อ Notify`; cons = notify; }
    rec.cons = cons;
    if (notify && !/^SAME\s+AS/i.test(notify)) rec.notify = notify;

    const cargo = body.filter(l => ySl + 3 < l.y && l.y < yT - 1);
    const contLine = t => contsIn(t).length > 0 && /^(?:R?\d{2}[A-Z]{0,3}'?)?\s*[A-Z]{3}[UJZ]\d{7}/.test(nospace(t));
    const noise = t => /^\d{2}'?[A-Z]*X\d+$/.test(nospace(t)) || nospace(t).startsWith("PARTICULARSFURNISHED")
      || ["SAIDTOCONTAIN", "SHIPPER'SLOAD,COUNT&WEIGHT"].includes(nospace(t));
    const marks = [], desc = [], typed = [];
    for (const l of cargo.slice().sort(byYX)) {
      const t = l.t;
      if (noise(t)) continue;
      if (isStatusLine(t) && t !== "CY/CY") { typed.push(t); continue; }
      if (l.x < xSl - 10) { if (!contLine(t)) marks.push(t); }
      else if (l.x < xGw) desc.push(t);
    }
    rec.marks = marks.join(" ");
    rec.desc = desc.join(" ");
    rec.cont = contsIn(...cargo.filter(l => l.x < xSl - 10 && contLine(l.t)).map(l => l.t), ...annots.map(a => a.t)).join(";");
    for (const l of body) if ((l.y >= yT - 1 || l.y <= ySl + 3) && isStatusLine(l.t) && l.t !== "CY/CY") typed.push(l.t);
    const labelled = applyLabelledAnnots(rec, annots, lc, ln);
    for (const a of annots) {
      if (labelled.has(a.t)) continue;
      typed.push(...a.lines.filter(isStatusLine));
      rec.dg = rec.dg || findDg(a.t);
      rec.reefer = rec.reefer || findReefer(a.t);
      if (TRANSIT_RE.test(a.t)) rec.transit_raw = a.t;
    }
    rec.status_raw = rec.status_raw || cleanStatus(typed);
    return { [rec.bl]: finish(rec) };
  }

  function applyLabelledAnnots(rec, annots, lc, ln) {
    const used = new Set();
    const L = [["pkgs", "(?:QUANTITY\\s*PACKAGE|NO\\.?\\s*OF\\s*PKGS?|PACKAGES?)"], ["desc", "(?:CARGO\\s*DESCRIPTION|DESCRIPTION)"],
      ["marks", "MARKS?\\s*&\\s*NOS?\\.?"], ["gw", "GROSS\\s*WEIGHT"], ["meas", "(?:MEAS(?:UREMENT)?|CBM)"], ["status", "STATUS"]];
    for (const a of annots) {
      const hits = {}, free = [];
      for (const p of a.lines) {
        let hit = false;
        for (const [k, pat] of L) {
          const m = p.match(new RegExp(`^${pat}\\s*[:=]\\s*([\\s\\S]*)$`, "i"));
          if (m) { hits[k] = m[1].trim(); hit = true; break; }
        }
        if (!hit) free.push(p);
      }
      const inConsBox = !!(lc && ln && topOf(lc) - 10 <= a.y0 && a.y0 < topOf(ln) - 2);
      const nh = Object.keys(hits).length;
      if (nh || inConsBox) {
        used.add(a.t);
        if (free.length && (inConsBox || nh >= 2) && !contsIn(free[0]).length) rec.cons = free[0];
        if ("pkgs" in hits) [rec.pkgs, rec.pkgtype] = splitPkg(hits.pkgs);
        for (const k of ["desc", "marks"]) if (k in hits) rec[k] = hits[k];
        if ("gw" in hits) rec.gw = num(hits.gw);
        if ("meas" in hits) rec.meas = num(hits.meas);
        if ("status" in hits) rec.status_raw = cleanStatus([hits.status]);
      }
    }
    return used;
  }

  function parseRider(lines) {
    const bl = (lines.find(l => BL_FULL_RE.test(nospace(l.t))) || { t: "" }).t;
    const mi = lines.find(l => nospace(l.t).includes("MARKINFORMATION"));
    const marks = [], typed = [];
    for (const l of lines.slice().sort(byYX)) {
      if (isStatusLine(l.t)) { typed.push(l.t); continue; }
      if (mi && l.y > mi.y + 3 && l.x < 250) marks.push(l.t);
    }
    return [nospace(bl), marks.join(" "), typed];
  }

  // ── ใบ MARK&NO. / NO. OF PKGS / DESCRIPTION OF GOODS (2 B/L ต่อหน้า)
  function parseMarkno(lines, page) {
    const out = {};
    const lbls = lines.filter(l => nospace(l.t) === "B/LNO.:").sort((a, b) => a.y - b.y);
    const hdr = lines.find(l => nospace(l.t) === "MARK&NO.");
    let prevY = hdr ? hdr.y + 5 : 0;
    for (const lb of lbls) {
      const val = lines.find(l => Math.abs(l.y - lb.y) < 3 && l.x > lb.x + 20 && BL_FULL_RE.test(nospace(l.t)));
      if (!val) continue;
      const rec = blank(nospace(val.t), page);
      const blk = lines.filter(l => prevY < l.y && l.y <= lb.y + 3);
      prevY = lb.y + 3;
      const cl = blk.find(l => nospace(l.t) === "CNTR:");
      const yCntr = cl ? cl.y : lb.y - 20;
      const marks = [], desc = [], mid = [];
      for (const l of blk.slice().sort(byYX)) {
        const t = l.t;
        if (t.startsWith("G.W.")) rec.gw = num(t);
        else if (t.startsWith("MEAS.")) rec.meas = num(t);
        else if ([yCntr, lb.y].some(y => Math.abs(l.y - y) < 3) && l.x < 200) { if (Math.abs(l.y - yCntr) < 3) rec.cont = contsIn(t).join(";") || rec.cont; }
        else if (l.x < 200) marks.push(t);
        else if (l.x < 340) mid.push(l);
        else if (l.y < yCntr) desc.push(t);
      }
      const cnt = mid.find(l => /^[\d,]+$/.test(l.t));
      if (cnt) {
        rec.pkgs = parseInt(cnt.t.replace(/,/g, ""), 10);
        const pt = mid.find(l => l !== cnt && l.y < cnt.y + 30 && !/^[\d,]+$/.test(l.t));
        rec.pkgtype = pt ? pt.t : "";
        rec.cons = mid.filter(l => l !== cnt && l !== pt && l.y < yCntr).map(l => l.t).join(" ");
      }
      rec.marks = marks.join(" ");
      rec.desc = desc.join(" ");
      out[rec.bl] = finish(rec);
    }
    return out;
  }

  // ── ตาราง AMENDMENT (Cargoport / ASIA ATLANTIC / INITIAL / Handle ฯลฯ)
  function statusValue(lines, l) {
    const m = l.t.match(/^\s*STATUS\s*[:=]?\s*([\s\S]*)$/i);
    let v = (m ? m[1] : l.t).replace(/^[\s:]+|[\s:]+$/g, "");
    if (!v) v = lines.filter(o => o !== l && Math.abs(o.y - l.y) < 3 && o.x > l.x && o.t.replace(/[\s:]/g, ""))
      .sort((a, b) => a.x - b.x).map(o => o.t).join(" ");
    return cleanStatus([v]);
  }
  function parseTablePage(lines, annots, page) {
    const hdr = lines.find(l => l.t.toUpperCase() === "CONSIGNEE");
    const hy = hdr.y;
    const col = {};
    for (const l of lines.filter(l => Math.abs(l.y - hy) < 6)) {
      const u = l.t.toUpperCase();
      if (u.startsWith("MARKS")) col.MARKS = l.x;
      else if (u.startsWith("QUANTITY")) col.QTY = l.x;
      else if (u.startsWith("DESCRIPTION")) col.DESC = l.x;
      else if (u === "CONSIGNEE") col.CONS = l.x;
      else if (u.startsWith("GROSS WEIGHT")) col.GW = l.x;
      else if (u.startsWith("HB/L")) col.HBL = l.x;
    }
    const bounds = Object.entries(col).map(([k, x]) => [x - 45, k]).sort((a, b) => a[0] - b[0]);
    const colOf = x => { let n = "BL"; for (const [lft, k] of bounds) if (x >= lft) n = k; return n; };
    const ends = lines.filter(l => (l.y > hy + 5 && (l.t.toUpperCase() === "TOTAL" || l.t.toUpperCase().startsWith("TOTAL ")))
      || /^Page \d+ of \d+/.test(l.t) || l.t.startsWith("[FM-")).map(l => l.y);
    const endY = Math.min(...ends, 1e9);
    const anchors = [], used = new Set();
    for (const l of lines.slice().sort((a, b) => a.y - b.y)) {
      if (l.y > hy + 5 && l.y < endY && colOf(l.x) === "BL" && /^HASL[A-Z0-9]+$/.test(l.t)) {
        let bl = l.t;
        const suf = lines.find(o => o !== l && Math.abs(o.x - l.x) < 5 && o.y - l.y > 3 && o.y - l.y < 16 && /^[A-Z0-9]{1,2}$/.test(o.t));
        if (suf) { bl += suf.t; used.add(suf); }
        used.add(l);
        anchors.push([l.y, bl]);
      }
    }
    const rows = {};
    anchors.forEach(([ay, bl], i) => {
      const ny = i + 1 < anchors.length ? anchors[i + 1][0] : endY;
      const rl = lines.filter(l => ay - 4 <= l.y && l.y < ny - 4 && !used.has(l));
      const rec = blank(bl, page);
      const parts = { BL: [], MARKS: [], QTY: [], DESC: [], CONS: [], GW: [], HBL: [] };
      const st = [];
      for (const l of rl.slice().sort(byYX)) {
        if (/^\s*STATUS\b/i.test(l.t)) { st.push(statusValue(rl, l)); continue; }
        parts[colOf(l.x)].push(l.t);
      }
      rec.marks = parts.MARKS.join(" "); rec.desc = parts.DESC.join(" "); rec.cons = parts.CONS.join(" ");
      [rec.pkgs, rec.pkgtype] = splitPkg(parts.QTY.join(" "));
      const gwt = parts.GW.join(" ");
      const mg = gwt.match(/G\.?\s*W\.?\s*([\d,]+\.?\d*)/) || gwt.match(/([\d,]+\.\d+)\s*KGS/);
      const mm = gwt.match(/([\d,]+\.?\d*)\s*(?:M3|CBM)/);
      let mmV = null;
      if (!mm) { const ns = gwt.match(/[\d,]+\.\d+/g) || []; mmV = ns.length > 1 ? num(ns[1]) : null; }
      rec.gw = mg ? num(mg[1]) : null;
      rec.meas = mm ? num(mm[1]) : mmV;
      rec.cont = contsIn(...parts.BL, ...parts.MARKS, ...parts.DESC).join(";");
      rec.status_raw = st.join(" ");
      rows[bl] = rec;
    });
    const firstY = anchors.length ? anchors[0][0] : null;
    const pageStatus = lines.filter(l => /^\s*STATUS\b/i.test(l.t) && !(firstY !== null && firstY - 4 <= l.y && l.y < endY)).map(l => statusValue(lines, l));
    const pageConts = contsIn(...lines.filter(l => l.y < hy).map(l => l.t), ...annots.map(a => a.t));
    for (const a of annots) if (isStatusLine(a.t)) pageStatus.push(cleanStatus([a.t]));
    return [rows, pageStatus.filter(Boolean), pageConts];
  }

  // ── จดหมายแก้ไขแบบข้อความอิสระ
  const LETTER_STOP = /^(CONSIGNEE|CNEE|CONTAINER|CONT\.?\s*NO|TOTAL|-{5,}|WE |YOUR|THANK|STATUS|REMARK|NOTIFY|PLACE OF|VESSEL|จ่าย)/i;
  const L_CONS = "(?:CONSIGNEE\\s*NAME|CONSIGNEE|CNEE)", L_PKG = "(?:PACKAGES?|NO\\.\\s*OF\\s*PCS|QUANTITY\\s*PACKAGE)";
  const L_GW = "(?:WEIGHT\\s*&\\s*CBM|GROSS\\s*(?:WEIGHT|WT\\.?))", L_MEAS = "(?:GROSS\\s*CBM|CBM\\.?)";
  const L_MARKS = "(?:MARKS?\\s*(?:AND|&)\\s*(?:NUMBERS?|NOS?\\.?|NO\\.?))", L_DESC = "(?:CARGO\\s*DESCRIPTION|DESCRIPTION)";
  const strip = (s, ch) => s.replace(new RegExp(`^[${ch}]+|[${ch}]+$`, "g"), "");
  function labelValue(lines, pat) {
    for (const l of lines.slice().sort((a, b) => a.y - b.y)) {
      const m = l.t.match(new RegExp(`^${pat}\\s*[:=]?\\s*([\\s\\S]*)$`, "i"));
      if (!m) continue;
      let v = strip(m[1], " :=");
      if (!v) v = strip(lines.filter(o => o !== l && Math.abs(o.y - l.y) < 6 && o.x > l.x).sort((a, b) => a.x - b.x).map(o => o.t).join(" "), " :=");
      if (!v) {
        const near = lines.filter(o => o.x > l.x + 40 && o.t.startsWith(":") && Math.abs(o.y - l.y) < 22).sort((a, b) => Math.abs(a.y - l.y) - Math.abs(b.y - l.y));
        if (near.length) v = strip(near[0].t, " :=");
      }
      if (v) return [v, l];
    }
    return ["", null];
  }
  function parseLetter(lines, annots, page) {
    const body = lines.filter(l => !inAnnot(l, annots));
    const txt = body.slice().sort(byYX).map(l => l.t).join(" ");
    const bls = txt.match(BL_IN_TXT_RE);
    if (!bls) return {};
    const rec = blank(bls[0], page);
    rec.cons = labelValue(body, L_CONS)[0];
    rec.as_per_attached = /AS\s*PER\s*ATTACHED/i.test(rec.cons);
    rec.cont = contsIn(txt, ...annots.map(a => a.t)).join(";");
    const reM = new RegExp(L_MARKS, "i");
    const hm = body.slice().sort(byYX).find(l => reM.test(l.t) && l.t.search(reM) === 0 && !l.t.includes(":")
      && body.some(o => /DESCRIP/i.test(o.t) && !o.t.includes(":") && Math.abs(o.y - l.y) < 8));
    if (hm) {
      const hrow = body.filter(o => Math.abs(o.y - hm.y) < 10);
      const hx = {};
      for (const o of hrow) {
        const u = o.t.toUpperCase(), isM = new RegExp(`^${L_MARKS}`).test(u);
        if (isM && !("MARKS" in hx)) hx.MARKS = o.x;
        if (/NUMBER$|NO\.?\s*OF\s*P|PKGS|PACKAGES/.test(u) && !isM && !("PKG" in hx)) hx.PKG = o.x;
        if (u.includes("DESCRIP") && !isM && !("DESC" in hx)) hx.DESC = o.x;
        if (u.includes("WEIGHT") && !isM && !("GW" in hx)) hx.GW = o.x;
      }
      const single = hrow.length === 1;
      const lastY = Math.max(...hrow.map(o => o.y));
      const rows = [];
      for (const l of body.filter(o => o.y > lastY + 3).sort(byYX)) {
        if (rows.length && l.y - Math.max(...rows.map(r => r.y)) > 45) break;
        if (LETTER_STOP.test(l.t) || isStatusLine(l.t)) { if (rows.length) break; continue; }
        rows.push(l);
      }
      const marks = [], desc = [], gwm = [], numsOnly = [];
      const pkgL = rows.find(r => PKG_WORD_RE.test(splitPkg(r.t)[1] || "") || (/^\d[\d,]*$/.test(r.t) && "PKG" in hx && Math.abs(r.x - hx.PKG) < 40));
      if (pkgL) {
        [rec.pkgs, rec.pkgtype] = splitPkg(pkgL.t);
        if (!rec.pkgtype) { const pw = rows.find(r => PKG_WORD_FULL.test(r.t.trim()) && Math.abs(r.x - pkgL.x) < 40); rec.pkgtype = pw ? pw.t : ""; }
      }
      const pkgX = pkgL ? pkgL.x : ("PKG" in hx ? hx.PKG : 200), descX = hx.DESC, gwX = hx.GW;
      const UNIT = /^(KGS?|CBM|M3)\.?$/i;
      for (const r of rows) {
        const t = r.t;
        if (r === pkgL || (PKG_WORD_FULL.test(t.trim()) && Math.abs(r.x - pkgX) < 40)) continue;
        if (/^[\d,]+\.\d+$/.test(t)) numsOnly.push(r);
        else if (UNIT.test(t.trim())) continue;
        else if (/KGS?\b|CBM|M3\b/i.test(t) || (gwX !== undefined && r.x >= gwX - 30)) gwm.push(t);
        else if (r.x < pkgX - 5) marks.push(t);
        else if (single || descX === undefined || r.x >= (pkgX + descX) / 2) desc.push(t);
      }
      rec.marks = marks.join(" "); rec.desc = desc.join(" ");
      for (const r of numsOnly) {
        const unit = rows.find(u => UNIT.test(u.t.trim()) && u.y - r.y > 0 && u.y - r.y < 30 && Math.abs(u.x - r.x) < 25);
        gwm.push(`${r.t} ${unit ? unit.t : ""}`.trim());
      }
      const g = gwm.join(" ");
      const mg = g.match(/([\d,]+\.?\d*)\s*KGS?/i);
      const mm = g.match(/([\d,]+\.?\d*)\s*(?:CBM|M3)/i) || g.match(/CBM\.?\s*=?\s*([\d,]+\.?\d*)/i);
      const ns = g.match(/\d[\d,]*\.\d+/g) || [];
      rec.gw = mg ? num(mg[1]) : (ns.length ? num(ns[0]) : null);
      rec.meas = mm ? num(mm[1]) : (ns.length > 1 ? num(ns[1]) : null);
      if (rec.gw === null) { const v = labelValue(body, L_GW)[0]; rec.gw = v ? num(v) : null; }
      if (rec.meas === null) { const v = labelValue(body, L_MEAS)[0]; rec.meas = v ? num(v) : null; }
    } else {
      let v = labelValue(body, L_PKG)[0];
      if (v) [rec.pkgs, rec.pkgtype] = splitPkg(v);
      v = labelValue(body, L_GW)[0];
      if (v) { rec.gw = num(v); const mm = v.match(/([\d,]+\.?\d*)\s*(?:M3|CBM)/i); if (mm) rec.meas = num(mm[1]); }
      if (rec.meas === null) { v = labelValue(body, L_MEAS)[0]; rec.meas = v ? num(v) : null; }
      for (const [key, pat] of [["marks", L_MARKS], ["desc", L_DESC]]) {
        const [val, l] = labelValue(body, pat);
        if (!l) continue;
        const vx = l.x + 60, more = [];
        for (const o of body.filter(o => l.y + 5 < o.y && o.y < l.y + 60 && o.x >= vx).sort((a, b) => a.y - b.y)) {
          if (LETTER_STOP.test(o.t) || /^[A-Z .&'/-]+\s*:/.test(o.t) || o.t.startsWith(":") || body.some(p => p.x < vx && Math.abs(p.y - o.y) < 4)) break;
          more.push(o.t);
        }
        rec[key] = [val, ...more].join(" ").trim();
      }
    }
    const typed = body.slice().sort(byYX).filter(l => isStatusLine(l.t) && !new RegExp(`^${L_DESC}`, "i").test(l.t)).map(l => l.t);
    rec.status_raw = cleanStatus([...typed, ...annots.flatMap(a => a.lines.filter(isStatusLine))]);
    rec.combined_text = true;
    for (const a of annots) { rec.dg = rec.dg || findDg(a.t); rec.reefer = rec.reefer || findReefer(a.t); }
    return { [rec.bl]: finish(rec, ...annots.map(a => a.t)) };
  }

  // ── ฟอร์มของเว็บพี่น้อง: ส่งเฉพาะหน้านั้นให้ตัวอ่านเดิม
  const SHARED = { SHARED_MACNELS: "CheckMacNels", SHARED_CMS: "CheckEastern", SHARED_DSV: "CheckDSV" };
  function parseShared(kind, pages, pnos, blRe) {
    const lib = root[SHARED[kind]];
    if (!lib) throw new Error(`ไม่พบตัวอ่าน ${SHARED[kind]} (โหลด ${kind.slice(7).toLowerCase()}.js ไม่สำเร็จ)`);
    const sub = pnos.map(p => {
      const pg = pages[p - 1];
      const items = pg.items.map(it => ({ ...it }));
      if (kind === "SHARED_CMS")   // ตัวอ่าน CMS นับข้อความที่ลูกค้าพิมพ์ทับเป็นเนื้อหาตามตำแหน่ง (เหมือนเว็บ CMS)
        for (const a of pg.annots) a.lines.forEach((t, k) => items.push({ x: a.x0 + 2, y: a.y0 + 10 + k * 12, w: 0, t }));
      return { items, annots: pg.annots.map(a => ({ y: a.y0, text: a.text })) };
    });
    const { ENT } = lib.parseEnter(sub, blRe);
    const pageBls = {};
    for (const p of pnos) for (const bl of pages[p - 1].text.match(BL_IN_TXT_RE) || []) (pageBls[bl] = pageBls[bl] || []).push(p);
    const out = {};
    for (const [bl, rec] of Object.entries(ENT || {})) if (pageBls[bl]) { rec.pages = uniqSort(pageBls[bl]); out[bl] = rec; }
    return out;
  }
  function fixDsv(recs, pages) {
    for (const r of Object.values(recs)) {
      const cs = contsIn(r.pages.map(p => pages[p - 1].text).join(" "));
      if (cs.length) r.cont = cs.join(";");
      r.desc = String(r.desc || "").split(/(?<=\))\s+|\s{2,}/).filter(w => !/^[A-Z]{3}[UJZ]\d{7}\s*\(?[^)]*\)?$/.test(w.trim())).join(" ");
      r.desc = r.desc.replace(/\s*[A-Z]{3}[UJZ]\d{7}\s*\(\d{2}[A-Z]{0,2}\)/g, "").trim();
    }
  }

  // pages: [{items:[{x,y,w,t}], annots:[{x0,y0,x1,y1,text,lines,t}], text}]
  function parseEnter(pages, blRe) {
    const L = pages.map(pg => pg.lines.filter(i => i.t));   // บรรทัดแบบ PyMuPDF (ตัวอ่านเว็บพี่น้องใช้ pg.items เดิม)
    const A = pages.map(pg => pg.annots);
    const kinds = L.map(pageKind);
    const ENT = {}, LETTERS = {}, scans = [];
    const put = recs => {
      for (const [bl, rec] of Object.entries(recs)) {
        if (ENT[bl]) {
          ENT[bl].pages = uniqSort([...ENT[bl].pages, ...rec.pages]);
          for (const [k, v] of Object.entries(rec)) if (k !== "pages" && v !== null && v !== "" && (ENT[bl][k] === null || ENT[bl][k] === "" || ENT[bl][k] === undefined)) ENT[bl][k] = v;
        } else ENT[bl] = rec;
      }
    };
    const groups = [];
    for (let i = 0; i < kinds.length;) {
      let j = i;
      while (j + 1 < kinds.length && kinds[j + 1] === kinds[i] && kinds[i].startsWith("SHARED")) j++;
      groups.push([kinds[i], Array.from({ length: j - i + 1 }, (_, k) => i + 1 + k)]);
      i = j + 1;
    }
    let tableDoc = null;
    const tableDocs = [];
    for (const [kind, pnos] of groups) {
      const p = pnos[0], ls = L[p - 1], an = A[p - 1];
      if (kind.startsWith("SHARED")) {
        const recs = parseShared(kind, pages, pnos, blRe);
        if (kind === "SHARED_DSV") fixDsv(recs, pages);
        const blYs = ls.filter(l => BL_ANY_RE.test(nospace(l.t))).map(topOf);
        const top = blYs.length ? Math.min(...blYs) : 0;
        const headSt = cleanStatus(an.filter(a => a.y0 < top).flatMap(a => a.lines.filter(isStatusLine)));
        if (headSt) for (const r of Object.values(recs)) if (!r.status_raw) { r.status_raw = headSt; r.pages = uniqSort([...r.pages, p]); }
        if (kind === "SHARED_CMS") for (const r of Object.values(recs)) {
          const q = Math.min(...r.pages), ql = L[q - 1];
          const d = ql.find(l => nospace(l.t).includes(r.bl));
          if (q > pnos[0] && d && !ql.some(l => l.t.includes("KGS") && l.y < d.y)) r.pages = uniqSort([...r.pages, q - 1]);
        }
        put(recs);
        continue;
      }
      if (kind !== "TABLE") tableDoc = null;
      if (kind === "BLFORM") put(parseBlform(ls, an, p));
      else if (kind === "RIDER") {
        const [bl, marks, typed] = parseRider(ls);
        if (bl) {
          const r = ENT[bl] || (ENT[bl] = blank(bl, p));
          r.pages = uniqSort([...r.pages, p]);
          if (marks) r.rider_marks = marks;
          if (typed.length && !r.status_raw) r.status_raw = cleanStatus(typed);
        }
      } else if (kind === "MARKNO") put(parseMarkno(ls, p));
      else if (kind === "TABLE") {
        const isHead = ls.some(l => l.t.toUpperCase().includes("WE WOULD LIKE TO REQUEST") || l.t.toUpperCase() === "AMENDMENT");
        if (isHead || !tableDoc) { tableDoc = { rows: {}, status: [], conts: [], stPages: [], ctPages: [] }; tableDocs.push(tableDoc); }
        const [rows, st, cs] = parseTablePage(ls, an, p);
        Object.assign(tableDoc.rows, rows);
        tableDoc.status.push(...st);
        for (const c of cs) if (!tableDoc.conts.includes(c)) tableDoc.conts.push(c);
        if (st.length) tableDoc.stPages.push(p);
        if (cs.length) tableDoc.ctPages.push(p);
      } else if (kind === "LETTER") Object.assign(LETTERS, parseLetter(ls, an, p));
      else if (kind === "SCAN") scans.push({ page: p, conts: contsIn(...an.map(a => a.t)) });
    }
    for (const td of tableDocs) for (const [bl, rec] of Object.entries(td.rows)) {
      if (!rec.status_raw && td.status.length) { rec.status_raw = [...new Set(td.status)].join(" "); rec.pages = uniqSort([...rec.pages, ...td.stPages]); }
      if (!rec.cont && td.conts.length) { rec.cont = td.conts.join(";"); rec.pages = uniqSort([...rec.pages, ...td.ctPages]); }
      put({ [bl]: finish(rec) });
    }
    for (const rec of Object.values(ENT)) {
      const rm = rec.rider_marks || "";
      delete rec.rider_marks;
      if (!rm) continue;
      if (/AS\s*PER\s*ATTACHED\s*RIDER/i.test(rec.marks)) rec.marks = rec.marks.replace(/AS\s*PER\s*ATTACHED\s*RIDER/gi, rm).trim();
      else if (!rec.marks) rec.marks = rm;
    }
    for (const [bl, lr] of Object.entries(LETTERS)) {
      const esc_ = bl.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const sibs = Object.keys(ENT).filter(b => b === bl || (lr.as_per_attached && new RegExp(`^${esc_}[A-Z]?$`).test(b)));
      if (sibs.length) {
        for (const b of sibs) {
          const e = ENT[b];
          e.pages = uniqSort([...e.pages, ...lr.pages]);
          if (!e.status_raw && lr.status_raw) e.status_raw = lr.status_raw;
          if (b === bl) for (const k of ["cont", "dg", "reefer", "transit_raw"]) if (!e[k] && lr[k]) e[k] = lr[k];
        }
      } else { delete lr.as_per_attached; ENT[bl] = lr; }
    }
    for (const r of Object.values(ENT)) delete r.as_per_attached;
    return { ENT, declared: {}, info: { kinds, scans } };
  }

  // ── กฎเฉพาะ HAHM ก่อนเทียบ
  const ASPER_RE = new RegExp("\\*?\\s*" + "ASPERATTACHED".split("").join("\\s*") + "\\s*(?:" + "SHEET".split("").join("\\s*") + "|"
    + "RIDER".split("").join("\\s*") + ")?\\s*\\*?", "gi");
  function subseqExtra(mTxt, eTxt) {
    const mClean = String(mTxt || "").toUpperCase().replace(ASPER_RE, " ");
    const mCore = mClean.replace(/[^A-Z0-9]/g, "");
    const eWords = String(eTxt || "").split(/\s+/).filter(Boolean);
    if (!mCore || !eWords.length) return null;
    const eCore = String(eTxt).toUpperCase().replace(/[^A-Z0-9]/g, "");
    if (!eCore.includes(mCore)) {
      const eTok = String(eTxt).toUpperCase().match(/[A-Z0-9]+/g) || [];
      let k = 0;
      for (const w of mClean.match(/[A-Z0-9]+/g) || []) {
        while (k < eTok.length && eTok[k] !== w) k++;
        if (k >= eTok.length) return null;
        k++;
      }
    }
    const extra = [];
    let rest = mCore;
    for (const w of eWords) {
      const c = w.toUpperCase().replace(/[^A-Z0-9]/g, "");
      if (c && rest.startsWith(c)) rest = rest.slice(c.length);
      else if (c) extra.push(w);
    }
    return extra.join(" ");
  }
  function applyHahmRules(MAN, ENT) {
    for (const m of Object.values(MAN)) {
      if ((m.reefer && !m.reefer.includes("°C")) || (!m.reefer && MINUS_TEMP_RE.test(m.desc || "")))
        m.reefer = findReefer(m.reefer || "", m.desc || "") || m.reefer || "";
    }
    for (const [bl, e] of Object.entries(ENT)) {
      if (/\bREELS?\b/i.test(e.pkgtype || "")) e.pkgtype = "REEL (ROLL)";
      const m = MAN[bl];
      if (!m || e.src !== "BLFORM") continue;
      for (const k of ["marks", "desc"]) {
        if (sameDetail(e[k], m[k]) === false) {
          const extra = subseqExtra(m[k], e[k]);
          if (extra !== null) { e[k + "_full"] = e[k]; e[k + "_extra"] = extra; e[k] = m[k]; }
        }
      }
    }
  }

  // ── เลขหน้า ENTER / ข้อความเต็ม / หน้าสแกน ลงรายงาน (เหมือน annotate_report())
  function annotateReport(report, MAN, ENT, info) {
    const scanPages = info.scans.map(s => s.page);
    const byCont = {};
    for (const s of info.scans) for (const c of s.conts) for (const [bl, m] of Object.entries(MAN)) if (m.cont.includes(c)) (byCont[bl] = byCont[bl] || []).push(s.page);
    for (const row of report.rows) {
      const bl = row.bl, e = ENT[bl], note = row.cells.note;
      let prefix = "";
      if (e && e.pages && e.pages.length) {
        prefix = `[ENTER หน้า ${e.pages.join(", ")}] `;
        if (e.cons_note) prefix += e.cons_note + " ; ";
        for (const [k, key, lbl] of [["marks", "mkE", "MARKS"], ["desc", "dsE", "DESC."]]) {
          if (!e[k + "_full"]) continue;
          row.cells[key].v = e[k + "_full"];
          if (e[k + "_extra"]) prefix += `${lbl} ตรง (MANIFEST เป็นส่วนหนึ่งของ B/L) — ENTER มีเพิ่ม: ${e[k + "_extra"]} ; `;
        }
      } else if (MAN[bl] && !e) {
        let msg = "";
        if (byCont[bl]) msg = `ENTER น่าจะเป็นหน้า ${byCont[bl].join(", ")} (ภาพสแกน อ่านอัตโนมัติไม่ได้ — เลขตู้ที่พิมพ์ทับตรงกับ MANIFEST) — ตรวจด้วยตา`;
        else if (scanPages.length) msg = `ไม่พบในข้อความ ENTER — มีหน้าภาพสแกนที่อ่านอัตโนมัติไม่ได้: หน้า ${scanPages.join(", ")} — ตรวจด้วยตา`;
        if (msg) {
          prefix = `[${msg}] `;
          for (const k of KEYS) row.cells[k].s = "warn";
          if (byCont[bl]) row.cells.urgent.v = `${WARN} ENTER เป็นภาพสแกน — ตรวจด้วยตา`;
        }
      }
      if (prefix) note.v = prefix + (note.v || "");
    }
    return report;
  }

  // SHED ตาม STATUS: LCL / LCL/CFS → shedLcl, อื่น ๆ → shedOther (STATUS = คำสั่งใน ENTER ถ้ามี ไม่มีจึงใช้ MANIFEST)
  function shedByStatus(shedLcl, shedOther) {
    return (m, e) => {
      const st = canonStatus((e || {}).status_raw || "") || canonStatus(m.status || "");
      return st === "LCL" || st === "LCL/CFS" ? shedLcl : shedOther;
    };
  }

