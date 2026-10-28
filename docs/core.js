/* CHECK EDI VARIOUS FORMS (HAHM) — ตรวจ MANIFEST (.xls) เทียบ ENTER (.pdf ชุดรวมหลาย forwarder ของ HAHM)
   พอร์ตจาก scripts/hahm_edi.py ทำงานในเบราว์เซอร์ล้วน ๆ ไม่มีการอัปโหลด
   ไฟล์นี้สร้างด้วย src/assemble.py: ตัวช่วยร่วม/parseManifest/buildReport มาจากเว็บ Check EDI Freight Links
   (github.com/Rattanao/Freight-links) + ตัวอ่าน ENTER ชุดรวม HAHM (src/hahm_enter.js) — แก้ที่ src/ แล้วรัน assemble ใหม่ */
(function (root) {
  "use strict";

  // ───────────── regex / ตัวช่วยร่วม ─────────────
  const CONT_RE = /[A-Z]{4}\d{6,7}/;
  const TRANSIT_RE = /IN[\s-]?TRANSIT|TRANS[\s-]?S?HIP(?:MENT)?|TRANSIT\s*PORT/i;
  const DG_RE = /\bCLASS\s*[0-9]|\bUN\s*[- ]?\d{3,4}\b/i;
  const REEFER_RE = /REEFER|TEMP|อุณหภูมิ/i;
  const WARN = "⚠", OK = "✔";

  function findDg(...txts) {
    for (const t of txts) {
      if (t && DG_RE.test(t)) {
        const cl = t.match(/CLASS\s*([0-9](?:\.[0-9])?)/i), un = t.match(/UN\s*[- ]?\s*(\d{3,4})/i);
        const s = [cl ? `CLASS ${cl[1]}` : "", un ? `UN ${un[1]}` : ""].filter(Boolean).join(" ");
        return s || t.trim();
      }
    }
    return "";
  }
  function findTemp(...txts) {
    for (const t of txts) {
      if (!t) continue;
      if (REEFER_RE.test(t)) {
        const m = t.match(/(-?\d{1,2}(?:\.\d)?)\s*['°]?\s*(?:deg\.?)?\s*C\b/i);
        return m ? m[1].replace(/^\+/, "") + "°C" : t.replace(/\s+/g, " ").trim();
      }
    }
    return "";
  }
  function extractTransitPhrase(txt) {
    if (!txt) return "";
    txt = txt.replace(/\bLAO\s*P\.?\s*D\.?\s*R\.?/gi, "LAO PDR");   // the dots in P.D.R. are not a sentence end
    // prefer TRANSHIPMENT when a sentence has both words — matches the 3-TRANSHIPMENT rule
    const m = txt.match(/(TRANS[\s-]?S?HIP(?:MENT)?)/i) || txt.match(/(IN\s*TRANSIT)/i);
    if (!m) return txt.replace(/\s+/g, " ").trim();
    const kw = m[1].toUpperCase().replace(/ /g, "").includes("TRANSIT") ? "INTRANSIT" : "TRANSHIPMENT";
    const rest = txt.slice(m.index + m[0].length);
    const tos = [...rest.matchAll(/\bTO\s+([A-Za-z][A-Za-z,.\s]*?)(?=\s+(?:VIA|BY|FROM)\b|["\.]|$)/gi)];
    if (tos.length) return `${kw} TO ${tos[tos.length - 1][1].replace(/\s+/g, " ").trim().replace(/^[ ,.]+|[ ,.]+$/g, "")}`;
    return kw;
  }

  // ───────────── 4 กฎเพิ่มเติม: SHED / ประเทศปลายทาง / CARGO MOVEMENT / TAX ID ─────────────
  const DEST_CODE = { "LAOS": "LA", "LAO PDR": "LA", "LAO": "LA", "MYANMAR": "MM", "CHINA": "CN",
    "MALAYSIA": "MY", "PHILIPPINES": "PH", "CAMBODIA": "KH", "MARSHALL ISLAND": "MH", "MARSHALL ISLANDS": "MH", "INDIA": "IN" };
  const LAOS_NAMES = ["LAOS", "LAO PDR", "LAO"];

  function findDestCountry(txt) {
    if (!txt) return "";
    const u = txt.toUpperCase();
    for (const name of Object.keys(DEST_CODE).sort((a, b) => b.length - a.length)) if (u.includes(name)) return name;
    return "";
  }
  function checkShed(portDischarge, hasDg, hasUsedEngine, destCountry, shedRaw, expectedShed) {
    const pd = (portDischarge || "").toUpperCase();
    const m = (shedRaw || "").match(/\d{3,4}/);
    const actual = m ? m[0].padStart(4, "0") : "";
    let expected = "", label = "";
    if (expectedShed && !LAOS_NAMES.includes(destCountry)) return [actual === expectedShed, expectedShed, "SHED ที่แจ้ง"];
    if (pd.includes("THLKR")) [expected, label] = ["0332", "THLKR"];
    else if (pd.includes("BMT")) [expected, label] = ["0110", "BMT"];
    else if (pd.includes("SCT")) [expected, label] = ["0302", "SCT"];
    else if (pd.includes("UNITHAI")) [expected, label] = ["0113", "UNITHAI"];
    else if (pd.includes("LAEM CHABANG") && hasDg) [expected, label] = ["2826", "DG ที่ DISCHARGE LAEM CHABANG"];
    else if (pd.includes("BANGKOK") && LAOS_NAMES.includes(destCountry)) [expected, label] = ["0124", "LAOS ที่ DISCHARGE BANGKOK"];
    else if (pd.includes("BANGKOK") && hasUsedEngine) [expected, label] = ["0126", "USED ENGINE ที่ DISCHARGE BANGKOK"];
    if (!expected) return [null, "", ""];
    return [actual === expected, expected, label];
  }
  function checkDestCode(transitText, destCodeRaw) {
    const country = findDestCountry(transitText);
    if (!country) return [null, "", ""];
    const expected = DEST_CODE[country] || "";
    if (!expected) return [null, "", ""];
    const m = (destCodeRaw || "").match(/\(([A-Za-z]{2})\)/);
    const actual = m ? m[1].toUpperCase() : "";
    return [actual === expected, expected, actual];
  }
  // TRANSHIPMENT (in MANIFEST or ENTER) must always be 3-TRANSHIPMENT, any destination incl. Laos — checked before the INTRANSIT rules
  const TRANSHIP_RE = /TRANS[\s-]?S?HIP/i;
  function checkMovement(transitText, destCountry, mvManifest, mvEnter, enterTransit) {
    const mv = (mvManifest || "").trim().toUpperCase();
    if (TRANSHIP_RE.test(transitText || "") || TRANSHIP_RE.test(enterTransit || "")) return [mv.startsWith("3"), "3-TRANSHIPMENT", mv];
    if (!transitText) return [null, "", ""];
    if (LAOS_NAMES.includes(destCountry)) {
      const exp = (mvEnter || "").trim().toUpperCase();
      if (!exp) return ["review", "", mv];
      return [(exp === "7" && mv.startsWith("7")) || (exp.startsWith("L") && mv.startsWith("L")), `ตาม ENTER = ${exp}`, mv];
    }
    return [mv.startsWith("7"), "7-TRANSIT", mv];
  }
  function checkTaxid(taxIdRaw, notifyName, enterNotify) {
    const m = (taxIdRaw || "").match(/\(([^)]*)\)/);
    const name = (m ? m[1] : "").trim();
    if (!name) return [null, "", notifyName || ""];
    const a = normCo(name), b = normCo(notifyName || "");
    let ok = !!(a && b && comatch(a, b));
    if (!ok && enterNotify) { const c = normCo(enterNotify); if (a && c && comatch(a, c)) ok = true; }
    return [ok, name, notifyName || ""];
  }
  function computeExtraChecks(m, e, expectedShed) {
    const hasUsedEngine = (m.desc || "").toUpperCase().includes("USED ENGINE");
    const destCountry = findDestCountry(m.transit);
    const out = {};
    let [ok, exp, lbl] = checkShed(m.port_discharge || "", !!m.dg, hasUsedEngine, destCountry, m.shed_no || "", expectedShed);
    const raw = m.shed_no || "";
    if (ok === null) out.shed = { ok: null, text: "-", note: null };
    else if (ok && lbl === "SHED ที่แจ้ง") out.shed = { ok: true, text: "-", note: null };
    else {
      const txt = raw + (ok ? "" : lbl === "SHED ที่แจ้ง" ? `  (ต้องเป็น ${exp})` : `  (ต้องเป็น ${exp} — ${lbl})`);
      out.shed = { ok, text: txt, note: ok ? null : `SHED NO. ไม่ตรง — ${lbl} ต้องเป็น ${exp} (MANIFEST: ${raw || "(ว่าง)"})` };
    }
    let act;
    [ok, exp, act] = checkDestCode(m.transit, m.dest_code_raw || "");
    if (ok === null) out.dest = { ok: null, text: "-", note: null };
    else {
      const disp = act ? `(${act})` : "(ว่าง)";
      out.dest = { ok, text: ok ? "-" : disp + `  (ต้องเป็น (${exp}) — ${destCountry})`,
        note: ok ? null : `ประเทศปลายทางไม่ตรง — ${destCountry} ต้องเป็น (${exp}) (MANIFEST: ${disp})` };
    }
    [ok, exp, act] = checkMovement(m.transit, destCountry, m.movement_raw || "", (e || {}).movement_raw || "", (e || {}).transit_raw || "");
    const rawMv = m.movement_raw || "";
    if (ok === null) out.mv = { ok: null, text: "-", note: null };
    else if (ok === "review") out.mv = { ok: "review", text: (rawMv || "(ว่าง)") + "  (ต้องเทียบ ENTER)",
      note: "CARGO MOVEMENT ไปลาว — ไม่พบค่า CARGO MOVEMENT ใน ENTER ของ B/L นี้ ต้องตรวจสอบด้วยคน" };
    else out.mv = { ok, text: ok ? "-" : (rawMv || "(ว่าง)") + `  (ต้องเป็น ${exp})`,
      note: ok ? null : `CARGO MOVEMENT ไม่ตรง — ต้องเป็น ${exp} (MANIFEST: ${rawMv || "(ว่าง)"})` };
    const enterNotify = m.transit ? (e || {}).notify : null;
    let name, notify;
    [ok, name, notify] = checkTaxid(m.tax_id_raw || "", m.notify, enterNotify);
    if (ok === null) out.tax = { ok: null, text: "-", note: null };
    else out.tax = { ok, text: name + (ok ? "" : `  (N:/NOTIFY = ${notify || "(ว่าง)"})`),
      note: ok ? null : `TAX ID ไม่ตรงกับ N:/NOTIFY PARTY — TAX ID: ${name} / N: ${notify || "(ว่าง)"}` + (enterNotify ? ` / ENTER NOTIFY: ${enterNotify}` : "") };
    return out;
  }

  // ───────────── ตัวช่วยเทียบข้อมูล ─────────────
  const PKG_KIND = { CS: "CASE", PX: "PALLET", PL: "PALLET", CT: "CARTON", CTN: "CARTON", BX: "BOX", PK: "PACKAGE", PKG: "PACKAGE",
    SX: "SET", DR: "DRUM", RO: "ROLL", RL: "ROLL", BG: "BAG", BE: "BALE", BL: "BALE", CR: "CRATE", UN: "UNIT", PC: "PIECE" };
  function canonPkg(s) {
    if (!s) return "";
    const u = String(s).toUpperCase().replace(/\bCTNS?\b/g, "CARTON").replace(/\bPKGS?\b/g, "PACKAGE").replace(/\bPLTS?\b/g, "PALLET");
    for (const w of ["WOODEN CASE", "PALLET", "CARTON", "PACKAGE", "CASE", "BOX", "SET", "DRUM", "ROLL", "BAG", "BALE", "CRATE", "UNIT", "PIECE", "SKID", "BUNDLE"])
      if (u.includes(w)) return w === "WOODEN CASE" ? "CASE" : w;
    let m = u.match(/^\s*[\d,]*\s*([A-Z]{2,3})\b/);
    if (m && PKG_KIND[m[1]]) return PKG_KIND[m[1]];
    m = u.match(/\(([A-Z]+?)\(S\)\)/) || u.match(/\(([A-Z ]+?)\)/);
    if (m) return m[1].trim().replace(/S+$/, "") || m[1].trim();
    return u.replace(/[\d,()]/g, "").trim();
  }
  function canonStatus(s) {
    if (!s) return "";
    const th = String(s);
    if (th.includes("เปิดตู้")) return "LCL/CFS";
    if (th.includes("ลากตู้")) return "CY";
    if (th.includes("ขน")) return "LCL";
    const u = th.replace(/STATUS/gi, "").replace(/[^A-Za-z/]/g, "").toUpperCase();
    if (!u) return th.trim();
    if (u.includes("CFS")) return "LCL/CFS";
    if (u === "LCL") return "LCL";
    if (u.includes("CY") || u.includes("FCL")) return "CY";
    if (u.includes("LCL")) return "LCL";
    return u;
  }
  function num(s) {
    const m = String(s == null ? "" : s).match(/-?\d[\d,]*\.?\d*/);
    return m ? parseFloat(m[0].replace(/,/g, "")) : null;
  }
  function pkgnum(s) {
    const m = String(s == null ? "" : s).match(/\d[\d,]*/);
    return m ? parseInt(m[0].replace(/,/g, ""), 10) : null;
  }
  function normCo(s) {
    return String(s || "").toUpperCase().replace(/[^A-Z0-9 ]/g, " ")
      .replace(/\b(CO|LTD|COMPANY|LIMITED|PUBLIC|CORP|CORPORATION|INC|GROUP|THE)\b/g, " ").replace(/\s+/g, " ").trim();
  }
  function comatch(a, b) {
    const sa = new Set(a.split(" ").filter(Boolean)), sb = new Set(b.split(" ").filter(Boolean));
    if (!sa.size || !sb.size) return false;
    const inter = [...sa].filter(x => sb.has(x));
    const subA = [...sa].every(x => sb.has(x)), subB = [...sb].every(x => sa.has(x));
    if (sa.size === sb.size && subA) return true;
    if (inter.length && (subA || subB) && Math.min(sa.size, sb.size) >= 2) return true;
    return inter.length >= 2;
  }
  function sameDetail(a, b) {
    // 'CARGO MOVEMENT n' เป็นรหัสศุลกากร ไม่ใช่รายละเอียดสินค้า — MANIFEST บางไฟล์พิมพ์ไว้ท้าย DESC ('*CARGO MOVEMENT 7*') ไม่นับในการเทียบ
    const mv = s => String(s || "").toUpperCase().replace(/\*?\s*CARGO\s*MOVEMENT\s*[:=]?\s*[0-9A-Z][0-9A-Z\-]*\s*\*?/g, " ");
    const na = mv(a).replace(/[^A-Z0-9]/g, ""), nb = mv(b).replace(/[^A-Z0-9]/g, "");
    if (!na && !nb) return null;
    return na === nb;
  }
  const fmt3 = x => x.toLocaleString("en-US", { minimumFractionDigits: 3, maximumFractionDigits: 3 });
  const r2 = x => Math.round(x * 100) / 100, r3 = x => Math.round(x * 1000) / 1000;
  const blSort = (a, b) => a.length - b.length || (a < b ? -1 : a > b ? 1 : 0);

  // ───────────── 1) MANIFEST.xls ─────────────
  function cell(row, i) {
    if (i >= row.length) return "";
    let v = row[i];
    if (v === null || v === undefined) return "";
    if (typeof v === "number" && Number.isInteger(v)) v = String(v);
    return String(v).replace(/[\x00-\x1f]+/g, " ").trim();
  }
  function detectBlPrefix(rows) {
    const cand = [];
    for (const row of rows) {
      const v = row && row[0];
      if (typeof v === "number" || v == null) continue;
      const m = String(v).trim().match(/^([A-Z]{3,8})[A-Z0-9]{5,}$/);
      if (m) cand.push(m[1]);
    }
    if (cand.length) {
      let p = cand[0];
      for (const c of cand) { let i = 0; while (i < p.length && i < c.length && p[i] === c[i]) i++; p = p.slice(0, i); }
      p = (p.match(/^[A-Z]*/) || [""])[0];
      if (p.length < 3) {
        const cnt = {}; cand.forEach(c => cnt[c] = (cnt[c] || 0) + 1);
        p = Object.keys(cnt).sort((a, b) => cnt[b] - cnt[a])[0];
      }
      return new RegExp(`^${p}[A-Z0-9]{5,}$`);
    }
    return /^[A-Z]{3,8}[A-Z0-9]{6,}$/;
  }
  function detectMarksCol(rows) {
    const counts = {};
    for (const row of rows) row.forEach((v, idx) => { if (/^\d+\.\s*[A-Z]{4}\d{6,7}/.test(String(v == null ? "" : v).trim())) counts[idx] = (counts[idx] || 0) + 1; });
    const k = Object.keys(counts).sort((a, b) => counts[b] - counts[a])[0];
    return k === undefined ? 5 : +k;
  }

  function parseManifest(rowsIn) {
    const rows = rowsIn.map(r => r.map(v => (v === null || v === undefined ? "" : v)));
    const blRe = detectBlPrefix(rows), MARKS_COL = detectMarksCol(rows);
    const n = rows.length;
    let vessel = "";
    const declared = {};
    const MAN = {};
    const END = /^(PORT TOTAL|GRAND TOTAL|TOTAL\b|PortOfDischarge)/i;
    // MARKS / DESCRIPTION (col 10) use a stricter end marker: real cargo text can start with TOTAL (e.g. marks "TOTAL 4 CTNS ONLY")
    const END_CARGO = /^(PORT TOTAL|GRAND TOTAL|TOTAL\s*(PACK|GROSS\s*WEIGHT|MEASUREMENT)\b|PortOfDischarge)/i;
    const ST_TOK = ["CY/CY", "CY-CY", "LCL/CFS", "CFS/CFS", "FCL/CFS", "CY", "LCL", "CFS", "FCL"];
    const pdAt = new Array(n).fill("");
    let curPd = "";
    rows.forEach((row, ridx) => {
      const v0 = cell(row, 0);
      if (v0.toUpperCase() === "VESSEL & VOYAGE" && !vessel) {
        for (const c of row.slice(1)) { const cs = String(c).trim(); if (cs) { vessel = cs; break; } }
      }
      for (let k = 0; k < row.length; k++) {
        if (cell(row, k).toUpperCase() === "PORT OF DISCHARGE") {
          for (let k2 = k + 1; k2 < row.length; k2++) { const v = cell(row, k2); if (v) { curPd = v; break; } }
          break;
        }
      }
      pdAt[ridx] = curPd;
      const rowTxt = row.map((_, k) => cell(row, k)).join(" ");
      if (/GRAND TOTAL/i.test(rowTxt)) {
        const joined = rows.slice(ridx, ridx + 4).map(r2_ => r2_.map((_, k) => cell(r2_, k)).join(" ")).join(" ");
        const mpk = joined.match(/TOTAL PACK\s*:\s*([\d,]+)/i), mgw = joined.match(/TOTAL GROSSWEIGHT\s*:\s*([\d,.]+)/i),
          mme = joined.match(/TOTAL MEASUREMENT\s*:\s*([\d,.]+)/i);
        if (mpk) declared.pkg = parseInt(mpk[1].replace(/,/g, ""), 10);
        if (mgw) declared.gw = parseFloat(mgw[1].replace(/,/g, ""));
        if (mme) declared.meas = parseFloat(mme[1].replace(/,/g, ""));
      }
    });

    let i = 0;
    while (i < n) {
      const bl = cell(rows[i], 0);
      if (!blRe.test(bl)) { i++; continue; }
      const rec = { bl, pkg_hdr: cell(rows[i], 10), gw: cell(rows[i], 16), cons: "", notify: "", cont: [], status: "", meas: "",
        marks: [], desc: [], transit: "", dg: "", reefer: "", port_discharge: pdAt[i], shed_no: "", dest_code_raw: "", movement_raw: "", tax_id_raw: "" };
      const h5 = cell(rows[i], MARKS_COL);
      if (h5 && !/^\d+\.\s*([A-Z]{4}\d{6,7}|\s*$)/.test(h5)) rec.marks.push(h5);
      let transitOpen = false;
      let j = i + 1;
      while (j < n) {
        const a = cell(rows[j], 0);
        if (blRe.test(a)) break;
        if (rows[j].some((_, k) => (k === MARKS_COL || k === 10 ? END_CARGO : END).test(cell(rows[j], k)))) { j++; break; }
        const numbered = /^[123]\./.test(a);
        if (a.startsWith("C :")) rec.cons = a.slice(3).trim();
        else if (a.startsWith("N :")) rec.notify = a.slice(3).trim();
        else if (a.startsWith("S :")) { /* shipper */ }
        else if (rec.cons && !rec.notify && !numbered) rec.cons += " " + a;
        else if (rec.notify && !numbered) rec.notify += " " + a;
        const c5 = cell(rows[j], MARKS_COL);
        if (c5) {
          const mm = c5.match(/^\d+\.\s*([A-Z]{4}\d{6,7})/);
          if (mm) rec.cont.push(mm[1]);
          else if (!/^\d+\.\s*$/.test(c5)) rec.marks.push(c5);
        }
        for (const cc of [cell(rows[j], MARKS_COL), cell(rows[j], 10)]) {
          if (cc && DG_RE.test(cc) && !rec.dg) rec.dg = findDg(cc);
          if (cc && REEFER_RE.test(cc) && !rec.reefer) rec.reefer = findTemp(cc);
        }
        const c10 = cell(rows[j], 10);
        if (c10) {
          const u = c10.toUpperCase();
          if (DG_RE.test(c10) || REEFER_RE.test(c10)) rec.desc.push(c10);
          else if (transitOpen || TRANSIT_RE.test(u)) { rec.transit = (rec.transit + " " + c10).trim(); transitOpen = true; rec.desc.push(c10); }
          else if (ST_TOK.includes(u) && !rec.status) rec.status = c10;
          else if (/^\d{0,2}[A-Z]\d[A-Z0-9]{1,3}(?:\s|$)/.test(u) || u.includes("PART CNTR") || /^[\d,]+\s*(PK|CS|PX|CT|BX|SX|BE|DR|RO|BG)\b/.test(u)) { /* โครงสร้าง */ }
          else rec.desc.push(c10);
        }
        const c16 = cell(rows[j], 16);
        if (c16 && c16.toUpperCase().includes("MTQ") && !rec.meas) rec.meas = c16;
        const c15 = cell(rows[j], 15);
        if (c15) {
          if (c15.startsWith("1)") && !rec.shed_no) { const msh = c15.match(/\d{3,4}/); rec.shed_no = msh ? msh[0] : ""; }
          else if (c15.startsWith("2)") && !rec.dest_code_raw) rec.dest_code_raw = c15;
          else if (c15.startsWith("3)") && !rec.movement_raw) { const mmv = c15.match(/\(([^)]*)\)/); rec.movement_raw = mmv ? mmv[1].trim() : ""; }
          else if (c15.startsWith("4)") && !rec.tax_id_raw) rec.tax_id_raw = c15;
        }
        j++;
      }
      rec.marks = rec.marks.join(" ");
      rec.desc = rec.desc.join(" ");
      rec.cont = [...new Set(rec.cont)].sort();
      MAN[bl] = rec;
      i = j;
    }
    return { MAN, vessel, declared, blRe };
  }

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


  // MANIFEST บางรายการพิมพ์จำนวน/ชนิดหีบห่อซ้ำในวงเล็บต้น DESCRIPTION ('(2 WOODEN CASES) ...') ซึ่ง ENTER อยู่ในคอลัมน์
  // QUANTITY — ถ้าตรงกับ QUANTITY ของ ENTER พอดี เติมวงเล็บเดียวกันหน้า DESC ของ ENTER ไม่ให้นับเป็น DESCRIPTION ไม่ตรง
  function alignQtyPrefix(MAN, ENT) {
    const n = s => String(s || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
    for (const [bl, e] of Object.entries(ENT)) {
      const mm = ((MAN[bl] || {}).desc || "").match(/^(\([^()]*\))\s*/);
      if (mm && e.pkgs && n(mm[1]) === n(`${e.pkgs}${e.pkgtype || ""}`) && !(e.desc || "").startsWith("("))
        e.desc = `${mm[1]} ${e.desc || ""}`.trim();
    }
  }

  // ───────────── 3) รายงาน ─────────────
  const KEYS = ["urgent", "bl", "cneeE", "cneeM", "tax", "cntE", "cntM", "stE", "stM", "tpE", "pkE", "tpM", "pkM",
    "gwE", "gwM", "meE", "meM", "mkE", "mkM", "dsE", "dsM", "reefer", "dg", "transit", "shed", "dest", "mv", "note"];
  const HEADERS = ["สรุปผลเร่งด่วน", "B/L NO.", "CNEE. (ENTER)", "CNEE. (MANIFEST)", "TAX ID vs NOTIFY", "CNTRS NO. (ENTER)", "CNTRS NO. (MANIFEST)",
    "STATUS (ENTER)", "STATUS (MANIFEST)", "TOTAL PKG (ENTER)", "PKG. (ENTER)", "TOTAL PKG (MANIFEST)", "PKG. (MANIFEST)",
    "G.W. KGM (ENTER)", "G.W. KGM (MANIFEST)", "MEAS. MTQ (ENTER)", "MEAS. MTQ (MANIFEST)", "MARKS (ENTER)", "MARKS (MANIFEST)",
    "DESC. (ENTER)", "DESC. (MANIFEST)", "REEFER TEMP", "DG", "TRANSIT/TRANSHIPMENT", "SHED NO.", "ประเทศปลายทาง", "CARGO MOVEMENT", "หมายเหตุ / จุดที่ไม่ตรงกัน"];
  const NUMKEYS = new Set(["gwE", "gwM", "meE", "meM"]), WIDEKEYS = new Set(["mkE", "mkM", "dsE", "dsM"]);
  const LBL = { cneeE: "CONSIGNEE", cneeM: "CONSIGNEE", cntE: "CONTAINER", cntM: "CONTAINER", stE: "STATUS", stM: "STATUS",
    tpE: "PACKAGE", tpM: "PACKAGE", pkE: "PACKAGING", pkM: "PACKAGING", gwE: "G.W.", gwM: "G.W.", meE: "MEAS.", meM: "MEAS.",
    mkE: "MARKS", mkM: "MARKS", dsE: "DESC", dsM: "DESC", reefer: "REEFER", dg: "DG", transit: "TRANSIT",
    shed: "SHED", dest: "ปลายทาง", mv: "MOVEMENT", tax: "TAX ID" };
  const OWN_URGENT = new Set(["reefer", "dg", "transit", "shed", "dest", "mv", "tax"]);

  const pyf = x => (x === null || x === undefined ? "None" : Number.isInteger(x) ? x.toFixed(1) : String(x));

  function buildReport(MAN, ENT, vessel, manDecl, entDecl, opts) {
    opts = opts || {};
    alignQtyPrefix(MAN, ENT);
    const expectedShed = (opts.expectedShed || "").replace(/\D/g, "").padStart(4, "0").slice(-4);
    const exShed = opts.expectedShed && /\d{3,4}/.test(opts.expectedShed) ? opts.expectedShed.match(/\d{3,4}/)[0].padStart(4, "0") : "";
    const rows = [], crit = [];
    let issueRows = 0;
    const TOT = { pe: 0, pm: 0, ge: 0, gm: 0, me: 0, mm: 0, pke: {}, pkm: {}, cte: {}, ctm: {}, st_bad: 0, transit: 0, dg: 0, reefer: 0,
      shed_bad: 0, dest_bad: 0, mv_bad: 0, mv_review: 0, tax_bad: 0 };
    const sub = Object.keys(MAN).sort(blSort);
    const mkCell = (v, s, bad) => ({ v: v === null || v === undefined ? "" : v, s: s || "ok", bad: !!bad });

    for (const bl of sub) {
      const m = MAN[bl], e = ENT[bl];
      const notes = [], urgent = [], cs = {};
      const mk = (keys, ok) => keys.forEach(k => { cs[k] = ok ? "ok" : "diff"; });

      if (!e) {
        const pm_ = pkgnum(m.pkg_hdr), gm_ = num(m.gw), mm_ = num(m.meas);
        const xc = computeExtraChecks(m, null, opts.shedFor ? opts.shedFor(m, null) : exShed);
        const extraNotes = ["shed", "dest", "mv", "tax"].map(k => xc[k].note).filter(Boolean);
        const vals = { urgent: `${WARN} ไม่มีเอกสาร ENTER เทียบ`, bl: `${WARN} ${bl}`, cneeE: "—", cneeM: m.cons || m.notify, tax: xc.tax.text,
          cntE: "—", cntM: m.cont.join(";"), stE: "—", stM: canonStatus(m.status) || m.status || "(ว่าง)",
          tpE: "—", pkE: "—", tpM: pm_ !== null ? pm_ : (m.pkg_hdr || "-"), pkM: canonPkg(m.pkg_hdr) || "-",
          gwE: "—", gwM: gm_ !== null ? fmt3(gm_) : "-", meE: "—", meM: mm_ !== null ? fmt3(mm_) : "-",
          mkE: "—", mkM: m.marks, dsE: "—", dsM: m.desc, reefer: "-", dg: "-", transit: extractTransitPhrase(m.transit) || "-",
          shed: xc.shed.text, dest: xc.dest.text, mv: xc.mv.text, note: "ไม่พบ B/L นี้ในเอกสาร ENTER" + (extraNotes.length ? " ; " + extraNotes.join(" ; ") : "") };
        const okx = { shed: xc.shed.ok, dest: xc.dest.ok, mv: xc.mv.ok, tax: xc.tax.ok };
        if (xc.shed.ok === false) TOT.shed_bad++;
        if (xc.dest.ok === false) TOT.dest_bad++;
        if (xc.mv.ok === false) TOT.mv_bad++;
        if (xc.mv.ok === "review") TOT.mv_review++;
        if (xc.tax.ok === false) TOT.tax_bad++;
        const cells = {};
        for (const k of KEYS) {
          const o = okx[k];
          cells[k] = mkCell(vals[k], o === false ? "diff" : o === "review" ? "warn" : "yel", ["urgent", "bl", "note"].includes(k) || o === false || o === "review");
        }
        if (pm_) TOT.pm += pm_; if (gm_) TOT.gm += gm_; if (mm_) TOT.mm += mm_;
        rows.push({ kind: "noenter", bl, cells, bad: true });
        issueRows++;
        crit.push(`${bl}: ไม่มีเอกสาร ENTER`);
        extraNotes.forEach(x => crit.push(`${bl}: ${x}`));
        continue;
      }

      // CONSIGNEE
      const ne = normCo(e.cons), mc = normCo(m.cons), mn = normCo(m.notify);
      const hitC = !!(ne && mc && comatch(ne, mc)), hitN = !!(ne && mn && comatch(ne, mn));
      const consOk = hitC || hitN;
      mk(["cneeE", "cneeM"], consOk);
      if (!consOk) notes.push(`CONSIGNEE ไม่ตรง (ENTER: ${e.cons} / MANIFEST C:${m.cons} N:${m.notify})`);
      else if (hitN && !hitC) { cs.cneeE = "warn"; notes.push(`CONSIGNEE(MANIFEST) ช่อง C: = '${m.cons}' ไม่ครบ — ชื่อเต็มตรงกับ Notify`); }

      // CONTAINER
      const eCs = [...new Set((e.cont || "").match(new RegExp(CONT_RE.source, "g")) || [])].sort(), eC = eCs.join(";");
      const contOk = eCs.length > 1 ? (eCs.length === m.cont.length && eCs.every(c => m.cont.includes(c))) : !!(eC && m.cont.includes(eC));
      mk(["cntE", "cntM"], contOk);
      if (!contOk) notes.push(`CONTAINER ไม่ตรง (ENTER: ${eC || "-"} / MANIFEST: ${m.cont.join(";") || "-"})`);

      // STATUS
      const seRaw = e.status_raw || "", smRaw = m.status || "";
      const se = canonStatus(seRaw), sm = canonStatus(smRaw);
      const stOk = se && sm ? se === sm : true;
      mk(["stE", "stM"], stOk);
      const dispE = seRaw ? seRaw + (se && seRaw.trim() !== se ? `  (= ${se})` : "") : "(ไม่ระบุใน ENTER)";
      const dispM = smRaw ? smRaw + (sm && smRaw !== sm ? `  (= ${sm})` : "") : "(ว่าง)";
      if (!stOk) { TOT.st_bad++; notes.push(`STATUS ไม่ตรง (ENTER ${seRaw} = ${se} / MANIFEST ${smRaw} = ${sm})`); }

      // PACKAGE
      const pe = e.pkgs, pm = pkgnum(m.pkg_hdr);
      const pkOk = pe !== null && pm !== null && pe === pm;
      mk(["tpE", "tpM"], pkOk);
      if (!pkOk) notes.push(`TOTAL PACKAGE ไม่ตรง (ENTER ${pe !== null ? pe : "-"} / MANIFEST ${pm !== null ? pm : "(ไม่มี)"})`);
      const ke = canonPkg(e.pkgtype), km = canonPkg(m.pkg_hdr);
      const pgOk = !!(ke && km && (ke === km || km.includes(ke) || ke.includes(km)));
      mk(["pkE", "pkM"], pgOk);
      if (!pgOk) notes.push(`PACKAGING ไม่ตรง (ENTER ${ke || e.pkgtype || "-"} / MANIFEST ${km || m.pkg_hdr || "(ไม่มี)"})`);

      // GW / MEAS
      const ge = e.gw, gm = num(m.gw);
      let gwOk = ge !== null && gm !== null && r2(ge) === r2(gm);
      if ((gm === 0 || gm === null) && ge) gwOk = false;
      mk(["gwE", "gwM"], gwOk);
      if (!gwOk) notes.push(`GROSS WEIGHT ไม่ตรง (ENTER ${pyf(ge)} / MANIFEST ${pyf(gm)})` + ((gm === 0 || gm === null) && ge ? " — MANIFEST = 0.000 (ค่าหาย) แต่ ENTER มีค่าจริง" : ""));
      const me = e.meas, mm = num(m.meas);
      let msOk = me !== null && mm !== null && r3(me) === r3(mm);
      if ((mm === 0 || mm === null) && me) msOk = false;
      mk(["meE", "meM"], msOk);
      if (!msOk) notes.push(`MEASUREMENT ไม่ตรง (ENTER ${pyf(me)} / MANIFEST ${pyf(mm)})` + ((mm === 0 || mm === null) && me ? " — MANIFEST = 0.000 (ค่าหาย) แต่ ENTER มีค่าจริง" : ""));

      // MARKS / DESCRIPTION (เข้มงวด — ตัดช่องว่าง/วรรคตอนเท่านั้น)
      const mdLvl = e.combined_text ? "warn" : "diff";
      const mdTxt = (lbl, a, b) => e.combined_text ? `${lbl} ไม่ตรงกัน (ไม่นับช่องว่าง) — (ENTER: ${a || "(ว่าง)"} / MANIFEST: ${b || "(ว่าง)"})`
        : `${lbl} ไม่เหมือนกัน (ENTER: ${a || "(ว่าง)"} / MANIFEST: ${b || "(ว่าง)"})`;
      if (sameDetail(e.marks, m.marks) === false) { cs.mkE = cs.mkM = mdLvl; notes.push(mdTxt("MARKS", e.marks, m.marks)); }
      if (sameDetail(e.desc, m.desc) === false) { cs.dsE = cs.dsM = mdLvl; notes.push(mdTxt("DESCRIPTION", e.desc, m.desc)); }

      // REEFER / DG
      const reE = e.reefer || "", reM = m.reefer || "", dgE = e.dg || "", dgM = m.dg || "";
      const both = (a, b) => [a ? `ENTER: ${a}` : "", b ? `MANIFEST: ${b}` : ""].filter(Boolean).join(" | ") || "-";
      const reefer = both(reE, reM), dg = both(dgE, dgM);
      if (reE || reM) {
        mk(["reefer"], false); TOT.reefer++;
        notes.push(`REEFER TEMP ${reefer} — ${!!reE !== !!reM ? "อีกฝั่งไม่ระบุ ; " : ""}ต้องสำแดงตู้เย็น/ยืนยันอุณหภูมิในใบขน`);
        urgent.push(`${WARN} REEFER ${reE || reM}`);
      }
      if (dgE || dgM) {
        mk(["dg"], false); TOT.dg++;
        notes.push(`DG ${dg} — ${!!dgE !== !!dgM ? "อีกฝั่งไม่ระบุ ; " : ""}ต้องสำแดงวัตถุอันตราย (CLASS/UN) + แนบ MSDS/DGD`);
        urgent.push(`${WARN} DG ${dgE || dgM}`);
      }

      // TRANSIT
      const parts = [extractTransitPhrase(e.transit_raw || ""), extractTransitPhrase(m.transit || "")].filter(Boolean);
      const transitDisp = parts.length ? [...new Set(parts)].join(" / ") : "-";
      if (parts.length) {
        mk(["transit"], false); TOT.transit++;
        notes.push(`พบคำ transit/transhipment — ${transitDisp} (ไม่ใช่ LOCAL, ตรวจประเภทใบขน: ถ่ายลำ/ผ่านแดน)`);
        urgent.push(`${WARN} ${transitDisp}`);
      }

      // 4 กฎเพิ่มเติม
      const xc = computeExtraChecks(m, e, opts.shedFor ? opts.shedFor(m, e) : exShed);
      for (const [key, tag] of [["shed", "SHED"], ["dest", "ปลายทาง"], ["mv", "MOVEMENT"], ["tax", "TAX ID"]]) {
        const o = xc[key].ok;
        if (o === null) continue;
        if (o === "review") {
          cs[key] = "warn"; TOT.mv_review++;
          notes.push(xc[key].note); urgent.push(`${WARN} ${tag} ต้องตรวจสอบด้วยคน`);
          crit.push(`${bl}: ${tag} ต้องตรวจสอบด้วยคน — ${xc[key].note}`);
        } else {
          cs[key] = o ? "ok" : "diff";
          if (!o) {
            if (key === "shed") TOT.shed_bad++; else if (key === "dest") TOT.dest_bad++; else if (key === "mv") TOT.mv_bad++; else TOT.tax_bad++;
            notes.push(xc[key].note); urgent.push(`${WARN} ${tag} ${xc[key].text}`);
            crit.push(`${bl}: ${xc[key].note}`);
          }
        }
      }

      if (e.low_confidence) {
        for (const k of Object.keys(cs)) if (cs[k] === "diff" && !["reefer", "dg", "transit", "shed", "dest", "tax"].includes(k)) cs[k] = "warn";
        notes.push("ENTER เอกสารนี้ดึงข้อมูลด้วยความมั่นใจต่ำ (layout ไม่ชัดเจน/พบ anchor น้อยเกินไป) — ตรวจสอบทุกช่องที่ไฮไลต์ส้มด้วยคน");
      }

      const diff = Object.keys(cs).filter(k => cs[k] === "diff"), warnCols = Object.keys(cs).filter(k => cs[k] === "warn");
      if (diff.length || warnCols.length) issueRows++;
      if (diff.length) {
        const tags = [...new Set(diff.filter(k => !OWN_URGENT.has(k)).map(k => LBL[k]))].sort();
        if (tags.length) { urgent.unshift(`${WARN} ไม่ตรง: ` + tags.join(", ")); crit.push(`${bl}: ไม่ตรง ` + tags.join(", ")); }
      } else if (warnCols.length) {
        const wt = [...new Set(warnCols.filter(k => LBL[k] && !OWN_URGENT.has(k)).map(k => LBL[k]))].sort();
        urgent.unshift(wt.length ? `${WARN} ต้องตรวจสอบ: ` + wt.join(", ") : `${WARN} ต้องตรวจสอบด้วยคน`);
      }
      if (urgent.length && !urgent[0].startsWith(WARN)) urgent[0] = `${WARN} ${urgent[0]}`;
      const urgentTxt = urgent.length ? [...new Set(urgent)].join(" ; ") : `${OK} ผ่าน`;

      if (pe !== null) TOT.pe += pe; if (pm !== null) TOT.pm += pm;
      if (ge !== null) TOT.ge += ge; if (gm !== null) TOT.gm += gm;
      if (me !== null) TOT.me += me; if (mm !== null) TOT.mm += mm;
      if (ke) TOT.pke[ke] = (TOT.pke[ke] || 0) + (pe || 0);
      if (km) TOT.pkm[km] = (TOT.pkm[km] || 0) + (pm || 0);
      eCs.forEach(x => { TOT.cte[x] = (TOT.cte[x] || 0) + 1; });
      m.cont.forEach(x => { TOT.ctm[x] = (TOT.ctm[x] || 0) + 1; });

      const vals = { urgent: urgentTxt, bl, cneeE: e.cons, cneeM: m.cons || m.notify, tax: xc.tax.text, cntE: eC, cntM: m.cont.join(";"),
        stE: dispE, stM: dispM, tpE: pe !== null ? pe : "-", pkE: ke || e.pkgtype || "-", tpM: pm !== null ? pm : "(ไม่มี)", pkM: km || m.pkg_hdr || "(ไม่มี)",
        gwE: ge !== null ? fmt3(ge) : "-", gwM: gm !== null ? fmt3(gm) : "-", meE: me !== null ? fmt3(me) : "-", meM: mm !== null ? fmt3(mm) : "-",
        mkE: e.marks, mkM: m.marks, dsE: e.desc, dsM: m.desc, reefer, dg, transit: transitDisp,
        shed: xc.shed.text, dest: xc.dest.text, mv: xc.mv.text, note: notes.length ? notes.join(" ; ") : "-" };
      const hasDiff = diff.length > 0, hasWarn = warnCols.length > 0;
      const cells = {};
      for (const k of KEYS) {
        const s = cs[k] || "ok";
        let v = vals[k];
        cells[k] = mkCell(v, s, s === "diff" || s === "warn");
        if ((s === "diff" || s === "warn") && v !== "" && v != null && !String(v).startsWith(WARN)) cells[k].v = `${WARN} ${v}`;
      }
      cells.urgent = hasDiff ? mkCell(urgentTxt, "diff", true) : hasWarn ? mkCell(urgentTxt, "warn", true) : mkCell(urgentTxt, "grn", true);
      const nv = vals.note;
      cells.note = (nv === "-" || nv === "") ? mkCell(nv, "ok", false)
        : mkCell((hasDiff || hasWarn) && !String(nv).startsWith(WARN) ? `${WARN} ${nv}` : nv, hasDiff ? "diff" : hasWarn ? "warn" : "ok", hasDiff || hasWarn);
      rows.push({ kind: "normal", bl, cells, bad: hasDiff || hasWarn });
    }

    // B/L ที่อยู่ใน ENTER แต่ไม่มีใน MANIFEST เลย
    for (const bl of Object.keys(ENT).filter(b => !(b in MAN)).sort(blSort)) {
      const e = ENT[bl];
      const generic = !!e.low_confidence;
      const head = generic ? `${WARN} พบข้อความคล้ายเลข B/L นี้ใน ENTER (ความมั่นใจต่ำ) แต่ไม่มีใน MANIFEST — ตรวจสอบด้วยคนว่าเป็น B/L จริงหรือขยะจากการอ่านข้อความ`
        : `${WARN} มีใน ENTER แต่ไม่มีใน MANIFEST — เสี่ยงตกหล่นจากใบขน`;
      const note = generic ? `${WARN} B/L นี้ไม่พบใน MANIFEST — ดึงข้อมูลมาด้วยความมั่นใจต่ำ ตรวจสอบด้วยคนก่อนแจ้งสายเรือ/ลูกค้า`
        : `${WARN} B/L นี้ไม่มีอยู่ใน MANIFEST เลย — ของอาจตกหล่นจากใบขนทั้งรายการ ต้องแจ้งสายเรือ/ลูกค้าด่วน`;
      const vals = { urgent: head, bl: `${WARN} ${bl}`, cneeE: e.cons, cneeM: "—", tax: "—", cntE: e.cont || "", cntM: "—",
        stE: e.status_raw || "(ไม่ระบุใน ENTER)", stM: "—", tpE: e.pkgs !== null ? e.pkgs : "-", pkE: canonPkg(e.pkgtype) || e.pkgtype || "-", tpM: "—", pkM: "—",
        gwE: e.gw !== null ? fmt3(e.gw) : "-", gwM: "—", meE: e.meas !== null ? fmt3(e.meas) : "-", meM: "—",
        mkE: e.marks, mkM: "—", dsE: e.desc, dsM: "—", reefer: e.reefer || "-", dg: e.dg || "-",
        transit: extractTransitPhrase(e.transit_raw) || "-", shed: "—", dest: "—", mv: "—", note };
      const cells = {};
      for (const k of KEYS) cells[k] = mkCell(vals[k], generic ? "warn" : "diff", true);
      rows.push({ kind: "extra", bl, cells, bad: true });
      if (e.pkgs) TOT.pe += e.pkgs; if (e.gw) TOT.ge += e.gw; if (e.meas) TOT.me += e.meas;
      issueRows++;
      crit.unshift(`${bl}: ` + (generic ? "⚠ อยู่ใน ENTER (layout ไม่รู้จัก) แต่ไม่มีใน MANIFEST — ตรวจสอบด้วยคน" : "⚠⚠ อยู่ใน ENTER แต่ไม่มีใน MANIFEST เลย (วิกฤต)"));
    }

    // แถวยอดรวม
    const peOk = TOT.pe === TOT.pm, geOk = r2(TOT.ge) === r2(TOT.gm), meOk = r3(TOT.me) === r3(TOT.mm);
    const pk = d => Object.keys(d).sort().map(k => `${k} ${d[k]}`).join(" / ") || "-";
    const ct = d => Object.keys(d).sort().join(", ") || "-";
    const na = sub.length;
    const allOk = peOk && geOk && meOk && !TOT.st_bad && !TOT.transit && !TOT.dg && !TOT.reefer && !TOT.shed_bad && !TOT.dest_bad && !TOT.mv_bad && !TOT.mv_review && !TOT.tax_bad;
    const dbits = [];
    const dtxt = d => ["pkg", "gw", "meas"].filter(k => k in d).map(k => (k === "pkg" ? String(d[k]) : fmt3(d[k]))).join(" / ");
    if (entDecl && Object.keys(entDecl).length) dbits.push("ENTER: " + dtxt(entDecl));
    if (manDecl && Object.keys(manDecl).length) dbits.push("MANIFEST(GRAND TOTAL): " + dtxt(manDecl));
    const declTxt = dbits.length ? "ยอดที่ระบุในเอกสาร — " + dbits.join("  |  ") + "  ||  " : "";
    // totals equal but other flags (TRANSIT/DG/...) — say what it is instead of "ยอดรวมไม่ตรง"
    const flags = [["st_bad", "STATUS ไม่ตรง"], ["transit", "TRANSIT"], ["dg", "DG"], ["reefer", "REEFER"], ["shed_bad", "SHED ไม่ตรง"],
      ["dest_bad", "ปลายทางไม่ตรง"], ["mv_bad", "MOVEMENT ไม่ตรง"], ["mv_review", "MOVEMENT ต้องตรวจ"], ["tax_bad", "TAX ID ไม่ตรง"]]
      .filter(([k]) => TOT[k]).map(([k, l]) => `${l} ${TOT[k]} B/L`);
    const T = {
      urgent: allOk ? `${OK} ยอดรวม ENTER = MANIFEST`
        : (peOk && geOk && meOk) ? `${WARN} ยอดรวมตรงกัน แต่มีจุดต้องตรวจ: ${flags.join(", ")}`
        : `${WARN} ยอดรวมไม่ตรง — ตรวจด่วน`, bl: "TOTAL / ยอดรวม",
      cneeE: `${na} B/L ย่อย`, cneeM: `${na} B/L ย่อย`, tax: `TAX ID ไม่ตรง ${TOT.tax_bad} B/L`,
      cntE: `${Object.keys(TOT.cte).length} ตู้: ${ct(TOT.cte)}`, cntM: `${Object.keys(TOT.ctm).length} ตู้: ${ct(TOT.ctm)}`,
      stE: `STATUS ไม่ตรง ${TOT.st_bad} B/L`, stM: `STATUS ไม่ตรง ${TOT.st_bad} B/L`,
      tpE: TOT.pe, pkE: pk(TOT.pke), tpM: TOT.pm, pkM: pk(TOT.pkm),
      gwE: fmt3(TOT.ge), gwM: fmt3(TOT.gm), meE: fmt3(TOT.me), meM: fmt3(TOT.mm),
      mkE: "", mkM: "", dsE: "", dsM: "", reefer: `${TOT.reefer} B/L`, dg: `${TOT.dg} B/L`, transit: `${TOT.transit} B/L`,
      shed: `SHED ไม่ตรง ${TOT.shed_bad} B/L`, dest: `ปลายทางไม่ตรง ${TOT.dest_bad} B/L`, mv: `MOVEMENT ไม่ตรง ${TOT.mv_bad} / ต้องตรวจ ${TOT.mv_review} B/L`,
      note: declTxt + `ตรวจยอดรวมจากรายการข้างบน: PACKAGE ${peOk ? "ตรง" : "ไม่ตรง"}, GROSS WEIGHT ${geOk ? "ตรง" : "ไม่ตรง"}, MEASUREMENT ${meOk ? "ตรง" : "ไม่ตรง"}` };
    const bad = {};
    if (!peOk) bad.tpE = bad.tpM = 1; if (!geOk) bad.gwE = bad.gwM = 1; if (!meOk) bad.meE = bad.meM = 1;
    if (TOT.st_bad) bad.stE = bad.stM = 1;
    if (TOT.reefer) bad.reefer = 1; if (TOT.dg) bad.dg = 1; if (TOT.transit) bad.transit = 1;
    if (TOT.shed_bad) bad.shed = 1; if (TOT.dest_bad) bad.dest = 1; if (TOT.mv_bad + TOT.mv_review) bad.mv = 1; if (TOT.tax_bad) bad.tax = 1;
    if (!(peOk && geOk && meOk)) bad.note = 1;
    const tcells = {};
    for (const k of KEYS) {
      const isBad = !!bad[k];
      tcells[k] = { v: isBad && k !== "note" ? `${WARN} ${T[k]}` : T[k], s: isBad ? "diff" : "ok", bad: isBad, total: true };
    }
    tcells.urgent = { v: T.urgent, s: allOk ? "grn" : "diff", bad: !allOk, total: true };
    tcells.bl = { v: T.bl, s: "navy", total: true };

    const masterBl = sub.length ? sub.slice().sort((a, b) => a.length - b.length)[0] : "";
    return { rows, total: { cells: tcells }, crit, issueRows, count: na, vessel, masterBl, TOT,
      stats: { bl: na, issue_rows: issueRows, allOk, totalsOk: peOk && geOk && meOk }, expectedShed: exShed };
  }

  // ───────────── อ่าน PDF ด้วย pdf.js → pages [{items:[{x,y,w,t}], annots:[{x0,y0,x1,y1,text,lines,t}], text}] ─────────────
  // เหมือน pdfPages ของเว็บพี่น้อง (รวมชิ้นข้อความแถวเดียวกันที่ชิดกัน, หมุนพิกัดกลับถ้าเนื้อหาถูกวาดหมุน 90°)
  // ต่างกันที่เก็บกรอบเต็มของ FreeText annotation ไว้ (ตัวอ่าน HAHM ใช้กรอบตัดข้อความที่ annotation วาดทับออก)
  // และไม่ใส่ข้อความ annotation ลงใน items (ตัวอ่าน CMS ได้รับแยกใน parseShared)
  async function pdfPages(pdfjsLib, buf) {
    const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
    const pages = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      const vp = page.getViewport({ scale: 1 }), W = vp.width, H = vp.height;
      const all = (await page.getTextContent()).items.filter(i => i.str);
      const src = all.filter(i => i.str.trim() !== "");
      let up = 0, down = 0, flat = 0;
      for (const i of src) {
        const [a, b] = i.transform;
        if (Math.abs(b) > Math.abs(a)) (b > 0 ? up++ : down++); else flat++;
      }
      const rot = up > flat && up >= down ? "up" : down > flat ? "down" : "";
      const conv = i => {
        const px = i.transform[4], py = i.transform[5];
        const pos = rot === "up" ? { x: py, y: px } : rot === "down" ? { x: H - py, y: W - px } : { x: px, y: H - py };
        // อักขระควบคุม (\x03 ฯลฯ) ที่บาง PDF ใช้แทนช่องว่าง → ช่องว่าง (เหมือน _clean_get_text ของ build_edi.py)
        return { x: pos.x, y: pos.y, w: i.width || 0, t: i.str.replace(/[\x00-\x08\x0b-\x1f\x7f]/g, " "),
          fs: Math.hypot(i.transform[2], i.transform[3]) || 8, eol: !!i.hasEOL };
      };
      const raw = src.map(conv);
      // บรรทัดแบบ PyMuPDF (ตัวอ่าน HAHM ใช้): ต่อคำตามลำดับที่ PDF วาด (content stream) เมื่ออยู่ baseline เดียวกันและห่างไม่เกิน ~1.5 ตัวอักษร
      // — ป้ายกับค่าที่วาดคนละช่วง (เช่น 'B/L NO. :' กับเลข B/L) จึงแยกบรรทัดเหมือน Python ; เก็บชิ้นที่เป็นช่องว่างล้วนไว้ด้วย
      // (ช่องว่างหลายช่องในต้นฉบับ เช่น 'CFS/CFE   แจ้ง…' คงไว้เหมือน PyMuPDF)
      const lines = [];
      let ln = null;
      for (const it of all.map(conv)) {
        if (!it.t.trim()) { if (ln && !ln.eol && Math.abs(it.y - ln.y) <= 1.5) ln.t += it.t; continue; }   // ไม่ขยายความกว้าง
        const gap = ln ? it.x - (ln.x + ln.w) : 0;
        if (ln && !ln.eol && Math.abs(it.y - ln.y) <= 1.5 && gap >= -1 && gap <= Math.max(it.fs, ln.fs) * 1.5) {
          ln.t += (gap > Math.max(it.fs, 4) * 0.15 && !ln.t.endsWith(" ") && !it.t.startsWith(" ") ? " " : "") + it.t;
          ln.w = it.x + it.w - ln.x;
          ln.eol = it.eol;
        } else { ln = { x: it.x, y: it.y, w: it.w, t: it.t, fs: it.fs, eol: it.eol }; lines.push(ln); }
      }
      raw.sort((a, b) => a.y - b.y || a.x - b.x);
      const items = [];
      let cur = null;
      for (const it of raw) {
        if (cur && Math.abs(it.y - cur.y) <= 1.5 && it.x - (cur.x + cur.w) <= 1.5 && it.x >= cur.x) {
          const gap = it.x - (cur.x + cur.w);
          cur.t += (gap > 1.5 && !cur.t.endsWith(" ") ? " " : "") + it.t;
          cur.w = it.x + it.w - cur.x;
        } else { cur = { x: it.x, y: it.y, w: it.w, t: it.t }; items.push(cur); }
      }
      items.forEach(i => { i.t = i.t.trim(); });   // คงช่องว่างตามต้นฉบับ (เหมือน PyMuPDF) ไม่ยุบหลายช่องเหลือช่องเดียว
      const annots = [];
      try {
        for (const a of await page.getAnnotations()) {
          const text = a.contents || (a.contentsObj && a.contentsObj.str);
          if (a.subtype !== "FreeText" || !text) continue;
          const [x1, y1, x2, y2] = a.rect;
          const box = rot === "up" ? { x0: y1, y0: x1, x1: y2, y1: x2 } : rot === "down" ? { x0: H - y2, y0: W - x2, x1: H - y1, y1: W - x1 }
            : { x0: x1, y0: H - y2, x1: x2, y1: H - y1 };
          const lines = String(text).split(/[\r\n]+/).map(s => s.replace(/\s+/g, " ").trim()).filter(Boolean);
          if (lines.length) annots.push({ ...box, text, lines, t: lines.join(" ") });
        }
      } catch (e) { /* ไม่มี annotation */ }
      const its = items.filter(i => i.t);
      // top = ขอบบนโดยประมาณ (baseline − 0.8 × ขนาดฟอนต์) เหมือน bbox ของ PyMuPDF — ใช้เฉพาะตอนเทียบกับกรอบ annotation
      const lns = lines.map(l => ({ x: l.x, y: l.y, top: l.y - 0.8 * l.fs, w: l.w, t: l.t.trim() })).filter(l => l.t);
      pages.push({ items: its, lines: lns, annots, text: lns.map(l => l.t).join("\n") });
    }
    return pages;
  }

  root.CheckHAHM = { pdfPages, HEADERS, KEYS, NUMKEYS, WIDEKEYS, parseManifest, parseEnter, alignQtyPrefix, applyHahmRules,
    buildReport, annotateReport, shedByStatus, WARN, OK,
    _t: { canonPkg, canonStatus, normCo, comatch, sameDetail, extractTransitPhrase, subseqExtra, pageKind } };
})(typeof window !== "undefined" ? window : globalThis);
