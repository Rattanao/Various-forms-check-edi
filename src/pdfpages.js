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
