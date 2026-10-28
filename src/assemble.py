"""ประกอบ docs/core.js ของ Check EDI Various forms (HAHM)
   = ตัวช่วยร่วม + parseManifest + buildReport จาก core.js ของเว็บ Freight Links (คัดลอก ไม่แก้ต้นฉบับ)
   + ตัวอ่าน ENTER ชุดรวม HAHM (src/hahm_enter.js) + pdfPages ที่เก็บกรอบ annotation
   รัน:  python src/assemble.py   (จากโฟลเดอร์ repo)"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
FL = os.path.join(os.path.dirname(REPO), 'check edi-Freight links', 'docs', 'core.js')
src = open(FL, encoding='utf-8').read().split('\n')


def find(prefix, start=0):
    for i in range(start, len(src)):
        if src[i].startswith(prefix):
            return i
    sys.exit(f'!! ไม่พบบรรทัด: {prefix}')


i_common = find('  // ───────────── regex / ตัวช่วยร่วม')
i_enter = find('  // ───────────── 2) ENTER.pdf')
i_align = find('  // MANIFEST บางรายการพิมพ์จำนวน')
i_pdf = find('  // ───────────── อ่าน PDF ด้วย pdf.js')
common = '\n'.join(src[i_common:i_enter])
# canon_pkg ของ build_edi.py: 'PLTS' = PALLET (เว็บ MAC-NELS มีแล้ว แต่ Freight Links ไม่มี)
_old = '.replace(/\\bPKGS?\\b/g, "PACKAGE");'
if common.count(_old) != 1:
    sys.exit('!! แพตช์ canonPkg ไม่ตรง')
common = common.replace(_old, '.replace(/\\bPKGS?\\b/g, "PACKAGE").replace(/\\bPLTS?\\b/g, "PALLET");')
report = '\n'.join(src[i_align:i_pdf])


def sub(text, old, new, count=1):
    if text.count(old) != count:
        sys.exit(f'!! แพตช์ไม่ตรง ({text.count(old)} ครั้ง): {old[:70]}')
    return text.replace(old, new)


# หลายตู้ต่อ B/L ต้องตรงครบชุด (เหมือน build_edi.py / เว็บ Union World)
report = sub(report,
             '      const ecm = (e.cont || "").match(CONT_RE), eC = ecm ? ecm[0] : "";\n'
             '      const contOk = !!(eC && m.cont.includes(eC));',
             '      const eCs = [...new Set((e.cont || "").match(new RegExp(CONT_RE.source, "g")) || [])].sort(), eC = eCs.join(";");\n'
             '      const contOk = eCs.length > 1 ? (eCs.length === m.cont.length && eCs.every(c => m.cont.includes(c))) : !!(eC && m.cont.includes(eC));')
report = sub(report, '      if (eC) TOT.cte[eC] = (TOT.cte[eC] || 0) + 1;',
             '      eCs.forEach(x => { TOT.cte[x] = (TOT.cte[x] || 0) + 1; });')
# จดหมายข้อความอิสระ (combined_text) → MARKS/DESC ไม่ตรง = ส้ม (เหมือน build_edi.py / เว็บ Able)
report = sub(report,
             '      if (sameDetail(e.marks, m.marks) === false) { mk(["mkE", "mkM"], false); notes.push(`MARKS ไม่เหมือนกัน (ENTER: ${e.marks || "(ว่าง)"} / MANIFEST: ${m.marks || "(ว่าง)"})`); }\n'
             '      if (sameDetail(e.desc, m.desc) === false) { mk(["dsE", "dsM"], false); notes.push(`DESCRIPTION ไม่เหมือนกัน (ENTER: ${e.desc || "(ว่าง)"} / MANIFEST: ${m.desc || "(ว่าง)"})`); }',
             '      const mdLvl = e.combined_text ? "warn" : "diff";\n'
             '      const mdTxt = (lbl, a, b) => e.combined_text ? `${lbl} ไม่ตรงกัน (ไม่นับช่องว่าง) — (ENTER: ${a || "(ว่าง)"} / MANIFEST: ${b || "(ว่าง)"})`\n'
             '        : `${lbl} ไม่เหมือนกัน (ENTER: ${a || "(ว่าง)"} / MANIFEST: ${b || "(ว่าง)"})`;\n'
             '      if (sameDetail(e.marks, m.marks) === false) { cs.mkE = cs.mkM = mdLvl; notes.push(mdTxt("MARKS", e.marks, m.marks)); }\n'
             '      if (sameDetail(e.desc, m.desc) === false) { cs.dsE = cs.dsM = mdLvl; notes.push(mdTxt("DESCRIPTION", e.desc, m.desc)); }')
# ตัวเลขในหมายเหตุ G.W./MEAS. พิมพ์แบบ Python (float: 20.0 / 0.0 ; None)
report = sub(report, '`GROSS WEIGHT ไม่ตรง (ENTER ${ge} / MANIFEST ${gm})`', '`GROSS WEIGHT ไม่ตรง (ENTER ${pyf(ge)} / MANIFEST ${pyf(gm)})`')
report = sub(report, '`MEASUREMENT ไม่ตรง (ENTER ${me} / MANIFEST ${mm})`', '`MEASUREMENT ไม่ตรง (ENTER ${pyf(me)} / MANIFEST ${pyf(mm)})`')
report = report.replace('  function buildReport(', '  const pyf = x => (x === null || x === undefined ? "None" : Number.isInteger(x) ? x.toFixed(1) : String(x));\n\n  function buildReport(', 1)
# SHED ตาม STATUS ต่อ B/L (opts.shedFor)
report = sub(report, 'computeExtraChecks(m, null, exShed)', 'computeExtraChecks(m, null, opts.shedFor ? opts.shedFor(m, null) : exShed)')
report = sub(report, 'computeExtraChecks(m, e, exShed)', 'computeExtraChecks(m, e, opts.shedFor ? opts.shedFor(m, e) : exShed)')
# ช่องหมายเหตุของแถวยอดรวม: build_edi.py ระบายแดงแต่ไม่เติม ⚠
report = sub(report, '      tcells[k] = { v: isBad ? `${WARN} ${T[k]}` : T[k], s: isBad ? "diff" : "ok", bad: isBad, total: true };',
             '      tcells[k] = { v: isBad && k !== "note" ? `${WARN} ${T[k]}` : T[k], s: isBad ? "diff" : "ok", bad: isBad, total: true };')

hahm = open(os.path.join(HERE, 'hahm_enter.js'), encoding='utf-8').read()
pdf = open(os.path.join(HERE, 'pdfpages.js'), encoding='utf-8').read()

head = '''/* CHECK EDI VARIOUS FORMS (HAHM) — ตรวจ MANIFEST (.xls) เทียบ ENTER (.pdf ชุดรวมหลาย forwarder ของ HAHM)
   พอร์ตจาก scripts/hahm_edi.py ทำงานในเบราว์เซอร์ล้วน ๆ ไม่มีการอัปโหลด
   ไฟล์นี้สร้างด้วย src/assemble.py: ตัวช่วยร่วม/parseManifest/buildReport มาจากเว็บ Check EDI Freight Links
   (github.com/Rattanao/Freight-links) + ตัวอ่าน ENTER ชุดรวม HAHM (src/hahm_enter.js) — แก้ที่ src/ แล้วรัน assemble ใหม่ */
(function (root) {
  "use strict";

'''
tail = '''
  root.CheckHAHM = { pdfPages, HEADERS, KEYS, NUMKEYS, WIDEKEYS, parseManifest, parseEnter, alignQtyPrefix, applyHahmRules,
    buildReport, annotateReport, shedByStatus, WARN, OK,
    _t: { canonPkg, canonStatus, normCo, comatch, sameDetail, extractTransitPhrase, subseqExtra, pageKind } };
})(typeof window !== "undefined" ? window : globalThis);
'''
out = head + common + '\n' + hahm + '\n' + report + '\n' + pdf + tail
os.makedirs(os.path.join(REPO, 'docs'), exist_ok=True)
open(os.path.join(REPO, 'docs', 'core.js'), 'w', encoding='utf-8', newline='\n').write(out)
print('docs/core.js', len(out.split(chr(10))), 'lines')
