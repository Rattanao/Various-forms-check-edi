"""ตรวจย้อนกลับ: ค่าทุกค่าที่รายงานใช้ หาเจอจริงในไฟล์ต้นฉบับหรือไม่
   ENTER  -> ค้นในข้อความ/annotation ของหน้า PDF ที่ระบุ
   MANIFEST -> ค้นในแถวดิบของ .xls ช่วงของ B/L นั้น"""
import sys, re, importlib.util
sys.stdout.reconfigure(encoding='utf-8')
import os
spec = importlib.util.spec_from_file_location('h', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'hahm_edi.py'))
h = importlib.util.module_from_spec(spec); spec.loader.exec_module(h)
import pymupdf, xlrd

man_path, ent_path = sys.argv[1], sys.argv[2]
MAN, vessel, decl, bl_re = h.be.parse_manifest(man_path)
ENT, info = h.parse_enter_hahm(ent_path, bl_re)

def nums(text):
    return {round(float(x.replace(',', '')), 3) for x in re.findall(r'\d[\d,]*\.?\d*', text)}

def alnum(s):
    return re.sub(r'[^A-Z0-9]', '', str(s or '').upper())

# ---- ENTER: ข้อความดิบต่อหน้า (ทั้งตัวพิมพ์และ FreeText annotation)
doc = pymupdf.open(ent_path)
ptext = {}
for i, pg in enumerate(doc, 1):
    t = pg.get_text() + ' ' + ' '.join((a.info.get('content') or '') for a in (pg.annots() or []))
    ptext[i] = t

# ---- MANIFEST: แถวดิบของแต่ละ B/L (จากบรรทัดเลข B/L ถึงก่อนเลข B/L ถัดไป)
sh = xlrd.open_workbook(man_path, logfile=open(os.devnull, 'w')).sheet_by_index(0)
rows = [' | '.join(str(sh.cell_value(r, c)) for c in range(sh.ncols)) for r in range(sh.nrows)]
starts = [(r, str(sh.cell_value(r, 0)).strip()) for r in range(sh.nrows)
          if re.fullmatch(r'HASL[A-Z0-9]+', str(sh.cell_value(r, 0)).strip())]
mraw = {}
for k, (r, bl) in enumerate(starts):
    end = starts[k + 1][0] if k + 1 < len(starts) else sh.nrows
    mraw[bl] = ' '.join(rows[r:end])

bad, checked, src_count = [], 0, {}
for bl in MAN:
    m = MAN[bl]
    raw = mraw.get(bl, '')
    # MANIFEST side
    for f in ('gw', 'meas'):
        v = h.be.num(m[f])
        if v is not None:
            checked += 1
            if round(v, 3) not in nums(raw):
                bad.append(f'{bl} MANIFEST {f}={v} ไม่พบในแถว .xls')
    for c in m['cont']:
        checked += 1
        if c not in raw:
            bad.append(f'{bl} MANIFEST cont {c} ไม่พบในแถว .xls')
    checked += 1
    if alnum(m['cons'])[:15] not in alnum(raw):
        bad.append(f'{bl} MANIFEST cons ไม่พบ')
    # ENTER side
    e = ENT.get(bl)
    if not e:
        continue
    pages = e.get('pages') or []
    text = ' '.join(ptext[p] for p in pages)
    src_count[bl] = pages
    for f in ('gw', 'meas', 'pkgs'):
        v = e.get(f)
        if v is not None:
            checked += 1
            if round(float(v), 3) not in nums(text):
                bad.append(f'{bl} ENTER {f}={v} ไม่พบในหน้า {pages}')
    for c in h.conts_in(e.get('cont') or ''):
        checked += 1
        if c not in text.upper():
            bad.append(f'{bl} ENTER cont {c} ไม่พบในหน้า {pages}')
    if e.get('cons'):
        checked += 1
        if alnum(e['cons'])[:15] not in alnum(text):
            bad.append(f'{bl} ENTER cons "{e["cons"]}" ไม่พบในหน้า {pages}')
    if e.get('status_raw'):
        checked += 1
        st = re.sub(r'\s+', '', e['status_raw'])[:6]
        if st not in re.sub(r'\s+', '', text).replace('ลากตต', 'ลากตู้'):
            bad.append(f'{bl} ENTER status "{e["status_raw"]}" ไม่พบในหน้า {pages}')

print(f'MANIFEST B/L: {len(MAN)} | อ่านค่าจาก ENTER ได้: {sum(1 for b in MAN if b in ENT)} | ไม่มีข้อความ ENTER: '
      f'{", ".join(b for b in MAN if b not in ENT)}')
print(f'ENTER B/L ที่ไม่อยู่ใน MANIFEST: {[b for b in ENT if b not in MAN]}')
print(f'ตรวจย้อนกลับ {checked} ค่า | ไม่พบในไฟล์ต้นฉบับ {len(bad)}')
for b in bad:
    print('  -', b)
print('\nตัวอย่างที่มาของค่า:')
for bl in [b for b in MAN if b in ENT][:5]:
    e, m = ENT[bl], MAN[bl]
    print(f'{bl}: ENTER หน้า {e["pages"]} -> {e["cons"]} | {e["cont"]} | {e["pkgs"]} {e["pkgtype"]} | {e["gw"]} kg | {e["meas"]} cbm')
    print(f'{"":17} MANIFEST .xls  -> {m["cons"]} | {";".join(m["cont"])} | {m["pkg_hdr"]} | {m["gw"]} | {m["meas"]}')
