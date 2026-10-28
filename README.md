# Check EDI Various Forms (HAHM)

โปรแกรมตรวจเอกสารขาเข้าเรือก่อนทำใบขนสินค้าขาเข้า — เทียบ **MANIFEST** (`.xls`/`.xlsx` ของสายเรือ)
กับ **ENTER** (`.pdf` ชุดรวมหลาย forwarder ของงาน HAHM) แล้วออกรายงาน **`EDI.xlsx`** แบบใส่สี ทีละ B/L:

- CONSIGNEE, CONTAINER NO., STATUS (CY / LCL / LCL-CFS), TAX ID vs NOTIFY
- TOTAL PACKAGE + ชนิดหีบห่อ, GROSS WEIGHT, MEASUREMENT
- MARKS, DESCRIPTION
- REEFER TEMP, DG (CLASS/UN), TRANSIT/TRANSHIPMENT (เจอเมื่อไหร่ขึ้นแดงทันที)
- SHED NO. ตาม STATUS (CFS/LCL = 0141, ที่เหลือ = 0122 — เปลี่ยนได้), ประเทศปลายทาง, CARGO MOVEMENT
- หมายเหตุทุก B/L บอกเลขหน้า ENTER ที่มาของค่า `[ENTER หน้า n]` เพื่อย้อนตรวจได้

**ความหมายของสี:** 🟥 แดง = ไม่ตรง · 🟧 ส้ม = ต้องให้คนตรวจยืนยัน / หน้าภาพสแกนต้องตรวจด้วยตา · 🟨 เหลือง = ไม่มี ENTER ให้เทียบ · 🟩 เขียว = ผ่าน

## ENTER ชุดรวมที่อ่านได้

ENTER ของ HAHM เป็น PDF ไฟล์เดียวรวมหลายฟอร์ม โปรแกรมแยกชนิดทีละหน้าแล้วอ่านด้วยตัวอ่านของแต่ละแบบ:

| ชนิดหน้า | อ่านอย่างไร |
|---|---|
| ARRIVAL NOTICE / สำเนา B/L ของ Heung-A (+ ATTACHED RIDER) | ตามตำแหน่งช่องบนฟอร์ม; MARKS ใน RIDER แทนคำ "AS PER ATTACHED RIDER" |
| ใบ MARK&NO. / NO. OF PKGS / DESCRIPTION OF GOODS | 2 B/L ต่อหน้า |
| ตาราง AMENDMENT (Cargoport / ASIA ATLANTIC / INITIAL / Handle ฯลฯ) | คอลัมน์จากหัวตาราง; STATUS/เลขตู้ที่ระบุครั้งเดียวใช้กับทุกแถวของฉบับ |
| จดหมายแก้ไขแบบข้อความอิสระ | ป้าย CNEE / CONTAINER / Status / MARKS … ; MARKS/DESC ไม่ตรง = ส้ม |
| MAC-NELS / THE M CORP (INBOUND CARGO MANIFEST) / DSV | ใช้ตัวอ่านเดียวกับเว็บ [MAC-NELS](https://rattanao.github.io/MAC-NELS/) / [CMS](https://rattanao.github.io/CMS/) / [DSV](https://rattanao.github.io/check-edi-DSV/) |
| หน้าภาพสแกน | อ่านอัตโนมัติไม่ได้ — รายงานขึ้นส้ม "ตรวจด้วยตา" พร้อมเลขหน้า |

ข้อความสีแดงที่ลูกค้าพิมพ์ทับบน PDF (FreeText) = คำสั่งจริงของลูกค้า (STATUS, DG, REEFER, แก้ชื่อผู้รับ ฯลฯ)

---

## ใช้ออนไลน์ (ง่ายสุด ไม่ต้องติดตั้ง)

เปิด **<https://rattanao.github.io/Various-forms-check-edi/>** → ลากไฟล์ MANIFEST + ENTER ใส่ → ตรวจ SHED NO. (ค่าเริ่มต้น 0141 / 0122) →
กด **ตรวจสอบ** → กด **ดาวน์โหลด EDI.xlsx**

ไฟล์ถูกประมวลผลในเบราว์เซอร์ของผู้ใช้เท่านั้น ไม่มีการอัปโหลดไปที่ใด

---

## ใช้แบบดับเบิลคลิก (ไม่ต้องมีอินเทอร์เน็ต)

1. ติดตั้ง **Python 3** จาก <https://www.python.org/downloads/> (ติ๊ก ☑ "Add python.exe to PATH")
2. หน้า repo กด **Code → Download ZIP** แล้วแตกไฟล์
3. ใส่ไฟล์ MANIFEST และ ENTER ในโฟลเดอร์ `input`
4. ดับเบิลคลิก **`run-check-edi.bat`** → กด Enter ใช้ SHED 0141/0122 หรือพิมพ์เลขใหม่
5. ผลอยู่ใน `output\EDI-HAHM.xlsx` (รอบถัดไปเป็น `EDI-HAHM-2.xlsx`, `-3` … ไม่ทับไฟล์เดิม)
   แล้วโปรแกรมรัน `verify_trace.py` ตรวจย้อนว่าทุกค่าในรายงานมาจากไฟล์จริง (ควรขึ้น "ไม่พบในไฟล์ต้นฉบับ 0")

## ใช้กับ Claude Code (skill)

คัดลอกโฟลเดอร์ repo นี้ไปเป็น `C:\Users\<ชื่อผู้ใช้>\.claude\skills\hahm-edi-check` (ต้องมี `SKILL.md` และ `scripts/`)
แล้วพิมพ์ "check-edi input" หรือ "ตรวจ EDI HAHM" ในโฟลเดอร์งาน — รายละเอียดใน [`SKILL.md`](SKILL.md)

---

## ไฟล์ใน repo

| ไฟล์ | หน้าที่ |
|---|---|
| `docs/index.html`, `docs/core.js` | หน้าเว็บ (GitHub Pages) |
| `docs/macnels.js`, `docs/cms.js`, `docs/dsv.js` | สำเนาตัวอ่านจากเว็บ MAC-NELS / CMS (Eastern Air) / DSV — ไม่แก้ไข |
| `src/` | ต้นฉบับของ `docs/core.js`: `hahm_enter.js` (ตัวอ่าน ENTER ชุดรวม), `pdfpages.js`, `assemble.py` (ประกอบ core.js จากเว็บ Freight Links + src) |
| `scripts/hahm_edi.py` | เวอร์ชัน Python (ตัวหลัก) |
| `scripts/build_edi.py` | ตัวตรวจร่วมของทุก forwarder (ใช้แบบอ่านอย่างเดียว) |
| `scripts/verify_trace.py` | ตรวจย้อนกลับว่าค่าในรายงานมาจากไฟล์ต้นฉบับจริง |

แก้ตัวอ่าน ENTER ที่ `src/hahm_enter.js` แล้วรัน `python src/assemble.py` เพื่อสร้าง `docs/core.js` ใหม่
ทุกครั้งที่แก้ ให้เทียบผลเว็บกับ `scripts/hahm_edi.py` บนงานจริงทีละช่อง (ทดสอบ 2026-10-28: 2,604 ช่องตรงกันทั้งค่าและสี
ต่างกันเฉพาะจำนวนช่องว่างซ้ำในข้อความ 11 ช่อง)

ห้าม commit ไฟล์งานของลูกค้า — `.gitignore` กันโฟลเดอร์ `input/`, `output/` และ `EDI*.xlsx` ไว้แล้ว
