---
name: hahm-edi-check
description: "Check EDI for HAHM: reconcile a carrier MANIFEST (.xls) against an ENTER PDF that is a bundle of many forwarders' forms (Heung-A arrival notices / B/L copies + ATTACHED RIDER, MARK&NO. sheets, AMENDMENT tables from Cargoport / ASIA ATLANTIC / INITIAL / Handle, free-text amendment letters, MAC-NELS, THE M CORP/CMS, DSV, scanned pages) and write the colour-coded EDI.xlsx report. Use whenever the user says \"check-edi input\" for a HAHM job, \"ตรวจ EDI HAHM\", \"เทียบ MANIFEST กับ ENTER ชุดรวม\", or points to a job folder with input/MANIFEST-HAHM*.xls and input/ENTER-HAHM*.pdf. Separate from manifest-enter-edi-check: never edit that skill or other forwarders' checks for HAHM."
---

# HAHM — MANIFEST × ENTER (ชุดรวม) → EDI.xlsx

HAHM's ENTER is one PDF bundling ~15 different forms. The shared `manifest-enter-edi-check` skill reads
a whole ENTER as one layout, so on HAHM it misreads most B/Ls. This skill splits the bundle by page type
and reads each type with the right reader, then builds the same EDI.xlsx report.

**Rule from the user (2026-10-26):** HAHM has its own format. Do not change or upload anything for the
other forwarders' EDI checks (the `manifest-enter-edi-check` skill, its forwarder websites/repos).
Fix HAHM problems only in this skill's `scripts/hahm_edi.py`.

## Files

- `scripts/hahm_edi.py` — the HAHM reader + report post-processing.
- `scripts/build_edi.py` — a frozen copy of the shared skill's script (used read-only: `parse_manifest`,
  `build_report`, and its MAC-NELS / CMS / DSV page parsers). `hahm_edi.py` loads the copy next to it
  first, so this folder works on its own on another PC.

Requirements: Python 3.10+, `pip install pymupdf openpyxl xlrd`.

## Run

From the job folder (the one that contains `input/`):

```bash
python ~/.claude/skills/hahm-edi-check/scripts/hahm_edi.py --shed-lcl 0141 --shed-other 0122 --name EDI-HAHM.xlsx
```

- Finds `input/MANIFEST*.xls` and `input/ENTER*.pdf` automatically (or pass `--manifest` / `--enter`).
- Writes `output/EDI.xlsx` (or `--outdir <folder>`). If `output/EDI.xlsx` is open in Excel the save
  fails with `PermissionError` — write to another `--outdir`, then ask the user before replacing; the
  user keeps old reports (do not overwrite without asking).
- SHED NO.: HAHM's usual rule (user, 2026-10-26) is **STATUS CFS/LCL → 0141, the rest → 0122**:
  `--shed-lcl 0141 --shed-other 0122`. The STATUS used is the customer's instruction in ENTER when there
  is one, otherwise MANIFEST's STATUS (so a B/L whose ENTER says เปิดตู้ but MANIFEST says CY is checked
  against 0141 and flagged). `--shed 0141` = one SHED for every B/L. Ask the user if no SHED was given.
- `--name EDI-HAHM-2.xlsx` sets the report file name. The script refuses to overwrite an existing file
  (the user keeps every report).

**Prove the comparison is real (trace check):** after a run, also run

```bash
python ~/.claude/skills/hahm-edi-check/scripts/verify_trace.py input/MANIFEST-xxx.xls input/ENTER-xxx.pdf
```

It looks up every value the report uses again in the raw files — ENTER values (G.W., MEAS., packages,
containers, consignee, STATUS) on the ENTER page(s) listed for that B/L, MANIFEST values in the B/L's
raw .xls rows — and prints how many were not found (should be 0), plus which B/Ls have no ENTER text.
Values stated once per document (a sheet-top STATUS note, a Cargoport header container) list that page
too, so they trace. 2026-10-26 HAHM job: 971 values, 0 not found.

Relay the console summary to the user, DG / REEFER / STATUS first, then MEAS. 0.000 in MANIFEST,
then MARKS/DESC, then B/Ls with no ENTER text (scanned pages).

## How the ENTER bundle is read (page types)

| Type | Recognised by | Reader |
|---|---|---|
| BLFORM | `SHIPPER'S LOAD` (Heung-A ARRIVAL NOTICE / B/L COPY, also no-space text) | own, coordinates relative to the SHIPPER'S LOAD row |
| RIDER | `== ATTACHED RIDER ==` | own; its marks replace `AS PER ATTACHED RIDER` |
| MARKNO | `MARK&NO.` + `CNTR :` + `B/L NO. :` (2 B/L per page) | own |
| TABLE | header row CONSIGNEE / QUANTITY / DESCRIPTION / GROSS WEIGHT | own; columns = header x − 45; split B/L (`…104` + `9`) joined; STATUS/container stated once per document apply to all rows |
| LETTER | anything else with text | own, label/value (`CNEE`, `CONTAINER NO`, `Status`, …) or a small MARKS/PKGS/DESCRIPTION table; MARKS/DESC mismatches orange |
| SHARED_MACNELS / CMS / DSV | MAC-NELS ATTACHED SHEETS / INBOUND CARGO MANIFEST / AMENDMENT OF MANIFEST + B/L CHANGED | only those pages are sent to `build_edi.parse_enter` |
| SCAN | no text layer | cannot be read (no OCR); flagged for checking by eye |

HAHM-specific rules (in `apply_hahm_rules` / `annotate_report`):

- Containers use ISO 6346 `[A-Z]{3}[UJZ]\d{7}` so SHIPMENT IDs (`SSGN0000001`) and B/L tails are not taken.
- Lines drawn by a FreeText annotation are recognised by its rectangle, not by text, so a real marks line
  that repeats the annotation text is kept.
- Customer STATUS is any typed line/annotation line with STATUS / ลากตู้ / เปิดตู้ / ขนส่ง / CFS (only that
  line of a multi-line note). A FreeText STATUS at the top of a shared sheet applies to all its B/Ls.
  A cover letter "CONSIGNEE AS PER ATTACHED SHEET" passes its STATUS to all sub-B/Ls.
- B/L consignee `TO ORDER OF …` → the Notify name is compared (noted in the report).
- BLFORM only: if the MANIFEST MARKS/DESC is a part of the B/L text (in order; `*AS PER ATTACHED SHEET*`
  ignored), it counts as a match; the ENTER column still shows the full B/L text and the note lists the
  extra text. Amendment forms (DSV, tables, letters) stay strict.
- `REEL` = ROLL for packaging. MANIFEST `TEMP : - 18'C` / `MINUS 18 DEGREES` is read as -18°C.
- The note column starts with `[ENTER หน้า n]` for every B/L. B/Ls with no ENTER text get an orange note
  listing the scanned pages; when a scanned page's typed container matches the MANIFEST it names that page.
- ENTER has no single grand total (bundle), so the TOTAL row compares against MANIFEST only.

## When a new HAHM bundle has a form this does not read

Dump each page's lines with x/y and its annotations (`page.get_text('dict')`, `page.annots()`), see which
type it falls into (`page_kind`), and add or adjust a reader in `hahm_edi.py`. Rerun the last HAHM job
before and after and compare every report cell; differences must be explained by the change.
