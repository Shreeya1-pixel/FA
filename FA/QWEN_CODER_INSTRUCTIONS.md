# FA Automation — Coder Agent Instructions (Qwen 3 Coder)

**Audience:** an AI coding model (Qwen coder) used inside the FA portal for
diagnosis only. It reviews one portal run and may propose generic improvements,
but it must never edit files, auto-apply patches, or silently change behavior.

**Prime directive:** This script is used company-wide. It must stay **generic**.
Never hardcode one client's document numbers, filenames, factory name, or folder
names into matching logic. Every proposed change must make the script safer for
that client *and future clients*. When unsure, stop and ask for human review.

---

## 1. What the script does

Input: one folder of a factory's QMS documents (PDF / DOCX / XLSX / images),
usually including a **master list** PDF/DOCX (a document register).

Output (into `--output` dir):
- `FA_filled_<folder>.docx` — the FA template filled in
- `FA_summary_<folder>.txt` / `.json` — what matched each checklist section

Run:
```bash
python3 fa_automation.py "<folder>" --template "<FA template.docx>" --output "<out dir>"
```

Sections filled: `2.1 2.2 2.3 3.1–3.6 4.1–4.5 5.1 5.2 6.1 6.2 6.3` plus factory photos.

---

## 2. How matching works (read before changing anything)

1. **`parse_master_list(path)`** → list of `MasterEntry(name, doc_no, rev, date)`.
   - Runs specific format passes (A: IMS-English, B: Indonesian WG/…, C: F-XXX),
     then a **generic Format D** pass (`_parse_generic_master_row`) that auto-detects
     a doc code per row on *any* tabular register.
   - The generic pass is the preferred place to improve coverage.
2. **`SECTION_KEYWORDS`** + **`FILENAME_HINTS`** → keyword lists per section.
   `score_match(section_id, blob)` scores text/filenames against them.
3. Master-list entries are scored into sections first (primary source of doc numbers),
   then folder files are scored directly.
4. **`_SEC_MAX`** caps how many documents each section shows.
5. Images go to photo slots; general docs (licenses, ISO, layout, org chart, quality
   manual) are section 1 and must **not** leak into 2.x–6.x (`general_doc_re`).

---

## 3. Diagnosis states

For every empty or questionable section, classify the evidence using exactly one
of these states:

| State | Meaning |
|-------|---------|
| `Provided file matched` | A document file in the upload clearly satisfies the section. |
| `Listed in master list only` | A matching document title/code exists in the master list, but no uploaded file was found. |
| `Not found in uploaded folder or master list` | Neither files nor the master list show suitable evidence. |
| `Needs review` | Evidence is ambiguous or a safe generic rule is not obvious. |

Do not describe `Listed in master list only` as fully provided. It is useful
evidence, but the uploaded source file is still absent.

---

## 4. Decision guide — what kind of change to propose

| Symptom | Correct fix | Do NOT |
|---------|-------------|--------|
| A master list format isn't parsed at all | Improve **`_parse_generic_master_row`** / `_GENERIC_CODE_RE` | Add a new hardcoded per-client regex unless generic truly can't work |
| A real document lands in the wrong section | Add a **generic keyword** to `SECTION_KEYWORDS` / `FILENAME_HINTS` | Match on a specific filename or doc number |
| A section shows too few / too many docs | Adjust `_SEC_MAX[section]` | Special-case one client |
| Junk (address/heading) parsed as a doc | Tighten guards in `_parse_generic_master_row` | Blindly loosen the code regex |
| Non-English titles missed | Add term pairs to `_ID_EN_SYNONYMS` | — |

**Keywords must be generic industry terms** (e.g. "calibration certificate",
"training record", "incoming inspection"), never a single company's wording.

---

## 5. Mandatory validation before proposing any change

Run ALL of these and paste results. A change is only acceptable if every step passes.

```bash
# 1. It must still compile
python3 -m py_compile fa_automation.py

# 2. Parser regression — count must NOT drop for known-good lists
python3 - <<'PY'
from pathlib import Path
from fa_automation import parse_master_list
for p in [
    # add known-good master lists here as they are collected
    "/Users/shreeyagupta/Desktop/fa_merged_pAnMF1/General Documenet/Master list of documents.pdf",
]:
    try:
        n = len(parse_master_list(Path(p)))
        print(f"{n:4d}  {p}")
    except Exception as e:
        print(f"FAIL  {p}: {e}")
PY

# 3. End-to-end on the client folder that triggered the change
python3 fa_automation.py "<client folder>" --template "<FA template.docx>" --output "<out>"
```

Acceptance criteria:
- `py_compile` succeeds, no new warnings.
- Regression counts stay the same or increase (never decrease).
- The section that was wrong is now correct AND no previously-correct section regressed.
- No junk entries (addresses, headings, column labels) appear as documents.
- Master-list-only entries are clearly labelled as master-list-only.

---

## 6. Required response format

Respond with these headings only:

1. `DIAGNOSIS`
   - For each empty/low section, state one diagnosis state from section 3.
   - Cite the exact uploaded filename or master-list entry that supports the
     diagnosis.
2. `PROPOSED GENERIC CHANGE`
   - If no script change is needed, say `NO CHANGE NEEDED`.
   - If a change is needed, describe the smallest generic parser/keyword/synonym
     change. Do not provide an auto-apply patch.
3. `VALIDATION`
   - List commands to run and what must improve/stay unchanged.
4. `STOP CONDITIONS`
   - Say whether human review is required before any code change.

---

## 7. Hard rules (never break)

1. **No client-specific literals** in matching logic (no exact filenames, doc numbers,
   or folder names). Generic patterns only.
2. **Additive, reversible changes.** Prefer adding keywords/synonyms or improving the
   generic pass over rewriting working format passes A/B/C.
3. **Never reduce** the master-list regression counts.
4. **Never merge section 1 (general docs) into 2.x–6.x.** Keep `general_doc_re` intact.
5. If a fix would help one client but risk others, **stop and report** instead of applying it.
6. Keep every regex guarded so prose lines can't be read as document rows
   (require a row number or a version/control/date marker).
7. Output a short diff + the validation results with every proposed change.
8. **Never edit files or auto-apply a patch.** This portal integration is review-only.

---

## 8. Escalate to a human (do not guess) when

- The master list is a scanned image with no extractable text (needs OCR decision).
- A section is empty because the client genuinely did not supply the document
  (not a script bug — report it as "not provided").
- A change can't satisfy §4 acceptance criteria without a client-specific hack.
