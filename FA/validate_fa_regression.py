#!/usr/bin/env python3
"""Regression checks for FA automation.

This is intentionally lightweight and local: it validates that known FA folders
still parse their master lists and can run end-to-end without requiring portal
state. It is safe to run before accepting any generic parser/keyword change.
"""

from __future__ import annotations

import argparse
import py_compile
import re
import subprocess
import sys
import tempfile
from pathlib import Path

FA_DIR = Path(__file__).resolve().parent
ROOT = FA_DIR.parent
DEFAULT_TEMPLATE = Path.home() / "Desktop" / "FA template.docx"

if str(FA_DIR) not in sys.path:
    sys.path.insert(0, str(FA_DIR))

import fa_automation as fa  # noqa: E402


JUNK_NAME_RE = re.compile(
    r"^(?:sl|id|version|report\s+name|documents?\s+control|"
    r"book\s*/?\s*computer|process\s+area|page)$",
    re.I,
)


def _has_supported_files(folder: Path) -> bool:
    suffixes = {".pdf", ".docx", ".xlsx", ".xls", ".jpg", ".jpeg", ".png", ".webp"}
    return any(p.is_file() and p.suffix.lower() in suffixes for p in folder.rglob("*"))


def _latest_matching_dirs(base: Path, pattern: str, limit: int) -> list[Path]:
    if not base.exists():
        return []
    dirs = [p for p in base.glob(pattern) if p.is_dir()]
    dirs.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return dirs[:limit]


def discover_folders(limit: int) -> list[Path]:
    """Find recent local FA sample folders without hardcoding client names."""
    candidates: list[Path] = []
    candidates.extend(_latest_matching_dirs(Path.home() / "Desktop", "fa_merged_*", limit))
    candidates.extend(_latest_matching_dirs(ROOT / "portal" / "fa_work", "run_*", limit))
    candidates.extend(_latest_matching_dirs(ROOT / "fa_portal" / "runs", "run_*", limit))

    usable: list[Path] = []
    seen: set[Path] = set()
    for folder in candidates:
        resolved = folder.resolve()
        if resolved in seen or not _has_supported_files(resolved):
            continue
        seen.add(resolved)
        usable.append(resolved)
    return usable[:limit]


def validate_master_list(folder: Path) -> tuple[bool, str]:
    master = fa.find_master_list(folder)
    if not master:
        return True, f"NO MASTER LIST  {folder}"
    try:
        entries = fa.parse_master_list(master)
    except Exception as exc:
        return False, f"FAIL PARSE      {master}: {exc}"
    if not entries:
        return False, f"FAIL EMPTY      {master}: 0 entries parsed"

    junk = [
        entry
        for entry in entries
        if JUNK_NAME_RE.search(entry.name.strip()) or JUNK_NAME_RE.search(entry.doc_no.strip())
    ]
    if junk:
        preview = ", ".join(f"{e.doc_no}:{e.name}" for e in junk[:5])
        return False, f"FAIL JUNK       {master}: {preview}"

    return True, f"OK PARSE        {len(entries):4d} entries  {master}"


def validate_match(folder: Path) -> tuple[bool, str]:
    try:
        result = fa.match_documents(folder)
    except Exception as exc:
        return False, f"FAIL MATCH      {folder}: {exc}"
    total = sum(len(matches) for matches in result.sections.values())
    master_only = sum(
        1
        for matches in result.sections.values()
        for match in matches
        if getattr(match, "status", "") in {"master_list_only", "register_only"}
    )
    return True, f"OK MATCH        {total:3d} section docs ({master_only} master-list-only)  {folder}"


def validate_end_to_end(folder: Path, template: Path) -> tuple[bool, str]:
    if not template.exists():
        return True, f"SKIP E2E        template not found: {template}"
    with tempfile.TemporaryDirectory(prefix="fa_regression_") as tmp:
        cmd = [
            sys.executable,
            str(FA_DIR / "fa_automation.py"),
            str(folder),
            "--template",
            str(template),
            "--output",
            tmp,
        ]
        proc = subprocess.run(
            cmd,
            cwd=str(FA_DIR),
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            timeout=180,
        )
    if proc.returncode != 0:
        return False, f"FAIL E2E        {folder}\n{proc.stdout}"
    return True, f"OK E2E          {folder}"


def _find_fa_mh_folder() -> Path | None:
    """Locate the FA MH sample folder used for Armourcoat-style register checks."""
    candidates = [
        ROOT / "portal" / "fa_work" / "run_FA_MH_1784015938",
        Path.home() / "Downloads" / "FA MH",
    ]
    for path in candidates:
        if path.is_dir() and _has_supported_files(path):
            return path.resolve()
    for path in _latest_matching_dirs(ROOT / "portal" / "fa_work", "run_FA_MH_*", 3):
        if _has_supported_files(path):
            return path.resolve()
    return None


def validate_fa_mh_expectations(folder: Path | None = None) -> tuple[bool, str]:
    """Regression expectations for FA MH-style packages (DCR + content extraction)."""
    folder = folder or _find_fa_mh_folder()
    if folder is None:
        return True, "SKIP FA_MH      no FA MH sample folder found"

    # DCR / register parsing
    registers = fa.find_document_registers(folder)
    dcr = next((p for p in registers if re.search(r"\bdcr\b", p.name, re.I)), None)
    if dcr is None:
        return False, f"FAIL FA_MH      DCR register not discovered in {folder}"
    try:
        entries = fa.parse_master_list(dcr)
    except Exception as exc:
        return False, f"FAIL FA_MH      DCR parse error: {exc}"
    doc_nos = {e.doc_no.upper() for e in entries}
    required_codes = {"WS-TT", "C-001", "P-001"}
    # WS-PL-S may normalize as WS-PL-S or WS-PLS depending on cell text
    has_pls = any(code.startswith("WS-PL") for code in doc_nos) or "WS-PLS" in doc_nos
    missing = sorted(required_codes - doc_nos)
    if missing or not has_pls:
        return False, f"FAIL FA_MH      DCR missing codes {missing} pls={has_pls} got={sorted(doc_nos)[:12]}"

    result = fa.match_documents(folder)

    def _blob(sid: str) -> str:
        return " ".join(
            f"{m.doc_no} {m.heading} {m.filename}".lower()
            for m in result.sections.get(sid, [])
        )

    checks: list[tuple[bool, str]] = []

    # Work instructions should land in production control, not vanish
    b34 = _blob("3.4")
    checks.append(("operator" in b34 or "worksheet" in b34 or "ws-" in b34, "3.4 work instructions"))

    # Calibration worksheets -> 4.4, not treated as certificate SOP
    b41 = _blob("4.1")
    b44 = _blob("4.4")
    checks.append(("check" in b44 or "worksheet" in b44 or "calibration" in b44, "4.4 calibration records"))
    checks.append((not b41 or "certificate" in b41, "4.1 not worksheet-as-cert"))

    # Finished goods / QC records should support 6.1 or 3.6
    b61 = _blob("6.1")
    b36 = _blob("3.6")
    checks.append(
        (
            any(k in b61 or k in b36 for k in ("qc", "sampling", "inspection testing", "aquawax", "test")),
            "6.1/3.6 QC or test records",
        )
    )

    # Customer complaint SOP must remain absent unless a real file exists
    b63 = _blob("6.3")
    checks.append(("quality manual" not in b63, "6.3 must not use QM narrative as SOP"))

    # Optional: content extraction from operator worksheet PDF
    wi = next(folder.rglob("Operator work instruction sheet Tactite*.pdf"), None)
    if wi and wi.exists():
        meta = fa.extract_content_meta(wi)
        checks.append(
            (
                "WS-TT" in (meta.get("doc_no") or "").upper() or "tactite" in (meta.get("title") or "").lower(),
                "content extract WS-TT",
            )
        )

    # Model review payload builds without inventing matches
    payload = fa.build_model_review_payload(folder, result)
    checks.append(("low_confidence_sections" in payload and "files" in payload, "model review payload"))
    before = {sid: [(m.doc_no, m.filename) for m in result.sections.get(sid, [])] for sid in fa.SECTION_IDS}
    fake = [
        {
            "section_id": "6.3",
            "doc_no": "FAKE-COMPLAINT-999",
            "name": "Invented Complaint SOP",
            "source_file": "does-not-exist.pdf",
            "status": "provided_file",
            "confidence": 0.99,
            "reason": "should be rejected",
        }
    ]
    reviewed = fa.apply_model_review_proposals(result, fake, folder)
    after = {sid: [(m.doc_no, m.filename) for m in reviewed.sections.get(sid, [])] for sid in fa.SECTION_IDS}
    checks.append((before == after, "model review rejects invented docs"))

    failed = [label for ok, label in checks if not ok]
    if failed:
        preview = {sid: _blob(sid)[:120] for sid in ("3.4", "4.1", "4.4", "6.1", "6.3")}
        return False, f"FAIL FA_MH      {failed} preview={preview}"
    return True, f"OK FA_MH        register+match expectations passed  {folder}"


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate FA automation regressions")
    parser.add_argument("folders", nargs="*", type=Path, help="FA folders to validate")
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument("--max-folders", type=int, default=6)
    parser.add_argument("--skip-e2e", action="store_true")
    parser.add_argument("--skip-fa-mh", action="store_true")
    args = parser.parse_args()

    ok = True
    py_compile.compile(str(FA_DIR / "fa_automation.py"), doraise=True)
    print(f"OK COMPILE      {FA_DIR / 'fa_automation.py'}")

    folders = [p.resolve() for p in args.folders] if args.folders else discover_folders(args.max_folders)
    if not folders:
        print("FAIL            no FA sample folders found; pass one explicitly")
        return 1

    for folder in folders:
        for check in (validate_master_list, validate_match):
            passed, message = check(folder)
            print(message)
            ok = ok and passed
        if not args.skip_e2e:
            passed, message = validate_end_to_end(folder, args.template)
            print(message)
            ok = ok and passed

    if not args.skip_fa_mh:
        passed, message = validate_fa_mh_expectations()
        print(message)
        ok = ok and passed

    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
