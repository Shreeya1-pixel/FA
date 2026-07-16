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
        if getattr(match, "status", "") == "master_list_only"
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


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate FA automation regressions")
    parser.add_argument("folders", nargs="*", type=Path, help="FA folders to validate")
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument("--max-folders", type=int, default=6)
    parser.add_argument("--skip-e2e", action="store_true")
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

    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
