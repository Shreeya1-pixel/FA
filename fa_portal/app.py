"""
FA Report Generator Portal — TÜV Rheinland
==========================================
Standalone Streamlit app for generating Factory Audit reports.
Upload factory documents → auto-fill the FA template → download the report.

Run with:
    streamlit run fa_portal/app.py
"""

from __future__ import annotations

import importlib
import io
import json
import sys
import time
import traceback
import zipfile
from datetime import datetime
from pathlib import Path

import streamlit as st

# ── paths ─────────────────────────────────────────────────────────────────────
PORTAL_DIR = Path(__file__).resolve().parent
ROOT = PORTAL_DIR.parent
FA_DIR = ROOT / "FA"
RUNS_DIR = PORTAL_DIR / "runs"
RUNS_DIR.mkdir(exist_ok=True)
SETTINGS_FILE = PORTAL_DIR / "fa_settings.json"
LOG_FILE = PORTAL_DIR / "fa_feedback.jsonl"
PYTHON = ROOT / ".venv" / "bin" / "python"
if not PYTHON.exists():
    PYTHON = Path(sys.executable)

for p in (str(FA_DIR), str(ROOT)):
    if p not in sys.path:
        sys.path.insert(0, p)

# ── load fa_automation ────────────────────────────────────────────────────────
try:
    import fa_automation as fa_mod  # type: ignore[import-not-found]
    importlib.reload(fa_mod)
    FA_AVAILABLE = True
except Exception as _e:
    FA_AVAILABLE = False
    _fa_import_error = str(_e)

# ── default settings ──────────────────────────────────────────────────────────
DEFAULT_SETTINGS: dict[str, str] = {
    "template_path": str(Path.home() / "Desktop" / "FA template.docx"),
    "output_dir": str(RUNS_DIR),
    "default_country": "KSA",
    "default_auditor": "",
    "qwen_model": "qwen2.5-coder:7b",
    "ollama_host": "http://localhost:11434",
}


def load_settings() -> dict[str, str]:
    if SETTINGS_FILE.exists():
        try:
            return {**DEFAULT_SETTINGS, **json.loads(SETTINGS_FILE.read_text())}
        except Exception:
            pass
    return dict(DEFAULT_SETTINGS)


def save_settings(s: dict[str, str]) -> None:
    SETTINGS_FILE.write_text(json.dumps(s, indent=2))


def load_runs() -> list[dict]:
    runs: list[dict] = []
    if not LOG_FILE.exists():
        return runs
    for line in LOG_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line:
            try:
                runs.append(json.loads(line))
            except Exception:
                pass
    return list(reversed(runs))


def append_log(entry: dict) -> None:
    with LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(json.dumps(entry) + "\n")


def extract_zip(uploaded_file, dest: Path) -> Path:
    """Extract uploaded ZIP into dest, skipping __MACOSX junk."""
    dest.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(io.BytesIO(uploaded_file.read())) as zin:
        for member in zin.namelist():
            if "__MACOSX" in member or member.endswith("/"):
                continue
            rel = Path(member)
            if not rel.name:
                continue
            parts = rel.parts
            out = dest.joinpath(*parts[1:]) if len(parts) > 1 else dest / rel.name
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_bytes(zin.read(member))

    subdirs = [p for p in dest.iterdir() if p.is_dir()]
    files = [p for p in dest.iterdir() if p.is_file()]
    if len(subdirs) == 1 and not files:
        return subdirs[0]
    return dest


def _safe_dest(base: Path, raw_name: str) -> Path:
    """Return a path inside base, ignoring unsafe absolute/parent segments."""
    parts = [
        part
        for part in Path(raw_name).parts
        if part not in {"", ".", ".."} and not Path(part).is_absolute()
    ]
    if not parts:
        parts = ["uploaded_file"]
    return base.joinpath(*parts)


def extract_uploaded_documents(uploaded_files: list, dest: Path) -> Path:
    """Save many uploaded files and unpack any ZIPs into one FA documents folder."""
    dest.mkdir(parents=True, exist_ok=True)
    for uploaded in uploaded_files:
        name = Path(uploaded.name).name
        data = uploaded.getvalue()
        if name.lower().endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(bytes(data))) as zin:
                for member in zin.namelist():
                    if "__MACOSX" in member or member.endswith("/"):
                        continue
                    rel = Path(member)
                    if not rel.name or rel.name.lower() in {".ds_store", "thumbs.db"}:
                        continue
                    parts = rel.parts[1:] if len(rel.parts) > 1 else rel.parts
                    out = _safe_dest(dest, "/".join(parts))
                    out.parent.mkdir(parents=True, exist_ok=True)
                    out.write_bytes(zin.read(member))
            continue

        if name.lower() in {".ds_store", "thumbs.db"}:
            continue
        out = _safe_dest(dest, uploaded.name)
        if out.exists():
            stem, suffix = out.stem, out.suffix
            idx = 2
            while out.exists():
                out = out.with_name(f"{stem}_{idx}{suffix}")
                idx += 1
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_bytes(bytes(data))

    subdirs = [p for p in dest.iterdir() if p.is_dir()]
    files = [p for p in dest.iterdir() if p.is_file()]
    if len(subdirs) == 1 and not files:
        return subdirs[0]
    return dest


# ── core pipeline (shared by first run and re-run) ─────────────────────────────
def generate_fa_report(
    docs_folder: Path,
    run_dir: Path,
    ts: int,
    template_path: Path,
    auditor_names: list[str] | None,
    audit_date_input: str,
    country_input: str,
    factory_name_override: str,
    progress=None,
):
    """Run the full FA matching + template fill. Returns (result, out_docx).

    Kept as a single function so the 'Generate' and 'Re-run' buttons share
    identical logic — a re-run must reproduce the same steps on the same docs.
    """
    def _p(pct, txt):
        if progress is not None:
            progress.progress(pct, text=txt)

    _p(10, "Matching documents…")
    result = fa_mod.match_documents(docs_folder)

    _p(20, "Matching images…")
    image_matches = fa_mod.match_images(docs_folder)
    result.images = image_matches

    _p(35, "Parsing factory info…")
    result.factory_info = fa_mod.parse_factory_info(docs_folder)
    if factory_name_override.strip():
        result.factory_info["name"] = factory_name_override.strip()

    _p(50, "Extracting audit metadata…")
    result.audit_meta = fa_mod.extract_audit_meta(
        docs_folder, result.factory_info, auditor_names
    )
    if audit_date_input.strip():
        result.audit_meta["audit_date"] = audit_date_input.strip()
    if country_input.strip():
        result.audit_meta["country"] = country_input.strip().upper()

    rn = fa_mod.generate_report_number(
        result.audit_meta.get("country", ""),
        result.audit_meta.get("auditor_names", []),
        result.audit_meta.get("audit_date", ""),
    )
    result.audit_meta["report_number"] = rn

    _p(65, "Building checklist…")
    master_path = Path(result.master_list) if result.master_list else None
    result.checklist = fa_mod.build_checklist(docs_folder, image_matches, master_path)
    fa_mod.write_checklist(result.checklist, run_dir / "FA_Checklist.txt")
    result.checklist_path = str(run_dir / "FA_Checklist.txt")

    _p(80, "Filling template…")
    out_docx = run_dir / f"FA_Report_{ts}.docx"
    fa_mod.fill_template(
        template_path, result, out_docx, folder=docs_folder, image_matches=image_matches
    )
    result.output_docx = str(out_docx)

    _p(95, "Writing summary…")
    summary_path = run_dir / "FA_summary.json"
    summary_path.write_text(json.dumps(result.to_dict(), indent=2), encoding="utf-8")
    result.summary_path = str(summary_path)
    _p(100, "Done!")
    return result, out_docx


def reload_fa_module():
    """Hot-reload fa_automation so script edits take effect without restarting."""
    global fa_mod
    importlib.reload(fa_mod)
    return fa_mod


# ── Qwen coder integration (local Ollama, propose-only) ────────────────────────
QWEN_INSTRUCTIONS_FILE = FA_DIR / "QWEN_CODER_INSTRUCTIONS.md"


def _ollama_available(host: str) -> bool:
    import urllib.request
    try:
        with urllib.request.urlopen(f"{host}/api/tags", timeout=2) as r:
            return r.status == 200
    except Exception:
        return False


def _ollama_models(host: str) -> list[str]:
    import urllib.request
    try:
        with urllib.request.urlopen(f"{host}/api/tags", timeout=2) as r:
            data = json.loads(r.read().decode("utf-8"))
    except Exception:
        return []
    models = []
    for item in data.get("models", []):
        name = item.get("name") or item.get("model")
        if name:
            models.append(str(name))
    return sorted(models)


def call_qwen(prompt: str, model: str, host: str = "http://localhost:11434") -> str:
    """Call a local Ollama model (e.g. qwen2.5-coder). Returns text or raises."""
    import urllib.request

    payload = json.dumps({
        "model": model,
        "prompt": prompt,
        "stream": False,
        "options": {"temperature": 0.1},
    }).encode("utf-8")
    req = urllib.request.Request(
        f"{host}/api/generate", data=payload,
        headers={"Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=180) as r:
        data = json.loads(r.read().decode("utf-8"))
    return data.get("response", "").strip()


STATUS_LABELS = {
    "provided_file": "Provided file matched",
    "master_list_only": "Listed in master list only",
    "not_found": "Not found in uploaded folder or master list",
    "needs_review": "Needs review",
}


def _status_label(status: str) -> str:
    return STATUS_LABELS.get(status or "", "Needs review")


def _match_status(match) -> str:
    status = getattr(match, "status", "")
    if status:
        return status
    filename = str(getattr(match, "filename", "") or "")
    if "master list" in filename.lower() and "file not" in filename.lower():
        return "master_list_only"
    return "provided_file" if filename else "needs_review"


def _checklist_status(item) -> str:
    status = getattr(item, "status", "")
    if status:
        return status
    if getattr(item, "present", False):
        return "provided_file"
    return "not_found"


def _master_list_preview(result, limit: int = 12) -> tuple[int, list[str]]:
    master_path = Path(result.master_list) if getattr(result, "master_list", None) else None
    if not master_path or not master_path.exists():
        return 0, []
    try:
        entries = fa_mod.parse_master_list(master_path)
    except Exception:
        return 0, []
    return len(entries), [f"{e.doc_no} :: {e.name}" for e in entries[:limit]]


def build_qwen_prompt(result, run_dir: Path) -> str:
    """Assemble a precise, guard-railed diagnosis prompt for the coder model."""
    guardrails = ""
    if QWEN_INSTRUCTIONS_FILE.exists():
        guardrails = QWEN_INSTRUCTIONS_FILE.read_text(encoding="utf-8")

    sections = result.sections or {}
    low = [sid for sid, m in sorted(sections.items()) if not m]

    # Master list entries actually parsed (names + codes)
    ml_lines = []
    master_path = Path(result.master_list) if result.master_list else None
    if master_path and master_path.exists():
        try:
            for e in fa_mod.parse_master_list(master_path)[:400]:
                ml_lines.append(f"{e.doc_no} :: {e.name}")
        except Exception:
            pass

    # Files present in the folder
    file_lines = []
    docs_folder = Path(result.folder) if getattr(result, "folder", None) else None
    if docs_folder and docs_folder.exists():
        for f in sorted(docs_folder.rglob("*")):
            if f.is_file():
                file_lines.append(str(f.relative_to(docs_folder)))

    section_state = "\n".join(
        f"  {sid}: {len(sections.get(sid, []))} doc(s)"
        for sid in getattr(fa_mod, "SECTION_IDS", sorted(sections))
    )
    section_details = []
    for sid in getattr(fa_mod, "SECTION_IDS", sorted(sections)):
        matches = sections.get(sid, [])
        if not matches:
            section_details.append(f"  {sid}: NOT FOUND")
            continue
        section_details.append(f"  {sid}:")
        for m in matches:
            status = _status_label(_match_status(m))
            doc_no = f"{m.doc_no} - " if getattr(m, "doc_no", "") else ""
            section_details.append(f"    - {doc_no}{m.heading} [{m.filename}] ({status})")

    return f"""{guardrails}

────────────────────────────────────────────────────────
CURRENT RUN DIAGNOSIS TASK
────────────────────────────────────────────────────────
You are diagnosing ONE factory folder. Do NOT rewrite the whole script.
Propose the smallest GENERIC change (per the rules above) that would correctly
fill the empty/low sections IF the documents for them actually exist.

Empty sections this run: {', '.join(low) or 'none'}

Section match counts:
{section_state}

Section match details:
{chr(10).join(section_details)}

MASTER LIST ENTRIES PARSED ({len(ml_lines)}):
{chr(10).join(ml_lines) or '  (none parsed — the master list may be unreadable)'}

FILES IN FOLDER ({len(file_lines)}):
{chr(10).join(file_lines) or '  (none)'}

Respond in this exact structure:
1. DIAGNOSIS — for each empty section, say whether the document appears to exist
   (in uploaded files, master list only, or neither) or is genuinely NOT PROVIDED.
2. PROPOSED CHANGE — only for sections where the doc exists but wasn't matched.
   Give exact generic additions (e.g. keywords for SECTION_KEYWORDS[sid]) with a
   one-line justification each. No client-specific literals. Do not provide an
   auto-apply patch.
3. VALIDATION — the commands you would run to confirm no regression.
If nothing should change, say "NO CHANGE NEEDED" and why.
"""


def render_qwen_panel(result, run_dir: Path) -> None:
    settings = load_settings()
    model = settings.get("qwen_model", "qwen2.5-coder:7b")
    host = settings.get("ollama_host", "http://localhost:11434")

    with st.expander("🤖 Analyze with Qwen coder (AI diagnosis — propose only)"):
        st.caption(
            "Sends this run's section results, parsed master list, and file list to "
            f"a local Qwen coder model (`{model}`). It **proposes** generic fixes for "
            "empty sections — it never edits the script automatically. A human applies "
            "any change, then use **Re-run** to verify."
        )
        if not _ollama_available(host):
            st.warning(
                f"Ollama not reachable at `{host}`. Start it and pull the model:\n\n"
                f"```\nollama pull {model}\n```"
            )
            return
        installed_models = _ollama_models(host)
        if model not in installed_models:
            st.warning(
                f"`{model}` is not installed in Ollama, so the diagnosis cannot run yet.\n\n"
                f"Installed models: {', '.join(f'`{m}`' for m in installed_models) or 'none'}\n\n"
                f"Install Qwen with:\n\n```\nollama pull {model}\n```\n\n"
                "Or change the model name in **Settings** to one of the installed models."
            )
            return
        if st.button("Run AI diagnosis"):
            with st.spinner(f"Asking {model}…"):
                try:
                    prompt = build_qwen_prompt(result, run_dir)
                    answer = call_qwen(prompt, model, host)
                    st.session_state["qwen_answer"] = answer
                except Exception as e:
                    st.error(f"Qwen call failed: {e}")
        if st.session_state.get("qwen_answer"):
            st.markdown(st.session_state["qwen_answer"])
            st.info(
                "Review the proposal against the guardrails in "
                "`FA/QWEN_CODER_INSTRUCTIONS.md`. Apply the change to "
                "`FA/fa_automation.py`, then click **Re-run** (with *reload script* on)."
            )


# ── page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="FA Report Generator — TÜV Rheinland",
    page_icon="🏭",
    layout="wide",
)

# ── sidebar nav ───────────────────────────────────────────────────────────────
st.sidebar.image(
    "https://upload.wikimedia.org/wikipedia/commons/thumb/4/4e/T%C3%9CV_Rheinland_Logo.svg/320px-T%C3%9CV_Rheinland_Logo.svg.png",
    width=200,
)
st.sidebar.title("FA Portal")
st.sidebar.caption("Factory Audit Report Generator")

page = st.sidebar.radio(
    "Navigate",
    ["Generate Report", "Run History", "Settings"],
    label_visibility="collapsed",
)

if not FA_AVAILABLE:
    st.error(f"⚠️  fa_automation could not be loaded: `{_fa_import_error}`")
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: Generate Report
# ══════════════════════════════════════════════════════════════════════════════
if page == "Generate Report":
    st.title("🏭 Factory Audit Report Generator")
    st.markdown(
        "Upload the factory's documents (**ZIP file(s), individual files, or both**). "
        "The system will auto-fill the TÜV Rheinland FA template and "
        "generate a ready-to-download DOCX report."
    )

    settings = load_settings()

    # ── Step 1: Upload ─────────────────────────────────────────────────────────
    with st.expander("Step 1 — Upload Documents", expanded=True):
        col_a, col_b = st.columns(2)
        with col_a:
            docs_uploads = st.file_uploader(
                "Factory Documents / ZIP file(s)",
                type=["zip", "pdf", "docx", "doc", "xlsx", "xls", "jpg", "jpeg", "png", "webp"],
                accept_multiple_files=True,
                help=(
                    "Upload as many FA documents as needed. You can upload one ZIP, "
                    "multiple ZIPs, individual PDFs/Excels/images, or a mix."
                ),
            )
        with col_b:
            audit_plan_pdf = st.file_uploader(
                "Audit Plan / Application Form (PDF) — optional",
                type=["pdf"],
                help="Pre-audit application form or factory audit plan containing auditor names, date, country.",
            )

    # ── Step 2: Optional metadata overrides ───────────────────────────────────
    with st.expander("Optional — Manual Overrides", expanded=False):
        st.caption(
            "Usually leave these blank. Use only if the audit plan/application "
            "form is missing a value or the parser needs a manual correction."
        )
        c1, c2, c3 = st.columns(3)
        with c1:
            auditor_input = st.text_input(
                "Auditor code(s)",
                value=settings.get("default_auditor", ""),
                help="Use PB, PD, NP, or MH. For two auditors, enter e.g. PB, MH.",
                placeholder="Optional, e.g. PB or PB, MH",
            )
        with c2:
            audit_date_input = st.text_input(
                "Audit date",
                value="",
                placeholder="Optional, DD/MM/YYYY",
                help="Will be auto-detected if present in the audit plan/application form.",
            )
        with c3:
            country_input = st.text_input(
                "Country code",
                value=settings.get("default_country", ""),
                help="2-letter ISO code for FA report number, e.g. GB, AE, KSA.",
                placeholder="Optional",
            )
        factory_name_override = st.text_input(
            "Factory name override (optional)",
            value="",
            placeholder="Leave blank to auto-detect from documents",
        )

    # ── Step 3: Run ────────────────────────────────────────────────────────────
    run_col, _ = st.columns([1, 3])
    with run_col:
        run_btn = st.button("Generate FA Report", type="primary", disabled=not docs_uploads)

    if run_btn and docs_uploads:
        ts = int(time.time())
        run_dir = RUNS_DIR / f"run_{ts}"
        run_dir.mkdir(parents=True, exist_ok=True)

        with st.spinner("Extracting documents…"):
            try:
                docs_folder = extract_uploaded_documents(docs_uploads, run_dir / "docs")
            except Exception as e:
                st.error(f"Could not prepare uploaded documents: {e}")
                st.stop()

        # Save audit plan/application form both for audit trail and inside the
        # scanned docs folder so factory info, dates, names, and signatures can
        # be parsed from it.
        if audit_plan_pdf:
            audit_bytes = audit_plan_pdf.getvalue()
            (run_dir / audit_plan_pdf.name).write_bytes(audit_bytes)
            audit_dest = _safe_dest(docs_folder, audit_plan_pdf.name)
            audit_dest.parent.mkdir(parents=True, exist_ok=True)
            audit_dest.write_bytes(audit_bytes)

        template_path = Path(settings["template_path"])
        if not template_path.exists():
            st.error(
                f"FA template not found at `{template_path}`.  "
                "Please update the path in **Settings**."
            )
            st.stop()

        auditor_names: list[str] | None = None
        if auditor_input.strip():
            auditor_names = [n.strip() for n in auditor_input.split(",") if n.strip()]

        progress = st.progress(0, text="Matching documents…")

        try:
            result, out_docx = generate_fa_report(
                docs_folder=docs_folder,
                run_dir=run_dir,
                ts=ts,
                template_path=template_path,
                auditor_names=auditor_names,
                audit_date_input=audit_date_input,
                country_input=country_input,
                factory_name_override=factory_name_override,
                progress=progress,
            )

            # Remember inputs so the "Re-run" button can reproduce this run
            st.session_state["last_inputs"] = {
                "run_dir": str(run_dir),
                "docs_folder": str(docs_folder),
                "template_path": str(template_path),
                "auditor_names": auditor_names,
                "audit_date_input": audit_date_input,
                "country_input": country_input,
                "factory_name_override": factory_name_override,
            }

            # Warn about fields that could not be auto-detected and were not provided
            missing_fields = []
            if not result.audit_meta.get("auditor_names"):
                missing_fields.append("**Auditor name** — not found in documents and not entered")
            if not result.audit_meta.get("audit_date"):
                missing_fields.append("**Audit date** — not found in documents and not entered")
            if not result.audit_meta.get("country"):
                missing_fields.append("**Country code** — not entered")
            if not result.audit_meta.get("report_number"):
                missing_fields.append("**Report number** — cannot be generated (requires auditor name, date, and country)")
            if missing_fields:
                st.warning(
                    "The following fields are **blank in the report** "
                    "(not assumed — must be filled manually):\n\n"
                    + "\n".join(f"- {f}" for f in missing_fields)
                )

            # ── store run in session so we can show results below ──────────────
            st.session_state["last_run"] = {
                "ts": ts,
                "result": result,
                "out_docx": str(out_docx),
                "run_dir": str(run_dir),
                "factory_name": result.factory_info.get("name", docs_folder.name),
            }
            append_log(
                {
                    "ts": ts,
                    "timestamp": datetime.fromtimestamp(ts).isoformat(),
                    "factory_name": result.factory_info.get("name", docs_folder.name),
                    "report_number": result.audit_meta.get("report_number", ""),
                    "audit_date": result.audit_meta.get("audit_date", ""),
                    "auditor": result.audit_meta.get("auditor_names", []),
                    "out_docx": str(out_docx),
                    "run_dir": str(run_dir),
                    "status": "ok",
                    "feedback": None,
                    "feedback_notes": "",
                }
            )

        except Exception:
            tb = traceback.format_exc()
            st.error("Report generation failed. See details below.")
            st.code(tb, language="python")
            append_log(
                {
                    "ts": ts,
                    "timestamp": datetime.fromtimestamp(ts).isoformat(),
                    "factory_name": factory_name_override or (docs_uploads[0].name if docs_uploads else "uploaded FA documents"),
                    "status": "error",
                    "error": tb,
                }
            )
            st.stop()

    # ── Results ────────────────────────────────────────────────────────────────
    if "last_run" in st.session_state:
        lr = st.session_state["last_run"]
        result = lr["result"]
        out_docx = Path(lr["out_docx"])

        st.success(f"Report generated — **{out_docx.name}**")

        # ── Action row: download + re-run ───────────────────────────────────────
        act1, act2, act3 = st.columns([2, 1, 1])
        with act1:
            if out_docx.exists():
                st.download_button(
                    label="⬇️  Download FA Report (.docx)",
                    data=out_docx.read_bytes(),
                    file_name=out_docx.name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
        with act2:
            rerun_clicked = st.button(
                "🔁 Re-run",
                help="Re-run on the same uploaded documents. Use this after the "
                     "script has been improved, or to regenerate the report.",
            )
        with act3:
            reload_first = st.checkbox(
                "reload script",
                value=True,
                help="Reload fa_automation.py so any script fixes take effect before re-running.",
            )

        if rerun_clicked and "last_inputs" in st.session_state:
            li = st.session_state["last_inputs"]
            docs_folder2 = Path(li["docs_folder"])
            if not docs_folder2.exists():
                st.error("Original documents are no longer available. Please re-upload.")
            else:
                if reload_first:
                    reload_fa_module()
                ts2 = int(time.time())
                run_dir2 = RUNS_DIR / f"run_{ts2}"
                run_dir2.mkdir(parents=True, exist_ok=True)
                prog2 = st.progress(0, text="Re-running…")
                try:
                    result2, out_docx2 = generate_fa_report(
                        docs_folder=docs_folder2,
                        run_dir=run_dir2,
                        ts=ts2,
                        template_path=Path(li["template_path"]),
                        auditor_names=li["auditor_names"],
                        audit_date_input=li["audit_date_input"],
                        country_input=li["country_input"],
                        factory_name_override=li["factory_name_override"],
                        progress=prog2,
                    )
                    st.session_state["last_run"] = {
                        "ts": ts2,
                        "result": result2,
                        "out_docx": str(out_docx2),
                        "run_dir": str(run_dir2),
                        "factory_name": result2.factory_info.get("name", docs_folder2.name),
                    }
                    st.session_state["last_inputs"]["run_dir"] = str(run_dir2)
                    append_log({
                        "ts": ts2,
                        "timestamp": datetime.fromtimestamp(ts2).isoformat(),
                        "factory_name": result2.factory_info.get("name", docs_folder2.name),
                        "report_number": result2.audit_meta.get("report_number", ""),
                        "out_docx": str(out_docx2),
                        "run_dir": str(run_dir2),
                        "status": "ok",
                        "rerun": True,
                    })
                    st.success("Re-run complete — results refreshed below.")
                    st.rerun()
                except Exception:
                    st.error("Re-run failed.")
                    st.code(traceback.format_exc(), language="python")

        # ── AI diagnosis (Qwen coder) — propose-only, human decides ─────────────
        render_qwen_panel(result, Path(lr["run_dir"]))

        # ── Preview tabs ───────────────────────────────────────────────────────
        tab_meta, tab_checklist, tab_sections, tab_images = st.tabs(
            ["Audit Details", "Document Checklist", "Section Matches", "Images Matched"]
        )

        with tab_meta:
            meta = result.audit_meta or {}
            fi = result.factory_info or {}
            cols = st.columns(2)
            with cols[0]:
                st.markdown("**Factory**")
                st.write(fi.get("name", "—"))
                st.markdown("**Country**")
                st.write(meta.get("country", "—"))
                st.markdown("**Report Number**")
                st.write(meta.get("report_number", "—"))
            with cols[1]:
                st.markdown("**Audit Date**")
                st.write(meta.get("audit_date", "—"))
                st.markdown("**Auditor(s)**")
                st.write(", ".join(meta.get("auditor_names", [])) or "—")
                st.markdown("**Regulation**")
                st.write(meta.get("regulation", "—"))

        with tab_checklist:
            if result.checklist:
                provided = [c for c in result.checklist if _checklist_status(c) == "provided_file"]
                master_only = [c for c in result.checklist if _checklist_status(c) == "master_list_only"]
                missing = [c for c in result.checklist if _checklist_status(c) == "not_found"]
                needs_review = [c for c in result.checklist if _checklist_status(c) == "needs_review"]
                action_items = missing + master_only + needs_review
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Template Items Checked", len(result.checklist))
                c2.metric("Matched", len(provided))
                c3.metric("Action Needed", len(action_items))
                c4.metric("Not Found", len(missing))

                st.caption(
                    "This action checklist covers the full FA template, including "
                    "documents and required factory/product images. Matched items are "
                    "hidden so users only see what still needs attention."
                )

                if missing:
                    st.markdown("#### Missing")
                    for c in missing:
                        st.markdown(f"- **{c.section}** — {c.label}")
                if master_only:
                    st.markdown("#### Listed in master list only")
                    st.caption("These are referenced in the master list, but the actual source file was not uploaded.")
                    for c in master_only:
                        st.markdown(f"- **{c.section}** — {c.label} -> `{c.matched_file}`")
                if needs_review:
                    st.markdown("#### Needs review")
                    for c in needs_review:
                        st.markdown(f"- **{c.section}** — {c.label} -> `{c.matched_file or 'uncertain'}`")
                if not action_items:
                    st.success("No missing or uncertain checklist items found.")
            else:
                st.info("No checklist data.")

        with tab_sections:
            ml_count, ml_preview = _master_list_preview(result)
            if result.master_list:
                st.markdown(f"**Master list parsed:** {ml_count} entr{'y' if ml_count == 1 else 'ies'}")
                if ml_preview:
                    with st.expander("Master list preview"):
                        for line in ml_preview:
                            st.markdown(f"- `{line}`")
                elif ml_count == 0:
                    st.warning(
                        "A master list file was found, but no document-register rows were parsed. "
                        "If it is a scanned image, OCR or manual review may be needed."
                    )
            else:
                st.info("No master list file detected.")

            if result.sections:
                for sid, matches in sorted(result.sections.items()):
                    if matches:
                        with st.expander(f"Section {sid} — {len(matches)} match(es)"):
                            for m in matches:
                                status = _status_label(_match_status(m))
                                doc_no = f"`{m.doc_no}` " if getattr(m, "doc_no", "") else ""
                                st.markdown(
                                    f"- **{m.heading}**  \n"
                                    f"  {doc_no}`{m.filename}`  \n"
                                    f"  Status: **{status}**"
                                )
                    else:
                        st.markdown(
                            f"**Section {sid}:** {_status_label('not_found')}"
                        )
            else:
                st.info("No section matches.")

        with tab_images:
            if result.images:
                for slot_id, matches in result.images.items():
                    if matches:
                        label = matches[0].slot_label if matches else slot_id
                        with st.expander(f"{label} — {len(matches)} file(s)"):
                            for m in matches:
                                st.markdown(f"- `{m.filename}`")
            else:
                st.info("No images matched.")

        # ── Feedback ──────────────────────────────────────────────────────────
        st.markdown("---")
        st.subheader("📝 Report Feedback")
        st.caption(
            "Your feedback helps improve the automation. "
            "Flag anything that looks wrong so we can fix it."
        )
        with st.form("feedback_form"):
            rating = st.select_slider(
                "Overall quality",
                options=["Poor", "Fair", "Good", "Great"],
                value="Good",
            )
            issues = st.multiselect(
                "What went wrong? (select all that apply)",
                options=[
                    "Wrong factory name",
                    "Wrong auditor name",
                    "Wrong audit date",
                    "Wrong report number",
                    "Missing photos",
                    "Photos in wrong section",
                    "Wrong ISO cert placed",
                    "Wrong document matched to section",
                    "Section answers incorrect",
                    "Signature not placed",
                    "Other",
                ],
            )
            notes = st.text_area(
                "Additional notes",
                placeholder="Describe the issue in detail…",
                height=100,
            )
            submit_fb = st.form_submit_button("Submit Feedback")
            if submit_fb:
                fb_entry = {
                    "ts": lr["ts"],
                    "timestamp": datetime.fromtimestamp(lr["ts"]).isoformat(),
                    "factory_name": lr["factory_name"],
                    "report_number": result.audit_meta.get("report_number", ""),
                    "status": "feedback",
                    "feedback_rating": rating,
                    "feedback_issues": issues,
                    "feedback_notes": notes,
                }
                append_log(fb_entry)
                st.success("Feedback saved — thank you!")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: Run History
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Run History":
    st.title("📋 Run History")
    runs = load_runs()
    if not runs:
        st.info("No runs yet. Generate your first FA report from the main page.")
        st.stop()

    # Summary metrics
    ok_runs = [r for r in runs if r.get("status") == "ok"]
    err_runs = [r for r in runs if r.get("status") == "error"]
    fb_runs = [r for r in runs if r.get("status") == "feedback"]
    issues_all: list[str] = []
    for r in fb_runs:
        issues_all.extend(r.get("feedback_issues", []))

    col1, col2, col3 = st.columns(3)
    col1.metric("Total runs", len(ok_runs) + len(err_runs))
    col2.metric("Errors", len(err_runs))
    col3.metric("Feedback entries", len(fb_runs))

    if issues_all:
        from collections import Counter
        counts = Counter(issues_all).most_common()
        st.subheader("Most reported issues")
        for issue, n in counts:
            st.progress(n / counts[0][1], text=f"{issue} ({n}×)")

    st.markdown("---")
    st.subheader("Recent runs")

    seen_ts: set[int] = set()
    for entry in runs:
        if entry.get("status") == "feedback":
            continue
        ts = entry.get("ts", 0)
        if ts in seen_ts:
            continue
        seen_ts.add(ts)
        label = (
            f"[{entry.get('timestamp', '?')[:19]}] "
            f"{entry.get('factory_name', '—')}  |  "
            f"{entry.get('report_number', '—')}"
        )
        icon = "✅" if entry.get("status") == "ok" else "❌"
        with st.expander(f"{icon} {label}"):
            st.json(entry)
            out = entry.get("out_docx", "")
            if out and Path(out).exists():
                st.download_button(
                    "⬇️ Download report",
                    data=Path(out).read_bytes(),
                    file_name=Path(out).name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"dl_{ts}",
                )

    # Learning summary section
    if fb_runs:
        st.markdown("---")
        st.subheader("🧠 Learning Log — Common Issues")
        st.caption(
            "Patterns extracted from feedback. Use these to prioritise improvements to the automation."
        )
        from collections import Counter, defaultdict
        factory_issues: dict[str, list[str]] = defaultdict(list)
        for r in fb_runs:
            for issue in r.get("feedback_issues", []):
                factory_issues[r.get("factory_name", "Unknown")].append(issue)

        for factory, issues in sorted(factory_issues.items()):
            cnt = Counter(issues)
            issues_str = ", ".join(f"{k} ({v}×)" for k, v in cnt.most_common(3))
            st.markdown(f"- **{factory}**: {issues_str}")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: Settings
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Settings":
    st.title("⚙️ Settings")
    settings = load_settings()

    with st.form("settings_form"):
        st.subheader("Paths")
        template = st.text_input(
            "FA Template path (.docx)",
            value=settings["template_path"],
            help="Full path to the FA template Word document.",
        )
        output_dir = st.text_input(
            "Output / runs directory",
            value=settings["output_dir"],
        )

        st.subheader("Defaults")
        country = st.text_input(
            "Default country code",
            value=settings["default_country"],
            help="Used in FA report number generation. E.g. AE, KSA, VN, TH.",
        )
        auditor = st.text_input(
            "Default auditor name",
            value=settings.get("default_auditor", ""),
            placeholder="Pre-fill the auditor name field on every new run",
        )

        st.subheader("AI diagnosis (Qwen coder)")
        qwen_model = st.text_input(
            "Qwen model (Ollama)",
            value=settings.get("qwen_model", "qwen2.5-coder:7b"),
            help="Local Ollama model used for propose-only diagnosis. Pull it with "
                 "`ollama pull qwen2.5-coder:7b` (or a larger coder variant).",
        )
        ollama_host = st.text_input(
            "Ollama host",
            value=settings.get("ollama_host", "http://localhost:11434"),
        )

        st.subheader("Status")
        template_ok = Path(template).exists()
        if template_ok:
            st.success(f"Template found ✅  `{template}`")
        else:
            st.warning(f"Template not found at `{template}`")

        if st.form_submit_button("Save Settings"):
            save_settings(
                {
                    "template_path": template,
                    "output_dir": output_dir,
                    "default_country": country,
                    "default_auditor": auditor,
                    "qwen_model": qwen_model,
                    "ollama_host": ollama_host,
                }
            )
            st.success("Settings saved.")

    st.markdown("---")
    st.subheader("About")
    st.markdown(
        """
        **FA Report Generator** — TÜV Rheinland  
        Automates filling of the Harmonized Factory Audit Report template from client documents.

        **Photo layout:** Factory photos are arranged in a 2-column table — two headings side by side.  
        **Report number format:** `FA + CountryCode + AuditorInitials + DDMMYYYY + 01`  
        **Feedback:** Logged to `fa_portal/fa_feedback.jsonl` for continuous improvement.
        """
    )
    fa_version = getattr(fa_mod, "__version__", "N/A") if FA_AVAILABLE else "unavailable"
    st.caption(f"fa_automation version: `{fa_version}`")
