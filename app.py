"""
Dossier — Employee Document Organizer
=======================================
A Streamlit app that takes a folder of unorganized employee documents
(Aadhar, PAN, Payslips, Employment Letters — mixed PDFs/images, random
filenames) and:

  01 EXTRACT   OCR + Groq LLM read each file, identify the employee name,
               document type, and ID number. Similar name spellings are
               clustered together. Result: a review-ready Excel sheet.

  02 ORGANIZE  You verify/correct the Excel by hand, upload it back, and
               the app builds one folder per employee containing only
               their documents — packaged as a downloadable ZIP.

Run with:  streamlit run app.py
"""

import os
import re
import io
import json
import base64
import shutil
import uuid
import zipfile
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st
from fuzzywuzzy import fuzz

try:
    import pypdfium2 as pdfium
except ImportError:
    pdfium = None

from groq import Groq


# ============================================================================
# CONFIG
# ============================================================================

WORKSPACE_ROOT = Path("./workspace")
DOC_TYPES = ["Aadhar", "PAN", "Bank Details", "Payslip", "Employment Letter"]
DOC_ICONS = {
    "Aadhar": "◆", "PAN": "▲", "Bank Details": "◈",
    "Payslip": "●", "Employment Letter": "■",
}

# Groq deprecates/rotates models over time (e.g. llama-3.3-70b-versatile was
# retired June 2026) — if extraction starts failing silently, check
# console.groq.com/docs/models for the current vision-capable model name
# and update this constant.
APP_BUILD = "2026-08-21-v4"  # bump this yourself if you want a quick sanity check that
                              # the running server picked up your latest file changes
GROQ_VISION_MODEL = "qwen/qwen3.6-27b"
NAME_MATCH_THRESHOLD = 80

st.set_page_config(
    page_title="Dossier — Employee Document Organizer",
    page_icon="🗂️",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ============================================================================
# THEME — ink-ledger / case-file aesthetic
# ============================================================================

def inject_theme():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Spectral:wght@400;500;600;700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&display=swap');

    :root {
        --bg: #10151C;
        --panel: #1A222C;
        --panel-2: #1F2933;
        --border: #2C3846;
        --text: #E9E4D8;
        --text-muted: #8B97A6;
        --text-faint: #56626F;
        --accent: #C1440E;
        --accent-soft: rgba(193, 68, 14, 0.14);
        --brass: #C79A56;
        --brass-soft: rgba(199, 154, 86, 0.14);
        --serif: 'Spectral', Georgia, serif;
        --sans: 'IBM Plex Sans', sans-serif;
        --mono: 'IBM Plex Mono', monospace;
    }

    .stApp {
        background:
            radial-gradient(1200px 600px at 15% -10%, rgba(193,68,14,0.06), transparent 60%),
            radial-gradient(900px 500px at 100% 0%, rgba(199,154,86,0.05), transparent 55%),
            var(--bg);
        color: var(--text);
        font-family: var(--sans);
    }

    /* Hide default streamlit chrome bits for a cleaner surface */
    #MainMenu, footer, header {visibility: hidden;}
    .block-container {padding-top: 2.2rem; max-width: 1180px;}

    /* ---------- Hero ---------- */
    .dossier-hero {
        border-bottom: 1px solid var(--border);
        padding-bottom: 22px;
        margin-bottom: 30px;
        display: flex;
        justify-content: space-between;
        align-items: flex-end;
        flex-wrap: wrap;
        gap: 14px;
    }
    .dossier-eyebrow {
        font-family: var(--mono);
        font-size: 12px;
        letter-spacing: 3px;
        color: var(--accent);
        text-transform: uppercase;
        margin-bottom: 6px;
    }
    .dossier-title {
        font-family: var(--serif);
        font-weight: 600;
        font-size: 42px;
        color: var(--text);
        letter-spacing: 0.5px;
        line-height: 1.1;
    }
    .dossier-sub {
        font-family: var(--sans);
        color: var(--text-muted);
        font-size: 14.5px;
        max-width: 480px;
        margin-top: 6px;
    }
    .dossier-stamp {
        font-family: var(--mono);
        font-size: 11px;
        letter-spacing: 2px;
        color: var(--text-faint);
        border: 1px dashed var(--border);
        padding: 8px 14px;
        text-align: center;
        transform: rotate(-2deg);
    }
    .dossier-stamp b { color: var(--brass); display:block; font-size: 13px; letter-spacing: 3px;}

    /* ---------- Step marker ---------- */
    .step-marker {
        display: flex;
        align-items: center;
        gap: 12px;
        margin: 6px 0 18px 0;
    }
    .step-num {
        font-family: var(--mono);
        font-size: 13px;
        color: var(--accent);
        border: 1px solid var(--accent-soft);
        background: var(--accent-soft);
        padding: 3px 10px;
        letter-spacing: 1px;
    }
    .step-label {
        font-family: var(--serif);
        font-size: 22px;
        font-weight: 600;
        color: var(--text);
    }
    .step-desc { color: var(--text-muted); font-size: 13.5px; margin: -8px 0 20px 0; }

    /* ---------- Panels ---------- */
    .panel {
        background: var(--panel);
        border: 1px solid var(--border);
        border-radius: 4px;
        padding: 22px 24px;
        margin-bottom: 18px;
    }

    /* ---------- Tabs as ledger sections ---------- */
    .stTabs [data-baseweb="tab-list"] { gap: 4px; border-bottom: 1px solid var(--border); }
    .stTabs [data-baseweb="tab"] {
        font-family: var(--mono);
        letter-spacing: 1.5px;
        font-size: 12.5px;
        color: var(--text-muted);
        background: transparent;
        padding: 10px 4px;
        margin-right: 22px;
    }
    .stTabs [aria-selected="true"] {
        color: var(--accent) !important;
        border-bottom: 2px solid var(--accent) !important;
    }

    /* ---------- Buttons ---------- */
    .stButton>button, .stDownloadButton>button {
        font-family: var(--mono);
        letter-spacing: 1px;
        font-size: 12.5px;
        text-transform: uppercase;
        background: var(--accent);
        color: #F6EFE6;
        border: none;
        border-radius: 3px;
        padding: 10px 22px;
    }
    .stButton>button:hover, .stDownloadButton>button:hover {
        background: #A5390B;
        color: #F6EFE6;
    }

    /* ---------- File uploader ---------- */
    [data-testid="stFileUploaderDropzone"] {
        background: var(--panel-2);
        border: 1px dashed var(--border);
        border-radius: 4px;
    }

    /* ---------- Folder tab cards ---------- */
    .folder-card {
        position: relative;
        background: var(--panel-2);
        border: 1px solid var(--border);
        border-radius: 2px;
        margin-top: 22px;
        padding: 18px 18px 14px 18px;
    }
    .folder-tab {
        position: absolute;
        top: -14px;
        left: 18px;
        background: var(--brass);
        color: #1A1305;
        font-family: var(--mono);
        font-size: 11px;
        font-weight: 600;
        letter-spacing: 1px;
        padding: 4px 12px;
        border-radius: 2px 2px 0 0;
    }
    .folder-name {
        font-family: var(--serif);
        font-size: 19px;
        font-weight: 600;
        color: var(--text);
        margin: 10px 0 10px 0;
    }
    .doc-row { display: flex; gap: 10px; flex-wrap: wrap; }
    .doc-chip {
        font-family: var(--mono);
        font-size: 11.5px;
        padding: 5px 10px;
        border-radius: 2px;
        letter-spacing: 0.5px;
    }
    .doc-present { background: var(--brass-soft); color: var(--brass); border: 1px solid var(--brass-soft); }
    .doc-missing { background: transparent; color: var(--text-faint); border: 1px solid var(--border); }

    /* ---------- Log lines ---------- */
    .log-line {
        font-family: var(--mono);
        font-size: 12px;
        color: var(--text-muted);
        padding: 3px 0;
        border-bottom: 1px solid rgba(255,255,255,0.03);
    }
    .log-line b { color: var(--text); }
    .log-tag {
        font-size: 10px;
        padding: 1px 6px;
        border-radius: 2px;
        margin-left: 6px;
        letter-spacing: 1px;
    }
    .tag-ok { background: var(--brass-soft); color: var(--brass); }
    .tag-warn { background: var(--accent-soft); color: var(--accent); }

    [data-testid="stMetricValue"] { font-family: var(--serif); color: var(--text); }
    [data-testid="stMetricLabel"] { font-family: var(--mono); letter-spacing: 1px; font-size: 11px; }
    </style>
    """, unsafe_allow_html=True)


def hero():
    st.markdown(f"""
    <div class="dossier-hero">
        <div>
            <div class="dossier-eyebrow">Employee Records · Auto-Filing</div>
            <div class="dossier-title">Dossier</div>
            <div class="dossier-sub">Feed it a folder of loose Aadhar cards, PAN cards, payslips and
            offer letters — it reads, identifies, and files each one under the right employee.</div>
        </div>
        <div class="dossier-stamp">BUILD {APP_BUILD}<br><b>VERIFIED FILING</b></div>
    </div>
    """, unsafe_allow_html=True)


def step_marker(num, label, desc):
    st.markdown(f"""
    <div class="step-marker">
        <span class="step-num">STEP {num}</span>
        <span class="step-label">{label}</span>
    </div>
    <div class="step-desc">{desc}</div>
    """, unsafe_allow_html=True)


# ============================================================================
# CORE LOGIC
# ============================================================================

def get_session_workspace() -> Path:
    if "session_id" not in st.session_state:
        st.session_state.session_id = uuid.uuid4().hex[:10]
    ws = WORKSPACE_ROOT / st.session_state.session_id
    (ws / "uploads").mkdir(parents=True, exist_ok=True)
    (ws / "organized").mkdir(parents=True, exist_ok=True)
    return ws


def get_groq_client():
    api_key = st.session_state.get("groq_api_key", "")
    if not api_key:
        return None
    return Groq(api_key=api_key)


def file_to_image_path(filepath: str):
    """Returns (image_path, error) for the given file. PDFs get their first
    page rendered to a temp JPEG; images pass through unchanged."""
    ext = Path(filepath).suffix.lower()
    if ext in (".jpg", ".jpeg", ".png"):
        return filepath, None
    if ext == ".pdf":
        if not pdfium:
            return None, "pypdfium2 not installed (pip install pypdfium2)"
        try:
            pdf = pdfium.PdfDocument(filepath)
            if len(pdf) == 0:
                pdf.close()
                return None, "PDF has no pages"
            page = pdf[0]
            bitmap = page.render(scale=2)
            image = bitmap.to_pil()
            tmp_path = filepath + ".page1.jpg"
            image.save(tmp_path, "JPEG", quality=90)
            image.close()
            bitmap.close()
            page.close()
            pdf.close()
            return tmp_path, None
        except Exception as e:
            return None, f"PDF render failed: {e}"
    return None, f"unsupported file type: {ext}"


def image_to_data_url(image_path: str) -> str:
    ext = Path(image_path).suffix.lower().lstrip(".")
    mime = "jpeg" if ext == "jpg" else ext
    with open(image_path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    return f"data:image/{mime};base64,{b64}"


EXTRACTION_PROMPT = """You are looking at a photographed or scanned page that may contain one or
more Indian employee documents — possibly rotated, upside down, folded, or photocopied — pasted or
printed on a plain sheet. Document types to recognize: Aadhar card, PAN card, Bank passbook/account
details sheet, Payslip / salary slip, Employment / offer / appointment letter.

Read any printed or handwritten name, even if the image is sideways or upside down. If more than one
document appears on the page, identify the primary/most complete one.

Respond with ONLY a JSON object, no markdown fences, no extra text:

{
  "employee_name": "<full name of the person this document belongs to, or null if genuinely unreadable>",
  "document_type": "<one of: Aadhar, PAN, Bank Details, Payslip, Employment Letter, Unknown>",
  "id_number": "<Aadhar number, PAN number, or account number if visible, else null>"
}

Filename (may be uninformative, e.g. a generic WhatsApp export name — do not rely on it): {filename}
"""


def extract_json_block(raw: str) -> dict:
    """Pulls the first {...} JSON object out of a string, tolerating stray
    reasoning text, code fences, or commentary around it."""
    raw = re.sub(r"<think>.*?</think>", "", raw, flags=re.DOTALL)
    raw = re.sub(r"```json|```", "", raw).strip()
    match = re.search(r"\{.*\}", raw, flags=re.DOTALL)
    if not match:
        raise ValueError(f"no JSON object found in model output: {raw[:200]!r}")
    return json.loads(match.group(0))


def extract_info_with_vision(client, filepath: str, filename: str) -> dict:
    image_path, err = file_to_image_path(filepath)
    if not image_path:
        return {"employee_name": None, "document_type": "Unknown", "id_number": None, "_error": err}

    raw = None
    try:
        data_url = image_to_data_url(image_path)
        response = client.chat.completions.create(
            model=GROQ_VISION_MODEL,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "text", "text": EXTRACTION_PROMPT.replace("{filename}", filename)},
                    {"type": "image_url", "image_url": {"url": data_url}},
                ],
            }],
            temperature=0.3,
            reasoning_effort="none",       # qwen3.6 defaults to "thinking" mode which
                                            # mixes reasoning text into the output; we
                                            # only want the final JSON.
            response_format={"type": "json_object"},
        )
        raw = response.choices[0].message.content
        return extract_json_block(raw)
    except Exception as e:
        detail = f"{type(e).__name__}: {e}"
        if raw:
            detail += f" | raw output: {raw[:150]!r}"
        return {"employee_name": None, "document_type": "Unknown", "id_number": None, "_error": detail}


def cluster_employee_names(records: list, threshold: int) -> list:
    canonical_names = []
    for rec in records:
        name = rec.get("employee_name")
        if not name:
            rec["Employee Name"] = "UNMATCHED"
            continue
        best_match, best_score = None, 0
        for canon in canonical_names:
            score = fuzz.token_set_ratio(name.lower(), canon.lower())
            if score > best_score:
                best_match, best_score = canon, score
        if best_score >= threshold:
            rec["Employee Name"] = best_match
        else:
            canonical_names.append(name)
            rec["Employee Name"] = name
    return records


def build_pivot(records: list) -> pd.DataFrame:
    df = pd.DataFrame(records)
    pivot = df.pivot_table(
        index="Employee Name",
        columns="document_type",
        values="file_name",
        aggfunc=lambda x: " | ".join(x),
    ).reset_index()
    for dt in DOC_TYPES:
        if dt not in pivot.columns:
            pivot[dt] = None
    ordered_cols = ["Employee Name"] + DOC_TYPES + [c for c in pivot.columns if c not in DOC_TYPES + ["Employee Name", "Unknown"]]
    if "Unknown" in pivot.columns:
        ordered_cols.append("Unknown")
    pivot = pivot[[c for c in ordered_cols if c in pivot.columns]]
    return pivot


def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Employee Documents")
    return buf.getvalue()


def make_zip_of_folder(folder: Path) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(folder):
            for f in files:
                full = Path(root) / f
                zf.write(full, arcname=full.relative_to(folder))
    return buf.getvalue()


# ============================================================================
# UI — STEP 1: EXTRACT
# ============================================================================

def render_step1():
    step_marker("01", "Extract & Identify", "Upload every document from the unorganized folder (select all files at once). Each one is OCR'd and read by the LLM to identify the employee and document type.")

    with st.container():
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        col1, col2 = st.columns([3, 1])
        with col1:
            uploaded_files = st.file_uploader(
                "Select all files from your unorganized documents folder",
                accept_multiple_files=True,
                type=["pdf", "jpg", "jpeg", "png"],
                key="step1_uploader",
            )
        with col2:
            st.session_state.groq_api_key = st.text_input(
                "Groq API key", type="password",
                value=st.session_state.get("groq_api_key", os.environ.get("GROQ_API_KEY", "")),
                help="From console.groq.com/keys",
            )
            threshold = st.slider("Name match strictness", 60, 95, NAME_MATCH_THRESHOLD, 5,
                                   help="Higher = stricter matching when grouping name variants across documents")
        st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_files:
        st.caption(f"{len(uploaded_files)} file(s) ready.")
        run = st.button("▸ Run extraction", disabled=not st.session_state.get("groq_api_key"))
        if not st.session_state.get("groq_api_key"):
            st.info("Enter your Groq API key to run extraction.")

        if run:
            client = get_groq_client()
            ws = get_session_workspace()
            records = []

            progress = st.progress(0.0)
            log_box = st.container()
            log_lines = []

            for i, uf in enumerate(uploaded_files):
                dest = ws / "uploads" / uf.name
                with open(dest, "wb") as f:
                    f.write(uf.getbuffer())

                info = extract_info_with_vision(client, str(dest), uf.name)

                records.append({
                    "employee_name": info.get("employee_name"),
                    "document_type": info.get("document_type", "Unknown"),
                    "id_number": info.get("id_number"),
                    "file_path": str(dest),
                    "file_name": uf.name,
                })

                if info.get("_error"):
                    tag, tag_text = "tag-warn", "ERROR"
                    detail = info["_error"][:160]
                elif info.get("employee_name"):
                    tag, tag_text = "tag-ok", "OK"
                    detail = f'{info.get("employee_name")} · {info.get("document_type", "Unknown")}'
                else:
                    tag, tag_text = "tag-warn", "UNCLEAR"
                    detail = "no name readable"

                log_lines.append(
                    f'<div class="log-line"><b>{uf.name}</b> → {detail} '
                    f'<span class="log-tag {tag}">{tag_text}</span></div>'
                )
                progress.progress((i + 1) / len(uploaded_files))

            with log_box:
                st.markdown("".join(log_lines), unsafe_allow_html=True)

            records = cluster_employee_names(records, threshold)
            pivot = build_pivot(records)
            st.session_state.pivot_df = pivot
            st.session_state.raw_records = records
            st.success(f"Processed {len(uploaded_files)} files → {len(pivot)} employee(s) identified.")

    if "pivot_df" in st.session_state:
        st.divider()
        m1, m2, m3 = st.columns(3)
        df = st.session_state.pivot_df
        m1.metric("Employees identified", len(df[df["Employee Name"] != "UNMATCHED"]))
        m2.metric("Unmatched documents", int((df["Employee Name"] == "UNMATCHED").sum()))
        m3.metric("Total documents", len(st.session_state.raw_records))

        st.markdown("**Review before download** — file names only need to stay exact in cells you don't touch; you can fix names, move a doc to the right column, or leave a cell blank.")
        edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="editor1")

        excel_bytes = df_to_excel_bytes(edited)
        st.download_button(
            "⬇ Download Excel for verification",
            data=excel_bytes,
            file_name=f"employee_documents_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption("Open this in Excel, fix anything wrong (especially UNMATCHED rows), save it, then move to Step 02 below.")


# ============================================================================
# UI — STEP 2: ORGANIZE
# ============================================================================

def render_step2():
    step_marker("02", "Verify & Organize", "Upload your corrected Excel sheet. The app matches each row's document filenames back to the originally uploaded files and builds one folder per employee.")

    ws = get_session_workspace()
    has_session_files = any((ws / "uploads").iterdir()) if (ws / "uploads").exists() else False

    st.markdown('<div class="panel">', unsafe_allow_html=True)
    verified_excel = st.file_uploader("Upload verified Excel sheet", type=["xlsx"], key="step2_excel")

    reupload_files = None
    if not has_session_files:
        st.warning("No documents found from this session — re-upload the original files so they can be copied into folders.")
        reupload_files = st.file_uploader("Re-upload original documents", accept_multiple_files=True, key="step2_files")
    st.markdown('</div>', unsafe_allow_html=True)

    if reupload_files:
        for uf in reupload_files:
            dest = ws / "uploads" / uf.name
            with open(dest, "wb") as f:
                f.write(uf.getbuffer())

    if verified_excel and st.button("▸ Build employee folders"):
        df = pd.read_excel(verified_excel)
        organized_root = ws / "organized"
        shutil.rmtree(organized_root, ignore_errors=True)
        organized_root.mkdir(parents=True, exist_ok=True)

        file_index = {p.name: p for p in (ws / "uploads").glob("*") if p.is_file()}
        cards = []

        for _, row in df.iterrows():
            employee = str(row.get("Employee Name", "")).strip()
            if not employee or employee.lower() in ("nan", "unmatched"):
                continue

            emp_folder = organized_root / employee
            emp_folder.mkdir(parents=True, exist_ok=True)
            present = {}

            for dt in DOC_TYPES:
                cell = row.get(dt)
                present[dt] = False
                if pd.isna(cell):
                    continue
                for fname in str(cell).split(" | "):
                    fname = fname.strip()
                    src = file_index.get(fname) or file_index.get(Path(fname).name)
                    if src and src.exists():
                        shutil.copy2(src, emp_folder / src.name)
                        present[dt] = True

            cards.append((employee, present))

        st.session_state.organized_cards = cards
        st.session_state.organized_root = str(organized_root)
        st.success(f"Built folders for {len(cards)} employee(s).")

    if st.session_state.get("organized_cards"):
        cols = st.columns(2)
        for i, (employee, present) in enumerate(st.session_state.organized_cards):
            chips = "".join(
                f'<span class="doc-chip {"doc-present" if present[dt] else "doc-missing"}">'
                f'{DOC_ICONS[dt]} {dt}</span>'
                for dt in DOC_TYPES
            )
            with cols[i % 2]:
                st.markdown(f"""
                <div class="folder-card">
                    <div class="folder-tab">FILE</div>
                    <div class="folder-name">{employee}</div>
                    <div class="doc-row">{chips}</div>
                </div>
                """, unsafe_allow_html=True)

        st.divider()
        zip_bytes = make_zip_of_folder(Path(st.session_state.organized_root))
        st.download_button(
            "⬇ Download organized folders (.zip)",
            data=zip_bytes,
            file_name=f"organized_employees_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
            mime="application/zip",
        )


# ============================================================================
# MAIN
# ============================================================================

def main():
    inject_theme()
    hero()

    tab1, tab2 = st.tabs(["01 · EXTRACT", "02 · ORGANIZE"])
    with tab1:
        render_step1()
    with tab2:
        render_step2()


if __name__ == "__main__":
    main()
