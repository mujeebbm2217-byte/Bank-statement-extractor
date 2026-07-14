"""
Bank Statement Keyword Amount Extractor - Streamlit Version
-----------------------------------------------------------------------------
Run locally with:   streamlit run app.py
Deploy free on:      https://share.streamlit.io  (Streamlit Community Cloud)
                      -- push this file + requirements.txt to a GitHub repo,
                      then "New app" on share.streamlit.io and point it at
                      app.py. No server setup needed.

Requires (see requirements.txt):
    streamlit pandas numpy pdfplumber openpyxl pillow reportlab

-----------------------------------------------------------------------------
WHAT CHANGED FROM THE DESKTOP (CustomTkinter) VERSION
-----------------------------------------------------------------------------
The actual data logic -- PDF table-strategy scoring, the wrapped-narration
merge heuristic, column auto-mapping, amount cleaning, PDF<->Excel
conversion -- is copied over UNCHANGED. Only the UI layer changed:

    tkinter/customtkinter windows      -->  Streamlit widgets & pages
    filedialog.askopenfilename         -->  st.file_uploader
    ttk.Treeview tables                -->  st.dataframe
    Toplevel wizard (Step 1 / Step 2)  -->  same 2-step flow, driven by
                                             st.session_state instead of
                                             separate windows
    In-memory .exe output files        -->  st.download_button (BytesIO)

This is a SCAFFOLD: the core workflow (upload -> pick header row -> confirm
columns -> search) all works, but the visual polish (navy/gold theme, page
thumbnails, etc.) is intentionally minimal so you can restyle it your way.

-----------------------------------------------------------------------------
IMPORTANT: LOGIN / PASSWORD NOTE
-----------------------------------------------------------------------------
Streamlit Community Cloud apps are PUBLIC by default (anyone with the link
can open them) unless you turn on Streamlit's built-in viewer auth in the
app settings. The username/password gate below is a lightweight extra layer,
same idea as the desktop app: we only store a SHA-256 fingerprint, never the
real password.

To set your own password:
    1. Run this once in any terminal (no need to run the app):
           python -c "import hashlib; print(hashlib.sha256(b'yourNewPassword').hexdigest())"
    2. Copy the printed hash into APP_PASSWORD_HASH below.
"""

import hashlib
import io
import re

import numpy as np
import pandas as pd
import pdfplumber
import openpyxl
from openpyxl.utils import get_column_letter
import streamlit as st

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
    from reportlab.lib.units import mm
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

NONE_OPTION = "-- None --"
DEVELOPER_NAME = "Mohammad Mujeeb"

# ----------------------------- Theme (navy/gold) -----------------------------
NAVY_BG = "#0f172a"
GOLD = "#eab308"
TEAL = "#14b8a6"

APP_USERNAME = "admin"
APP_PASSWORD_HASH = "707cfae1e410d17ac820cd4b290040561bb963b00fa1c13807d63723539fdba0"
# ^ hash for "changeme123" -- change this, see note above.


def _hash_password(plain_text_password: str) -> str:
    return hashlib.sha256(plain_text_password.encode("utf-8")).hexdigest()


# ----------------------------- Core parsing logic -----------------------------
# (Unchanged from the desktop version -- see original file for full comments.)

PDF_TABLE_STRATEGIES = [
    {"vertical_strategy": "lines", "horizontal_strategy": "lines"},
    {"vertical_strategy": "lines", "horizontal_strategy": "text",
     "text_x_tolerance": 1, "text_y_tolerance": 3,
     "snap_tolerance": 3, "join_tolerance": 3},
    {"vertical_strategy": "text", "horizontal_strategy": "text",
     "text_x_tolerance": 1, "text_y_tolerance": 3},
]

_DATE_CELL_RE = re.compile(r'^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$')


def map_columns(columns):
    """Map actual column names to standard names: reference, date, debit,
    credit, balance. Only a STARTING GUESS -- user can override in the UI."""
    mapping = {}
    col_lower = {c: str(c).lower().strip() for c in columns}

    date_kw = ['date', 'txn date', 'value date', 'transaction date']
    ref_kw = ['narration', 'description', 'particulars', 'reference', 'remarks',
              'transaction details', 'details', 'transaction remarks', 'recipient', 'payee']
    debit_kw = ['debit', 'withdrawal', 'dr amount', 'withdrawal amt', 'payment', 'payment amt']
    credit_kw = ['credit', 'deposit', 'cr amount', 'deposit amt', 'receipt', 'receipt amt']
    balance_kw = ['balance', 'closing balance', 'running balance']

    for col, low in col_lower.items():
        if 'reference' not in mapping and any(k in low for k in ref_kw):
            mapping['reference'] = col
        elif 'date' not in mapping and any(k in low for k in date_kw):
            mapping['date'] = col
        elif 'debit' not in mapping and any(k in low for k in debit_kw):
            mapping['debit'] = col
        elif 'credit' not in mapping and any(k in low for k in credit_kw):
            mapping['credit'] = col
        elif 'balance' not in mapping and any(k in low for k in balance_kw):
            mapping['balance'] = col
    return mapping


def clean_amount(val):
    """Convert amount string/number to float, handling commas, currency
    symbols, blanks."""
    if pd.isna(val):
        return np.nan
    s = str(val).strip()
    if s == '' or s.lower() in ['nan', '-', 'none']:
        return np.nan
    s = re.sub(r'[^\d.\-]', '', s)
    if s in ['', '-', '.']:
        return np.nan
    try:
        return float(s)
    except ValueError:
        return np.nan


def _is_blank_row(row):
    return all((c is None or str(c).strip() == '') for c in row)


def _score_strategy(tables):
    """Count rows containing a date-like cell -- a strong bank-statement
    -specific signal of a real transaction row."""
    score = 0
    for table in tables:
        for row in table:
            if not row:
                continue
            if any(c and _DATE_CELL_RE.match(str(c).strip()) for c in row):
                score += 1
    return score


def _extract_pdf_tables_best_strategy(pdf):
    """Try each strategy on a couple of sample pages and score how many real
    transaction rows each one recovers."""
    sample_pages = pdf.pages[:min(2, len(pdf.pages))]
    scores = []
    for settings in PDF_TABLE_STRATEGIES:
        total = 0
        for page in sample_pages:
            try:
                total += _score_strategy(page.extract_tables(settings))
            except Exception:
                pass
        scores.append(total)

    best_score = max(scores) if scores else 0
    threshold = best_score * 0.7
    for settings, score in zip(PDF_TABLE_STRATEGIES, scores):
        if score >= threshold:
            return settings
    return PDF_TABLE_STRATEGIES[0]


def extract_raw_rows(uploaded_file, ext):
    """
    Extract RAW rows (no header assumption yet) plus optional PDF page
    images for preview. `uploaded_file` is a Streamlit UploadedFile
    (file-like / BytesIO) -- pdfplumber and pandas both accept it directly.
    Returns (rows, page_images).
    """
    page_images = []

    if ext == 'pdf':
        rows = []
        uploaded_file.seek(0)
        with pdfplumber.open(uploaded_file) as pdf:
            settings = _extract_pdf_tables_best_strategy(pdf)

            for page in pdf.pages:
                try:
                    tables = page.extract_tables(settings)
                except Exception:
                    tables = []

                if tables:
                    main_table = max(tables, key=len)
                    for r in main_table:
                        if r is None or _is_blank_row(r):
                            continue
                        rows.append(list(r))

                try:
                    img = page.to_image(resolution=100).original
                    page_images.append(img)
                except Exception:
                    pass
        return rows, page_images

    uploaded_file.seek(0)
    if ext == 'csv':
        df_raw = pd.read_csv(uploaded_file, header=None, dtype=str)
    elif ext in ('xlsx', 'xls'):
        df_raw = pd.read_excel(uploaded_file, header=None, dtype=str)
    else:
        raise ValueError(f"Unsupported file type: {ext}")

    df_raw = df_raw.fillna('')
    rows = df_raw.values.tolist()
    return rows, page_images


def pad_row(row, width, merge_into=None):
    """Force a row to exactly `width` cells. Overflow cells (extra cells from
    wrapped narration text) are merged into `merge_into` instead of dropped."""
    row = ['' if c is None else c for c in row]

    if len(row) < width:
        row = row + [''] * (width - len(row))
    elif len(row) > width:
        target = merge_into if merge_into is not None else width - 1
        target = max(0, min(target, width - 1))

        overflow = row[width:]
        kept = row[:width]

        pieces = [str(kept[target])] + [str(c) for c in overflow if str(c).strip() != '']
        kept[target] = ' '.join(p for p in pieces if p.strip() != '')

        row = kept
    return row


def merge_wrapped_continuation_rows(rows, date_idx=None, debit_idx=None,
                                     credit_idx=None, narration_idx=None):
    """Fold "orphan" wrapped-narration rows into the transaction row above
    them. See the desktop-version docstring for the full explanation."""
    merged = []
    for row in rows:
        row = list(row)

        def cell(i):
            return row[i] if 0 <= i < len(row) and row[i] is not None else ''

        has_any_text = any(str(c).strip() != '' for c in row)

        is_continuation = False
        if merged and has_any_text:
            checks = []
            if date_idx is not None:
                checks.append(str(cell(date_idx)).strip() == '')
            if debit_idx is not None:
                checks.append(str(cell(debit_idx)).strip() == '')
            if credit_idx is not None:
                checks.append(str(cell(credit_idx)).strip() == '')
            if checks:
                is_continuation = all(checks)
            else:
                is_continuation = str(cell(0)).strip() == ''

        if is_continuation:
            extra_text = ' '.join(str(c).strip() for c in row if str(c).strip() != '')
            prev = merged[-1]
            target = narration_idx if narration_idx is not None else (len(prev) - 1)
            target = max(0, min(target, len(prev) - 1)) if prev else 0
            if prev:
                existing = str(prev[target]) if prev[target] is not None else ''
                prev[target] = (existing + ' ' + extra_text).strip() if existing else extra_text
            else:
                merged.append(row)
        else:
            merged.append(row)

    return merged


def build_clean_df(df_raw, col_map):
    """col_map keys used: 'reference' (required), 'debit' (optional),
    'credit' (optional)."""
    df = pd.DataFrame()
    df['Reference'] = df_raw[col_map['reference']].astype(str)

    if 'debit' in col_map:
        df['Debit'] = df_raw[col_map['debit']].apply(clean_amount)
    if 'credit' in col_map:
        df['Credit'] = df_raw[col_map['credit']].apply(clean_amount)

    df = df[df['Reference'].str.strip().str.lower().apply(lambda x: x not in ['', 'nan', 'none'])]
    df = df.reset_index(drop=True)
    return df


def guess_header_row(rows, max_scan=40):
    """Best-effort guess at which row is the real header row, so the UI can
    default the row-picker to something sensible instead of always 0."""
    keywords = ['date', 'narration', 'description', 'particulars', 'reference',
                'debit', 'credit', 'withdrawal', 'deposit', 'balance']
    for idx, row in enumerate(rows[:max_scan]):
        text = ' '.join(str(c).lower() for c in row if c is not None)
        if sum(k in text for k in keywords) >= 2:
            return idx
    return 0


# ----------------------------- PDF <-> Excel conversion -----------------------------

def convert_pdf_to_excel_bytes(uploaded_file):
    """Raw dump of every extracted PDF row into an in-memory Excel file.
    Returns (BytesIO, row_count)."""
    rows, _ = extract_raw_rows(uploaded_file, 'pdf')
    if not rows:
        raise ValueError("No data could be extracted from this PDF.")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Extracted Data"

    max_cols = 0
    for row in rows:
        ws.append(row)
        max_cols = max(max_cols, len(row))

    col_widths = [0] * max_cols
    for row in rows:
        for i, cell in enumerate(row):
            if i < max_cols:
                col_widths[i] = max(col_widths[i], len(str(cell)) if cell is not None else 0)
    for i, width in enumerate(col_widths):
        ws.column_dimensions[get_column_letter(i + 1)].width = min(max(width + 2, 10), 60)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, len(rows)


def convert_table_to_pdf_bytes(uploaded_file, ext):
    """Reads an Excel/CSV file and renders it as a formatted in-memory PDF.
    Returns (BytesIO, row_count)."""
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("reportlab is not installed. Add it to requirements.txt.")

    uploaded_file.seek(0)
    if ext == 'csv':
        df = pd.read_csv(uploaded_file, dtype=str)
    else:
        df = pd.read_excel(uploaded_file, dtype=str)
    df = df.fillna('')

    data = [list(df.columns)] + df.astype(str).values.tolist()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4),
        leftMargin=15 * mm, rightMargin=15 * mm, topMargin=15 * mm, bottomMargin=15 * mm
    )
    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(NAVY_BG)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor(GOLD)),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f0fdfa')]),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    doc.build([table])
    buf.seek(0)
    return buf, len(df)


# ----------------------------- Streamlit UI -----------------------------

st.set_page_config(page_title="Bank Statement Keyword Extractor", layout="wide")

st.markdown(
    f"""
    <style>
    .stApp {{ background-color: {NAVY_BG}; }}
    h1, h2, h3 {{ color: {GOLD}; }}
    </style>
    """,
    unsafe_allow_html=True,
)


def login_gate():
    st.title("🔐 Bank Statement Extractor -- Login")
    with st.form("login_form"):
        user = st.text_input("Username")
        pwd = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Login")
    if submitted:
        if user == APP_USERNAME and _hash_password(pwd) == APP_PASSWORD_HASH:
            st.session_state["logged_in"] = True
            st.rerun()
        else:
            st.error("Invalid username or password")
    st.caption(f"Developed by {DEVELOPER_NAME}")


def reset_workflow_state():
    for key in ("raw_rows", "page_images", "ext", "step", "df_raw", "df", "header_idx"):
        st.session_state.pop(key, None)


def keyword_search_tab():
    st.subheader("Upload a bank statement")
    uploaded = st.file_uploader(
        "PDF / Excel / CSV", type=["pdf", "xlsx", "xls", "csv"], key="stmt_upload"
    )

    if uploaded is not None and st.session_state.get("_last_upload_name") != uploaded.name:
        # New file selected -> reset the wizard state.
        reset_workflow_state()
        st.session_state["_last_upload_name"] = uploaded.name
        ext = uploaded.name.lower().split('.')[-1]
        with st.spinner("Extracting rows..."):
            try:
                rows, page_images = extract_raw_rows(uploaded, ext)
            except Exception as e:
                st.error(f"Failed to read file: {e}")
                return
        if not rows:
            st.error("No data rows could be extracted from this file.")
            return
        st.session_state["raw_rows"] = rows
        st.session_state["page_images"] = page_images
        st.session_state["ext"] = ext
        st.session_state["step"] = 1

    if "raw_rows" not in st.session_state:
        return

    rows = st.session_state["raw_rows"]
    ext = st.session_state["ext"]

    # ---------------- Step 1: pick header / data-start row ----------------
    if st.session_state.get("step", 1) == 1:
        st.markdown("### Step 1 · Where does your real data start?")
        st.caption(
            "Pick the row number where your Reference / Debit / Credit header actually "
            "begins. Everything above it will be ignored."
        )

        if st.session_state.get("page_images"):
            with st.expander("PDF page preview", expanded=False):
                cols = st.columns(min(4, len(st.session_state["page_images"])))
                for i, img in enumerate(st.session_state["page_images"][:8]):
                    with cols[i % len(cols)]:
                        st.image(img, caption=f"Page {i + 1}", use_container_width=True)

        max_cols = max((len(r) for r in rows), default=0)
        preview_n = min(40, len(rows))
        padded_preview = [pad_row(r, max_cols) for r in rows[:preview_n]]
        preview_df = pd.DataFrame(
            padded_preview, columns=[f"Col {i + 1}" for i in range(max_cols)]
        )
        st.dataframe(preview_df, use_container_width=True, height=320)
        if len(rows) > preview_n:
            st.caption(f"Showing first {preview_n} of {len(rows)} rows.")

        default_guess = guess_header_row(rows)
        header_idx = st.number_input(
            "Header / data-start row number (from the table above, 0 = first row)",
            min_value=0, max_value=max(len(rows) - 1, 0),
            value=min(default_guess, max(len(rows) - 1, 0)), step=1,
        )

        if st.button("Use this row as header →", type="primary"):
            header_width = len(rows[header_idx])
            header = [str(c).strip() if c is not None else '' for c in rows[header_idx]]

            seen = {}
            clean_header = []
            for i, h in enumerate(header):
                name = h if h else f"Column {i + 1}"
                if name in seen:
                    seen[name] += 1
                    name = f"{name} ({seen[name]})"
                else:
                    seen[name] = 1
                clean_header.append(name)

            auto_map_guess = map_columns(clean_header)
            if 'reference' in auto_map_guess:
                merge_col_idx = clean_header.index(auto_map_guess['reference'])
            else:
                merge_col_idx = header_width - 1

            date_idx = clean_header.index(auto_map_guess['date']) if 'date' in auto_map_guess else None
            debit_idx = clean_header.index(auto_map_guess['debit']) if 'debit' in auto_map_guess else None
            credit_idx = clean_header.index(auto_map_guess['credit']) if 'credit' in auto_map_guess else None

            raw_data_rows = rows[header_idx + 1:]
            merged_rows = merge_wrapped_continuation_rows(
                raw_data_rows, date_idx=date_idx, debit_idx=debit_idx,
                credit_idx=credit_idx, narration_idx=merge_col_idx
            )
            data = [pad_row(r, header_width, merge_into=merge_col_idx) for r in merged_rows]

            st.session_state["df_raw"] = pd.DataFrame(data, columns=clean_header)
            st.session_state["step"] = 2
            st.rerun()
        return

    # ---------------- Step 2: confirm columns ----------------
    if st.session_state.get("step") == 2:
        df_raw = st.session_state["df_raw"]
        auto_map = map_columns(df_raw.columns)
        columns = list(df_raw.columns)

        st.markdown("### Step 2 · Confirm your columns")
        st.caption(
            "Auto-detected values are pre-filled below. If your bank uses different "
            "wording (e.g. 'Payment', 'Receipt', 'Recipient'), just change the dropdown."
        )

        c1, c2, c3 = st.columns(3)
        with c1:
            ref_col = st.selectbox(
                "Reference / Narration column", columns,
                index=columns.index(auto_map['reference']) if 'reference' in auto_map else 0,
            )
        with c2:
            debit_options = [NONE_OPTION] + columns
            debit_default = auto_map.get('debit', NONE_OPTION)
            debit_col = st.selectbox(
                "Debit (money out) column", debit_options,
                index=debit_options.index(debit_default) if debit_default in debit_options else 0,
            )
        with c3:
            credit_options = [NONE_OPTION] + columns
            credit_default = auto_map.get('credit', NONE_OPTION)
            credit_col = st.selectbox(
                "Credit (money in) column", credit_options,
                index=credit_options.index(credit_default) if credit_default in credit_options else 0,
            )

        st.markdown("**Data preview (first 8 rows)**")
        st.dataframe(df_raw.head(8), use_container_width=True)

        bcol1, bcol2 = st.columns([1, 1])
        with bcol1:
            if st.button("← Back"):
                st.session_state["step"] = 1
                st.rerun()
        with bcol2:
            if st.button("Build table →", type="primary"):
                if debit_col == NONE_OPTION and credit_col == NONE_OPTION:
                    st.warning("Select at least one of Debit or Credit column.")
                elif debit_col == credit_col and debit_col != NONE_OPTION:
                    st.warning("Debit and Credit cannot be the same column.")
                else:
                    col_map = {'reference': ref_col}
                    if debit_col != NONE_OPTION:
                        col_map['debit'] = debit_col
                    if credit_col != NONE_OPTION:
                        col_map['credit'] = credit_col
                    st.session_state["df"] = build_clean_df(df_raw, col_map)
                    st.session_state["step"] = 3
                    st.rerun()
        return

    # ---------------- Step 3: keyword search ----------------
    if st.session_state.get("step") == 3 and "df" in st.session_state:
        df = st.session_state["df"]
        st.success(f"Table ready · {len(df)} rows")

        available_cols = [c for c in ['Debit', 'Credit'] if c in df.columns]

        s1, s2, s3 = st.columns([2, 1, 1])
        with s1:
            keyword = st.text_input("Keyword")
        with s2:
            col = st.selectbox("Column", available_cols)
        with s3:
            st.write("")
            st.write("")
            search_clicked = st.button("🔍 Search", type="primary")

        if st.button("↺ Start over with a new file"):
            reset_workflow_state()
            st.rerun()

        if search_clicked:
            if not keyword.strip():
                st.warning("Enter a keyword to search for.")
            else:
                mask = df['Reference'].str.contains(keyword, case=False, na=False, regex=False)
                matches = df[mask]
                amounts = matches[col].dropna()
                amounts = amounts[amounts > 0]

                if amounts.empty:
                    st.info(f"No '{col}' entries found for keyword '{keyword}'.")
                else:
                    amt_str = ''.join(
                        [f"+{int(a) if float(a) == int(a) else a}" for a in amounts]
                    )
                    st.text_area("Amounts", amt_str, height=100)
                    st.markdown(f"**Found {len(amounts)} entries · Total: ₹{amounts.sum():,.2f}**")

                    result_table = matches[matches[col] > 0][['Reference', col]]
                    st.dataframe(result_table, use_container_width=True)

                    csv_bytes = result_table.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        "Download matches as CSV", data=csv_bytes,
                        file_name="matches.csv", mime="text/csv",
                    )


def converter_tab():
    st.subheader("PDF → Excel")
    pdf_file = st.file_uploader("Select PDF to convert", type=["pdf"], key="pdf2xlsx")
    if pdf_file is not None and st.button("Convert to Excel"):
        try:
            with st.spinner("Converting..."):
                buf, row_count = convert_pdf_to_excel_bytes(pdf_file)
            st.success(f"Converted successfully! {row_count} rows.")
            st.download_button(
                "Download Excel file", data=buf,
                file_name=pdf_file.name.rsplit('.', 1)[0] + ".xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Conversion failed: {e}")

    st.divider()

    st.subheader("Excel/CSV → PDF")
    if not REPORTLAB_AVAILABLE:
        st.warning("reportlab is not installed -- add it to requirements.txt to enable this tool.")
    table_file = st.file_uploader(
        "Select Excel/CSV to convert", type=["xlsx", "xls", "csv"], key="tbl2pdf"
    )
    if table_file is not None and st.button("Convert to PDF", disabled=not REPORTLAB_AVAILABLE):
        ext = table_file.name.lower().split('.')[-1]
        try:
            with st.spinner("Converting..."):
                buf, row_count = convert_table_to_pdf_bytes(table_file, ext)
            st.success(f"Converted successfully! {row_count} rows.")
            st.download_button(
                "Download PDF file", data=buf,
                file_name=table_file.name.rsplit('.', 1)[0] + ".pdf",
                mime="application/pdf",
            )
        except Exception as e:
            st.error(f"Conversion failed: {e}")


def main_app():
    st.title("Bank Statement Keyword Amount Extractor")
    st.caption("Search transaction amounts by keyword · Convert PDF ⇄ Excel")

    tab1, tab2 = st.tabs(["🔍 Keyword Search", "🔁 Format Converter"])
    with tab1:
        keyword_search_tab()
    with tab2:
        converter_tab()

    st.divider()
    st.caption(f"Developed by {DEVELOPER_NAME}")


if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    login_gate()
else:
    main_app()
