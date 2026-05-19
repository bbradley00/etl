# GUI app: select 2 PDFs (Main + PO) and a Staffing Bible (Excel, locally OneDrive-synced),
# then produce one color-coded Excel grouped ROLE → SHIFT.
# Keeps B & F formatting from roster, colors C & D by shift:
#   default gold (#FFF2CC), E/E1 orange (#F4B183), 7P/E2 light green (#C6EFCE), N = #CCFFFF.

import os, re, traceback
from openpyxl.styles import Font, Border, Side, Alignment, Protection
from copy import copy

MISSING = []
try:
    from PyPDF2 import PdfReader
except Exception:
    MISSING.append("PyPDF2")
try:
    import pandas as pd
except Exception:
    MISSING.append("pandas")
try:
    from rapidfuzz import process, fuzz
except Exception:
    MISSING.append("rapidfuzz")
try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import PatternFill
except Exception:
    MISSING.append("openpyxl")

APP_TITLE = "Health First – All-in-One Staffing Extractor"

# PDF parsing
# (adapted from earlier staffing script)

BASE_SHIFT_HEADERS = {
    "N", "7A", "E", "7P", "D", "E1", "E2", "NAS", "7A ORIENT", "830 TO 17", "10T19Q"
}
ROLE_PREFIXES = ["RN", "CNA", "NSA", "HUC", "SA", "AA", "LPN", "CCN", "PO"]

# 10T19Q tolerant matcher (case-insensitive, optional internal spaces)
TEN_T_19_Q_RE = re.compile(r"^\s*10\s*[Tt]\s*19\s*[Qq]\s*$")

# shift + role on one line (allow optional space between shift and role)
header_both_re = re.compile(
    r"^(?P<shift>(?:7A|E|E1|E2|7P|N|D|NAS|7A ORIENT|830\s*TO\s*17|10\s*[-Tt]?\s*19\s*[Qq]))\s+"
    r"(?P<role>RN|CNA|NSA|HUC|SA|AA|LPN|CCN|PO)\b",
    re.IGNORECASE
)

PLANNED_RE = re.compile(r"^\s*Planned\b", re.IGNORECASE)
VARIANCE_RE = re.compile(r"^\s*Variance\b", re.IGNORECASE)
NOTES_PAT = re.compile(r"(x\s+FLT\s+POOL/[A-Z]{2,3}\b.*)$", re.IGNORECASE)
TIME_RANGE_PAT = re.compile(r"\d{1,2}:\d{2}\s*[AP]M\s*[-–—]\s*\d{1,2}:\d{2}\s*[AP]M.*$", re.IGNORECASE)
ORIENT_NOTE_PAT = re.compile(r"\bORIENT(?:ATION)?\b.*$", re.IGNORECASE)
COMPACT_SHIFT_PAT = re.compile(r"^[A-Z0-9][A-Z0-9 :/\-]{0,12}$")

def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def _nospace_upper(s: str) -> str:
    return re.sub(r"\s+", "", (s or "").replace("\u00A0", " ")).upper()

def is_shift_header(line: str, next_line: str | None) -> bool:
    if not line:
        return False
    up = line.upper().strip()

    # Direct match against known headers
    if up in {h.upper() for h in BASE_SHIFT_HEADERS}:
        return True

    # Handle 10T19Q variants (extra spaces or dash)
    if re.fullmatch(r"(?:7A|E|E1|E2|7P|N|D|NAS|7A ORIENT|830\s*TO\s*17|10T19Q)", line.strip().upper()):
        return True

    # Split-line header like "10T19Q" then "RN"
    if up in {"10T19Q", "10T19", "10 19Q", "10T 19Q"} and next_line and next_line.upper().strip() in ROLE_PREFIXES:
        return True

    # Compact guess: treat as header if next line starts with a role
    if COMPACT_SHIFT_PAT.match(up) and next_line:
        for rp in ROLE_PREFIXES:
            if next_line.startswith(rp + " ") or next_line == rp or next_line.startswith(rp + ","):
                return True
    return False

def extract_notes_and_clean(s: str) -> tuple[str, str]:
    notes = ""
    m = NOTES_PAT.search(s)
    if m:
        notes = norm(m.group(1))
        s = s.replace(m.group(1), "")

    # remove any trailing ORIENT/ORIENTATION fragments (e.g., "... PT OBS/PO ORIENTATION")
    s = ORIENT_NOTE_PAT.sub("", s)

    # strip trailing time-range text off the name
    s = TIME_RANGE_PAT.sub("", s).strip(" ,")
    s = re.sub(r"\s{2,}", " ", s)
    return s.strip(), notes

def parse_role_name(line: str) -> tuple[str | None, str | None, str]:
    if not line:
        return None, None, ""
    role = None
    for rp in ROLE_PREFIXES:
        if line.startswith(rp + " ") or line == rp or line.startswith(rp + ","):
            role = rp
            break
    if not role:
        return None, None, ""
    rest = line[len(role):].strip()
    if not rest:
        return role, None, ""
    rest, notes = extract_notes_and_clean(rest)
    name = rest if rest else None
    return role, name, notes

def join_wrapped_name(curr: str, nxt: str) -> str | None:
    if not curr or not nxt:
        return None
    c, n = curr.strip(), nxt.strip()
    if c.endswith(","):
        if re.fullmatch(r"[A-Z][A-Z'\-]+(?:\s+[A-Z][A-Z'\-]+)*", n):
            return f"{c} {n}"
    if re.fullmatch(r"[A-Z][A-Z'\-]+,?", c) and re.fullmatch(r"[A-Z][A-Z'\-]+(?:\s+[A-Z][A-Z'\-]+)*", n):
        if not c.endswith(","):
            c = c + ","
        return f"{c} {n}"
    return None

def parse_pdf_to_df(pdf_path: str) -> pd.DataFrame:
    reader = PdfReader(pdf_path)
    text_pages = [page.extract_text() or "" for page in reader.pages]
    raw_lines = [ln.rstrip() for ln in "\n".join(text_pages).splitlines()]
    lines = [ln.strip() for ln in raw_lines if ln.strip()]

    records = []
    current_shift = None
    current_role_hint = None
    block_emitted = False

    # Catch shift+role headers like "10T19Q RN"
    header_both_re = re.compile(
        r"^(?P<shift>(?:7A|E|E1|E2|7P|N|D|NAS|7A ORIENT|830\s*TO\s*17|10T19Q))\s+"
        r"(?P<role>RN|CNA|NSA|HUC|SA|AA|LPN|CCN|PO)\b",
        re.IGNORECASE
    )

    i = 0
    while i < len(lines):
        line = lines[i]
        next_line = lines[i + 1].strip() if i + 1 < len(lines) else ""

        # --- detect "Shift Role" headers like "10T19Q RN" ---
        m_both = header_both_re.match(line)
        if m_both:
            current_shift = m_both.group("shift").upper()
            current_role_hint = m_both.group("role").upper()
            block_emitted = False
            i += 1
            continue

        # --- detect shift-only header (e.g., "10T19Q" alone) ---
        if re.fullmatch(r"(7A|E|E1|E2|7P|N|D|NAS|7A ORIENT|830 TO 17|10T19Q)", line.strip().upper()):
            current_shift = line.strip().upper()
            # if next line is a role (RN/CNA/...), capture it
            next_up = next_line.upper()
            if next_up in ROLE_PREFIXES:
                current_role_hint = next_up
                i += 1  # consume that role line
            else:
                current_role_hint = None
            block_emitted = False
            i += 1
            continue

        # --- skip planned/variance/time junk ---
        if re.search(r"\bPlanned\b|\bVariance\b|\d{1,2}:\d{2}\s*[AP]M", line, re.IGNORECASE):
            i += 1
            continue

        # --- "Planned" placeholder block ---
        if PLANNED_RE.match(line):
            if current_shift and not block_emitted:
                records.append({
                    "SHIFT": current_shift,
                    "ROLE": "",
                    "NAME": "<Open Shift>",
                    "NOTES": "none assigned"
                })
                block_emitted = True
            current_role_hint = None
            i += 1
            continue

        # --- lines starting with a role prefix (RN, CNA, etc.) ---
        role, name, notes = parse_role_name(line)
        if role:
            if name:
                joined = join_wrapped_name(name, next_line)
                if joined:
                    name = joined
                    i += 1
                records.append({
                    "SHIFT": current_shift or "",
                    "ROLE": role,
                    "NAME": name,
                    "NOTES": notes
                })
                block_emitted = True
            else:
                current_role_hint = role
            i += 1
            continue

        # --- name-only lines under a role block ---
        if current_role_hint and line and not PLANNED_RE.match(line):
            # 1) ORIENTATION meta-lines: attach to previous record as a note, do NOT become a person
            if re.search(r"\bORIENT(?:ATION)?\b", line, re.IGNORECASE):
                if records:
                    prev_notes = records[-1].get("NOTES", "") or ""
                    note_text = norm(line)  # e.g. "P - ORIENTATION"
                    records[-1]["NOTES"] = f"{prev_notes}; {note_text}" if prev_notes else note_text
                i += 1
                continue

            # 2) Time-range lines (e.g. "6:42 AM-3:12 PM x PT OBS/PO"):
            #    also attach the "x PT OBS/PO" part as notes on the previous record.
            if TIME_RANGE_PAT.match(line):
                if records:
                    # extract_notes_and_clean strips the time, leaves the "x PT OBS/PO" part in notes
                    _, notes_line = extract_notes_and_clean(line)
                    note_text = notes_line or norm(line)
                    if note_text:
                        prev_notes = records[-1].get("NOTES", "") or ""
                        records[-1]["NOTES"] = f"{prev_notes}; {note_text}" if prev_notes else note_text
                i += 1
                continue

            # 3) Normal name lines (or "<Open Shift>")
            j = join_wrapped_name(line, next_line)
            if j:
                clean, notes2 = extract_notes_and_clean(j)
                if clean and not clean.lower().startswith("<open shift>"):
                    records.append({
                        "SHIFT": current_shift or "",
                        "ROLE": current_role_hint,
                        "NAME": clean,
                        "NOTES": notes2
                    })
                    block_emitted = True
                i += 2
                continue

            clean_line, notes3 = extract_notes_and_clean(line)
            if clean_line:
                name_val = "<Open Shift>" if clean_line.lower().startswith("<open shift>") else clean_line
                records.append({
                    "SHIFT": current_shift or "",
                    "ROLE": current_role_hint,
                    "NAME": name_val,
                    "NOTES": notes3
                })
                block_emitted = True
            i += 1
            continue

        i += 1

    if current_shift and not block_emitted:
        records.append({
            "SHIFT": current_shift,
            "ROLE": "",
            "NAME": "<Open Shift>",
            "NOTES": "none assigned"
        })

    df = pd.DataFrame.from_records(records)
    df = df.sort_values(by=["SHIFT", "ROLE", "NAME"], kind="stable").reset_index(drop=True)
    return df



# ---------------- MATCHING ----------------

def parse_last_first(s: str) -> tuple[str, str]:
    if not s: return "", ""
    up = " ".join(str(s).upper().replace("\u00A0", " ").split())
    if "," in up:
        last, first = up.split(",", 1)
        return last.strip(), first.strip()
    parts = up.split(" ")
    if len(parts) == 1: return parts[0], ""
    return parts[-1], " ".join(parts[:-1])

def roster_name_tokens(row_text: str) -> list[str]:
    up = " ".join((row_text or "").upper().replace("\u00A0", " ").split())
    parts = re.split(r"[^\w]+", up)
    toks = []
    for t in parts:
        if not t: continue
        if t == "NO": continue
        if len(t) == 1: continue
        toks.append(t)
    return toks

def build_roster_index(df_roster: pd.DataFrame, roster_name_col: str) -> list[list[str]]:
    tokens_per_row = []
    for _, r in df_roster.iterrows():
        tokens_per_row.append(roster_name_tokens(str(r.get(roster_name_col, ""))))
    return tokens_per_row

def fuzzy_match_first_last(toks: list[str], lastL: str, firstL: str) -> bool:
    def prefix_ok(a, b, m):
        s = min(len(a), len(b))
        return s >= m and (a.startswith(b) or b.startswith(a))
    if lastL:
        lastOK = False
        for t in toks:
            if t == lastL or prefix_ok(t, lastL, 4):
                lastOK = True; break
        if not lastOK and len(toks) >= 2:
            for i in range(len(toks) - 1):
                pair = toks[i] + toks[i + 1]
                if pair == lastL or prefix_ok(pair, lastL, 4):
                    lastOK = True; break
        if not lastOK: return False
    if firstL:
        firstOK = False
        for t in toks:
            if t == firstL or prefix_ok(t, firstL, 3):
                firstOK = True; break
        if not firstOK and len(toks) >= 2:
            for i in range(len(toks) - 1):
                pair = toks[i] + toks[i + 1]
                if pair == firstL or prefix_ok(pair, firstL, 3):
                    firstOK = True; break
        if not firstOK: return False
    return True

from rapidfuzz import process, fuzz, utils
import re

def normalize_person_name(s: str) -> str:
    """
    Normalize a person's name for comparison:
    - Uppercase
    - Remove hyphens, punctuation, and excess spaces
    - Remove trailing comments or departments after a dash or ' -IN '
    """
    if not s:
        return ""
    s = str(s).upper().replace("\u00A0", " ")
    s = re.sub(r"\s*[-–—]\s*IN .*", "", s)   # remove " -IN SMS AS ..." etc.
    s = re.sub(r"[-–—]", "", s)              # remove hyphens
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def match_pdf_to_roster(df_pdf: pd.DataFrame, df_roster: pd.DataFrame,
                        roster_name_col="D_NAME", coarse_threshold=85) -> pd.DataFrame:
    """
    Fuzzy match PDF names to roster names...ignroes hyphens and trailing notes, but only after a last/first sanity filter.
    """
    # Ensure we have a name column
    if roster_name_col not in df_roster.columns:
        guess = next((c for c in df_roster.columns
                      if str(c).upper() in ("NAME", "EMPLOYEE NAME", "FULLNAME", "FULL NAME", "STAFF NAME")), None)
        if not guess:
            raise RuntimeError("Could not locate a name column in roster; add 'D_NAME' or rename.")
        df_roster = df_roster.rename(columns={guess: roster_name_col})

    # Raw + normalized roster names
    roster_names_raw = df_roster[roster_name_col].astype(str).fillna("").tolist()
    roster_norm = [normalize_person_name(x) for x in roster_names_raw]

    # Precompute tokens per roster row for last/first sanity checks
    roster_tokens = [roster_name_tokens(name) for name in roster_names_raw]

    matched_rows = []

    for _, r in df_pdf.iterrows():
        # Get raw value from NAME column
        pdf_name_raw = r.get("NAME", "")

        # ---- 0) Skip actual NaN / missing names ----
        if pd.isna(pdf_name_raw):
            continue

        # Now safely convert to string
        pdf_name_raw = str(pdf_name_raw)

        # ---- 1) Skip empties and explicit open shifts ----
        if not pdf_name_raw.strip() or pdf_name_raw.upper().startswith("<OPEN SHIFT>"):
            # Don't try to match open shifts or blank lines
            continue

        # Normalized name (uppercased, punctuation stripped)
        pdf_norm = normalize_person_name(pdf_name_raw)

        # If normalization collapses to nothing or plain 'NAN', it's garbage
        if not pdf_norm or pdf_norm == "NAN":
            continue

        lastL, firstL = parse_last_first(pdf_name_raw)

        # --- 2) Filter roster rows by last/first compatibility ---
        candidate_indices = [
            i for i, toks in enumerate(roster_tokens)
            if fuzzy_match_first_last(toks, lastL, firstL)
        ]

        if not candidate_indices:
            # No one in the roster passes the last/first sanity check → treat as UNMATCHED
            continue

        # --- 3) Among candidates, pick the best fuzzy token_set_ratio match ---
        best_idx = None
        best_score = -1
        for i in candidate_indices:
            score = fuzz.token_set_ratio(pdf_norm, roster_norm[i])
            if score > best_score:
                best_score = score
                best_idx = i

        if best_idx is None or best_score < coarse_threshold:
            # Not a strong enough match → treat as UNMATCHED
            continue

        idx = best_idx

        # --- Build output row (keep your existing behavior) ---
        out = {**r.to_dict()}
        for j, col in enumerate(df_roster.columns[:6]):
            out[f"ROSTER_{j}"] = df_roster.iloc[idx][col]
        out["MATCH_SCORE"] = int(best_score)
        out["_ROSTER_INDEX"] = int(idx)
        out["MATCHED_NAME"] = roster_names_raw[idx]
        matched_rows.append(out)

    return pd.DataFrame(matched_rows)


# ---------------- OUTPUT / COLORING ----------------

FILL_DEFAULT = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")  # Gold Accent4 L80
FILL_ORANGE  = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")  # E/E1
FILL_GREEN   = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")  # 7P/E2
FILL_BLUE    = PatternFill(start_color="CCFFFF", end_color="CCFFFF", fill_type="solid")  # N

def choose_fill_for_shift(shift: str):
    """Return fill color based on shift code, tolerant of PDF artifacts and invisible characters."""
    if not shift:
        return FILL_DEFAULT

    # normalize all whitespace + unicode variants
    s = str(shift).upper()
    s = s.replace("\u00A0", "")   # non-breaking space
    s = s.replace("\u202F", "")   # narrow no-break space
    s = s.replace("\u2009", "")   # thin space
    s = s.replace("\u200A", "")
    s = s.replace("\u200B", "")   # zero-width space
    s = s.replace("\uFEFF", "")
    s = re.sub(r"[^A-Z0-9]", "", s)  # strip all punctuation and leftover junk

    # --- now match cleanly ---
    if s in ("E", "E1"):
        return FILL_ORANGE
    if s in ("7P", "E2"):
        return FILL_GREEN
    if s == "N":
        return FILL_BLUE
    if s in ("7A", "D", "10T19Q", "830TO17"):
        return FILL_DEFAULT
    return FILL_DEFAULT

def copy_style(dst_cell, src_cell):
    try:
        if src_cell.has_style:
            dst_cell.font = src_cell.font
            dst_cell.border = src_cell.border
            from copy import copy
            dst_cell.fill = copy(src_cell.fill)
            dst_cell.number_format = src_cell.number_format
            dst_cell.protection = src_cell.protection
            dst_cell.alignment = src_cell.alignment
    except Exception:
        pass

def write_grouped_workbook(
    matched: pd.DataFrame,
    roster_xlsx_path: str,
    roster_sheet="CRT STAFFING",
    out_path="Matched_Output.xlsx",
    unmatched_df: pd.DataFrame | None = None
):
    """Build output workbook grouped by ROLE/SHIFT. Also add unmatched names at top."""

    # --- Load source roster for formatting reference ---
    wb_src = load_workbook(roster_xlsx_path, data_only=False)
    if roster_sheet not in wb_src.sheetnames:
        wb_src.close()
        raise RuntimeError(f"Sheet '{roster_sheet}' not found in roster.")
    ws_src = wb_src[roster_sheet]

    # --- Create output workbook ---
    wb = Workbook()
    ws = wb.active
    ws.title = "Match Found"

    # Headers
    headers = ["SHIFT", "ROSTER A", "ROSTER B", "ROSTER C", "ROSTER D", "ROSTER E", "ROSTER F"]
    ws.append(headers)
    ws.freeze_panes = "B2"
    for col in ["A", "B", "C", "D", "E", "F", "G"]:
        ws.column_dimensions[col].width = 20

    row_cursor = 2  # next writable row after header

    # ---- Add UNMATCHED section at the top ----
    if unmatched_df is not None and not unmatched_df.empty:
        # pattern for LASTNAME, FIRSTNAME
        name_pat = re.compile(r"^[A-Z][A-Z'\-]+,\s*[A-Z][A-Z'\-]+", re.IGNORECASE)
        # exclude SA/AA roles
        filtered = unmatched_df[
            (~unmatched_df["ROLE"].isin(["SA", "AA"])) &
            unmatched_df["NAME"].astype(str).apply(lambda x: bool(name_pat.match(x)))
        ]

        if not filtered.empty:
            ws.append(["UNMATCHED ENTRIES (NAME, NAME FORMAT — EXCLUDING SA/AA)"])
            ws.cell(ws.max_row, 1).fill = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
            ws.append(["SHIFT", "ROLE", "NAME"])
            for _, r in filtered.iterrows():
                ws.append([
                    r.get("SHIFT", ""),
                    r.get("ROLE", ""),
                    r.get("NAME", "")
                ])
            ws.append([""])  # spacer
            row_cursor = ws.max_row + 1

    # ---- Grouped ROLE → SHIFT below ----
    matched_sorted = matched.sort_values(
        ["ROLE", "SHIFT", "NAME", "MATCH_SCORE"], ascending=[True, True, True, False]
    )

    for role, g_role in matched_sorted.groupby("ROLE", sort=True):
        ws.append([f"{role or '(No Role)'}"])
        role_row = ws.max_row
        ws.cell(role_row, 1).fill = PatternFill(start_color="DDD9C3", end_color="DDD9C3", fill_type="solid")

        for shift, g_shift in g_role.groupby("SHIFT", sort=True):
            ws.append([f"Shift: {shift}"])
            ws.cell(ws.max_row, 1).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

            for _, row in g_shift.iterrows():
                ws.append(["", "", "", "", "", "", ""])
                out_r = ws.max_row
                shift_val = str(row.get("SHIFT", ""))  # parsed shift from PDF (N / 7P / E / etc.)

                roster_index = int(row["_ROSTER_INDEX"]) if "_ROSTER_INDEX" in row else None
                if roster_index is not None:
                    src_row = 2 + roster_index

                    # 1) baseline gold for the whole row A..G
                    for col in range(1, 8):
                        ws.cell(out_r, col).fill = FILL_DEFAULT

                    # precompute the fill we want based on SHIFT
                    shift_fill = choose_fill_for_shift(shift_val)

                    for j in range(6):  # source A..F → output B..G
                        src_cell = ws_src.cell(src_row, j + 1)
                        dst_cell = ws.cell(out_r, j + 2)

                        # ---- VALUE COPY ----
                        # For the "shift" column (roster col C -> j == 2), override with today's shift
                        if j == 2:
                            dst_cell.value = shift_val  # use parsed shift instead of usual CRT value
                        else:
                            dst_cell.value = src_cell.value  # copy normal value

                        # ---- SAFE STYLE COPY ----
                        if src_cell.has_style:
                            dst_cell.font = copy(src_cell.font)
                            dst_cell.border = copy(src_cell.border)
                            dst_cell.alignment = copy(src_cell.alignment)
                            dst_cell.number_format = src_cell.number_format
                            dst_cell.protection = copy(src_cell.protection)

                            # LOCATION column (roster F -> j == 5): keep CRT STAFFING background fill
                            if j == 5 and src_cell.fill:
                                dst_cell.fill = copy(src_cell.fill)

                            # SKILL column (roster A -> j == 0) for RNs: keep SKILL banding
                            elif j == 0 and role == "RN" and src_cell.fill:
                                dst_cell.fill = copy(src_cell.fill)

                    # color ROSTER C & ROSTER D by SHIFT (Excel cols D & E)
                    for col in (4, 5):  # columns D and E
                        cell = ws.cell(out_r, col)
                        cell.fill = shift_fill

            ws.append([""])
        ws.append([""])

    wb_src.close()
    wb.save(out_path)
    return out_path



# ---------------- Streamlit Web UI ----------------

import io
import tempfile
import streamlit as st

APP_TITLE = "Health First – All-in-One Staffing Extractor"

st.set_page_config(page_title="Staffing Sheet Creator", layout="centered")
st.title("Staffing Sheet Creator")
st.write("Upload the Main Staffing PDF, PO Staffing PDF, and Staffing Bible Excel file. The app will generate a matched, color-coded Excel output.")

if MISSING:
    st.error("Missing packages: " + ", ".join(MISSING))
    st.stop()

main_pdf_file = st.file_uploader("Main Staffing PDF (RN/CNA/HUC)", type=["pdf"])
po_pdf_file = st.file_uploader("PO Staffing PDF", type=["pdf"])
roster_file = st.file_uploader("Staffing Bible Excel", type=["xlsx", "xlsm"])

coarse_threshold = st.slider("Match strictness", min_value=70, max_value=100, value=85, step=1)


def save_uploaded_to_temp(uploaded_file, suffix):
    """Save a Streamlit upload to a temp file and return its path."""
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.getbuffer())
    tmp.flush()
    tmp.close()
    return tmp.name


def build_output(main_upload, po_upload, roster_upload, threshold):
    main_path = save_uploaded_to_temp(main_upload, ".pdf")
    po_path = save_uploaded_to_temp(po_upload, ".pdf")
    roster_suffix = ".xlsm" if roster_upload.name.lower().endswith(".xlsm") else ".xlsx"
    roster_path = save_uploaded_to_temp(roster_upload, roster_suffix)

    df_main = parse_pdf_to_df(main_path)
    df_po = parse_pdf_to_df(po_path)
    df_all = pd.concat([df_main, df_po], ignore_index=True)

    df_roster = pd.read_excel(roster_path, sheet_name="CRT STAFFING")
    if df_roster.shape[1] >= 4:
        dcol = df_roster.columns[3]
        if "D_NAME" not in df_roster.columns:
            df_roster = df_roster.rename(columns={dcol: "D_NAME"})

    matched = match_pdf_to_roster(df_all, df_roster, roster_name_col="D_NAME", coarse_threshold=threshold)
    if matched.empty:
        raise RuntimeError("No matches found. Check the name column in roster column D and confirm the PDFs are the right ones.")

    try:
        pdf_people = df_all[~df_all["NAME"].str.contains("<OPEN SHIFT>", case=False, na=False)][["SHIFT", "ROLE", "NAME"]].copy()
        matched_names = set(matched["NAME"].astype(str))
        unmatched_df = pdf_people[~pdf_people["NAME"].astype(str).isin(matched_names)].copy()
    except Exception:
        unmatched_df = None

    out_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    out_tmp.close()
    write_grouped_workbook(
        matched,
        roster_path,
        roster_sheet="CRT STAFFING",
        out_path=out_tmp.name,
        unmatched_df=unmatched_df,
    )

    with open(out_tmp.name, "rb") as f:
        output_bytes = f.read()

    return output_bytes, len(df_all), len(matched)


ready = main_pdf_file is not None and po_pdf_file is not None and roster_file is not None

if st.button("Run and create Excel", disabled=not ready):
    try:
        with st.spinner("Parsing PDFs, matching names, and building Excel..."):
            output_bytes, parsed_count, matched_count = build_output(
                main_pdf_file,
                po_pdf_file,
                roster_file,
                coarse_threshold,
            )
        st.success(f"Done. Parsed {parsed_count} PDF entries and matched {matched_count} roster rows.")
        st.download_button(
            label="Download matched Excel",
            data=output_bytes,
            file_name="MATCHED_STAFFING_PLUS_PO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(str(e))
        with st.expander("Show technical details"):
            st.code(traceback.format_exc())

if not ready:
    st.info("Upload all three files to enable the run button.")
