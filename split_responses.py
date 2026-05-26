"""
split_responses.py
-------------------
Splits the Chapter 10 Google Form responses Excel into 3 clean result PDFs
matching the Satsang Vihar reference format:

  1. Balika PDF — Sr. No. | બાલિકાનું નામ | મંડળનું નામ | Score
  2. ND1 PDF    — Sr. No. | ક્ષેત્ર | મંડળનું નામ | બાળકનું નામ | ધોરણ | Score
  3. ND2 PDF    — Sr. No. | ક્ષેત્ર | મંડળનું નામ | બાળકનું નામ | ધોરણ | Score

Usage:
  1. Place this script in the project folder.
  2. pip install pandas openpyxl fpdf2 uharfbuzz requests
  3. python split_responses.py
"""

import re
import sys
import unicodedata
from pathlib import Path

import pandas as pd
from fpdf import FPDF

# ---------- CONFIG ----------
SHEET_NAME = "Form responses 1"
INPUT_DIR  = "input"
OUTPUT_DIR = "output"
FONTS_DIR  = "fonts"
MAX_SCORE  = 20
TITLE_PREFIX = "Satsang Vihar Online Test Result"

# Exact column names in your Google Form sheet
COL_SCORE       = "Score"
COL_BAL_NAME    = "બાલિકાનું નામ"
COL_BAL_MANDAL  = "મંડળનું નામ"
COL_ND1_MANDAL  = "નડિયાદ 1 ના મંડળનું નામ"
COL_ND1_NAME    = "બાળકનું નામ"
COL_ND1_STD     = "ધોરણ"
COL_ND2_MANDAL  = "નડિયાદ 2 ના મંડળનું નામ"
COL_ND2_NAME    = "બાળકનું નામ 2"
COL_ND2_STD     = "ધોરણ 2"


def ensure_font() -> str:
    font_path = Path(FONTS_DIR) / "HindVadodara-Medium.ttf"
    if not font_path.exists():
        print(f"ERROR: Font not found at {font_path}.")
        print("Please place HindVadodara-Medium.ttf in the fonts/ folder.")
        sys.exit(1)
    return str(font_path)


def nfc(text: str) -> str:
    """NFC-normalise Gujarati Unicode so matras are in canonical order."""
    return unicodedata.normalize("NFC", str(text))


def extract_chapter_label(input_file: Path) -> tuple[str, str]:
    """Return (chapter_label, language) extracted from the filename."""
    stem = input_file.stem
    label = re.sub(r'\s*\(?\s*responses?\s*\)?\s*$', '', stem, flags=re.IGNORECASE).strip()

    lang_match = re.search(
        r'(?:_|\s)(english|gujarati|hindi|marathi|sanskrit|french|spanish|german|chinese|japanese)(?:_|\s|$)',
        label, re.IGNORECASE,
    )
    if lang_match:
        language = lang_match.group(1).capitalize()
        label = re.sub(r'(?:_|\s)' + lang_match.group(1) + r'(?:_|\s|$)', '', label, flags=re.IGNORECASE).strip()
        label = re.sub(r'\s+', '_', label)
    else:
        language = "Gujarati"

    return label, language


# ---------- DATA PREP ----------

def fmt_score(v):
    if pd.isna(v) or v == "":
        return ""
    try:
        return f"{int(float(v))} / {MAX_SCORE}"
    except (ValueError, TypeError):
        return str(v)


def build_balika_df(df: pd.DataFrame) -> pd.DataFrame:
    sub = df[df[COL_BAL_NAME].notna() & (df[COL_BAL_NAME].astype(str).str.strip() != "")].copy()
    sub = sub.sort_values(by=COL_BAL_MANDAL, kind="stable", na_position="last")
    return pd.DataFrame({
        "Sr. No.":         range(1, len(sub) + 1),
        "બાલિકાનું નામ":   [nfc(v) for v in sub[COL_BAL_NAME].astype(str)],
        "મંડળનું નામ":     [nfc(v) for v in sub[COL_BAL_MANDAL].astype(str)],
        "Score":           [fmt_score(v) for v in sub[COL_SCORE]],
    })


def build_nd_df(df: pd.DataFrame, mandal_col, name_col, std_col, kshetra_label) -> pd.DataFrame:
    sub = df[df[name_col].notna() & (df[name_col].astype(str).str.strip() != "")].copy()
    sub = sub.sort_values(by=mandal_col, kind="stable", na_position="last")
    return pd.DataFrame({
        "Sr. No.":       range(1, len(sub) + 1),
        "ક્ષેત્ર":        [kshetra_label] * len(sub),
        "મંડળનું નામ":   [nfc(v) for v in sub[mandal_col].astype(str)],
        "બાળકનું નામ":   [nfc(v) for v in sub[name_col].astype(str)],
        "ધોરણ":          [nfc(v) for v in sub[std_col].astype(str).fillna("")],
        "Score":         [fmt_score(v) for v in sub[COL_SCORE]],
    })


# ---------- PDF RENDERING ----------

def render_pdf(df: pd.DataFrame, out_path: Path, title_text: str, font_path: str):
    if df.empty:
        print(f"  [WARN] No data, skipping {out_path}")
        return

    n_cols = len(df.columns)
    CONTENT_W = 180  # A4 210mm − 2×15mm margins

    if n_cols == 4:   # Balika: Sr.No | Name | Mandal | Score
        col_ws  = [18, CONTENT_W - 18 - 55 - 25, 55, 25]
        d_aligns = ["C", "L", "L", "C"]
    elif n_cols == 6:  # ND1/ND2: Sr.No | Kshetra | Mandal | Name | Dhoran | Score
        flex    = CONTENT_W - 15 - 18 - 45 - 22 - 25
        col_ws  = [15, 18, 45, flex, 22, 25]
        d_aligns = ["C", "C", "L", "L", "C", "C"]
    else:
        col_ws   = [CONTENT_W / n_cols] * n_cols
        d_aligns = ["L"] * n_cols

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_margins(left=15, top=15, right=15)
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.add_font("HV", fname=font_path)
    # Enable HarfBuzz shaping so Gujarati conjuncts and matras render correctly.
    pdf.set_text_shaping(use_shaping_engine=True, script="Gujr", language="GUJ")

    # ── Title ────────────────────────────────────────────────────────────────
    pdf.set_font("HV", size=14)
    pdf.multi_cell(0, 10, title_text, align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)

    ROW_H = 7
    pdf.set_font("HV", size=9)

    # ── Header row ───────────────────────────────────────────────────────────
    pdf.set_fill_color(44, 62, 80)
    pdf.set_text_color(255, 255, 255)
    for col, cw in zip(df.columns, col_ws):
        pdf.cell(cw, ROW_H, nfc(str(col)), border=1, fill=True, align="C")
    pdf.ln()

    # ── Data rows ────────────────────────────────────────────────────────────
    pdf.set_text_color(0, 0, 0)
    for i, (_, row) in enumerate(df.iterrows()):
        if i % 2 == 0:
            pdf.set_fill_color(245, 247, 250)
        else:
            pdf.set_fill_color(255, 255, 255)
        for val, cw, align in zip(row, col_ws, d_aligns):
            pdf.cell(cw, ROW_H, nfc("" if pd.isna(val) else str(val)), border=1, fill=True, align=align)
        pdf.ln()

    pdf.output(str(out_path))
    print(f"  [OK] {out_path} ({len(df)} rows)")


# ---------- MAIN ----------

def main():
    print("Checking font ...")
    font_path = ensure_font()

    Path(INPUT_DIR).mkdir(exist_ok=True)
    Path(OUTPUT_DIR).mkdir(exist_ok=True)

    excel_files = list(Path(INPUT_DIR).glob("*.xlsx")) + list(Path(INPUT_DIR).glob("*.xls"))
    if not excel_files:
        print(f"ERROR: No Excel files found in '{INPUT_DIR}/' folder.")
        sys.exit(1)

    print(f"Found {len(excel_files)} Excel file(s) in {INPUT_DIR}/ folder")

    for input_file in excel_files:
        print(f"\nProcessing {input_file.name} ...")
        try:
            df = pd.read_excel(input_file, sheet_name=SHEET_NAME)
            # NFC-normalise column names so they match the COL_* constants
            df.columns = [unicodedata.normalize("NFC", str(c)) for c in df.columns]
            print(f"  Loaded {len(df)} total responses")

            chapter, language = extract_chapter_label(input_file)
            title_base = f"{TITLE_PREFIX} : {chapter}"

            print("Building PDFs ...")
            render_pdf(
                build_balika_df(df),
                Path(OUTPUT_DIR) / f"{language} {TITLE_PREFIX} {chapter} (Balika).pdf",
                title_base,
                font_path,
            )
            render_pdf(
                build_nd_df(df, COL_ND1_MANDAL, COL_ND1_NAME, COL_ND1_STD, "ND1"),
                Path(OUTPUT_DIR) / f"{language} {TITLE_PREFIX} {chapter} (ND1).pdf",
                f"{title_base} : ND 1",
                font_path,
            )
            render_pdf(
                build_nd_df(df, COL_ND2_MANDAL, COL_ND2_NAME, COL_ND2_STD, "ND2"),
                Path(OUTPUT_DIR) / f"{language} {TITLE_PREFIX} {chapter} (ND2).pdf",
                f"{title_base} : ND 2",
                font_path,
            )
        except Exception as e:
            print(f"  [WARN] Error processing {input_file.name}: {e}")
            continue

    print("\nDone. Files are in ./output/")


if __name__ == "__main__":
    main()
