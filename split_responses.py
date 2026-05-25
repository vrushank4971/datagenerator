"""
split_responses.py
-------------------
Splits the Chapter 10 Google Form responses Excel into 3 clean result PDFs
matching the Satsang Vihar reference format:

  1. Balika PDF — Sr. No. | બાલિકાનું નામ | મંડળનું નામ | Score
  2. ND1 PDF    — Sr. No. | ક્ષેત્ર | મંડળનું નામ | બાળકનું નામ | ધોરણ | Score
  3. ND2 PDF    — Sr. No. | ક્ષેત્ર | મંડળનું નામ | બાળકનું નામ | ધોરણ | Score

Usage:
  1. Place this script + responses.xlsx in the same folder.
  2. pip install pandas openpyxl reportlab requests
  3. python split_responses.py
"""

import re
import sys
from pathlib import Path

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle

# ---------- CONFIG ----------
SHEET_NAME = "Form responses 1"
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

# ---------- FONT ----------
FONT_URL = (
    "https://github.com/google/fonts/raw/main/ofl/"
    "notosansgujarati/NotoSansGujarati%5Bwdth%2Cwght%5D.ttf"
)
FONT_NAME = "NotoGujarati"


def ensure_font():
    Path(FONTS_DIR).mkdir(exist_ok=True)
    font_path = Path(FONTS_DIR) / "NotoSansGujarati.ttf"
    if not font_path.exists():
        print(f"Downloading Gujarati font (first run only) → {font_path}")
        import requests
        r = requests.get(FONT_URL, timeout=60)
        r.raise_for_status()
        font_path.write_bytes(r.content)
    pdfmetrics.registerFont(TTFont(FONT_NAME, str(font_path)))


def extract_chapter_label(input_file: Path) -> str:
    stem = input_file.stem
    label = re.sub(r'\s*\(?\s*responses?\s*\)?\s*$', '', stem, flags=re.IGNORECASE).strip()
    return label


# ---------- DATA PREP ----------

def fmt_score(v):
    if pd.isna(v) or v == "":
        return ""
    try:
        return f"{int(float(v))} / {MAX_SCORE}"
    except (ValueError, TypeError):
        return str(v)


def build_balika_df(df):
    sub = df[df[COL_BAL_NAME].notna() & (df[COL_BAL_NAME].astype(str).str.strip() != "")].copy()
    sub = sub.sort_values(by=COL_BAL_MANDAL, kind="stable", na_position="last")
    return pd.DataFrame({
        "Sr. No.":         range(1, len(sub) + 1),
        "બાલિકાનું નામ":   sub[COL_BAL_NAME].astype(str).values,
        "મંડળનું નામ":     sub[COL_BAL_MANDAL].astype(str).values,
        "Score":           [fmt_score(v) for v in sub[COL_SCORE].values],
    })


def build_nd_df(df, mandal_col, name_col, std_col, kshetra_label):
    sub = df[df[name_col].notna() & (df[name_col].astype(str).str.strip() != "")].copy()
    sub = sub.sort_values(by=mandal_col, kind="stable", na_position="last")
    return pd.DataFrame({
        "Sr. No.":       range(1, len(sub) + 1),
        "ક્ષેત્ર":        [kshetra_label] * len(sub),
        "મંડળનું નામ":   sub[mandal_col].astype(str).values,
        "બાળકનું નામ":   sub[name_col].astype(str).values,
        "ધોરણ":          sub[std_col].astype(str).fillna("").values,
        "Score":         [fmt_score(v) for v in sub[COL_SCORE].values],
    })


# ---------- PDF RENDERING ----------

def render_pdf(df, out_path, title_text):
    if df.empty:
        print(f"  ⚠  No data, skipping {out_path}")
        return

    doc = SimpleDocTemplate(
        str(out_path),
        pagesize=A4,
        leftMargin=15 * mm, rightMargin=15 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm,
        title=title_text,
    )

    title_style = ParagraphStyle(
        "TitleGu", fontName=FONT_NAME, fontSize=14,
        alignment=1, spaceAfter=10, leading=18,
    )
    cell_style = ParagraphStyle(
        "CellGu", fontName=FONT_NAME, fontSize=9, leading=11,
    )
    head_style = ParagraphStyle(
        "HeadGu", fontName=FONT_NAME, fontSize=9, leading=11,
        textColor=colors.white, alignment=1,
    )

    story = [Paragraph(title_text, title_style)]

    headers = [Paragraph(str(c), head_style) for c in df.columns]
    rows = [headers]
    for _, r in df.iterrows():
        rows.append([Paragraph("" if pd.isna(v) else str(v), cell_style) for v in r])

    n_cols = len(df.columns)
    page_w = A4[0] - 30 * mm

    if n_cols == 4:   # Balika
        col_w = [18 * mm, page_w - 18 * mm - 55 * mm - 25 * mm, 55 * mm, 25 * mm]
    elif n_cols == 6: # ND1 / ND2
        col_w = [
            15 * mm,                                                # Sr. No.
            18 * mm,                                                # Kshetra
            45 * mm,                                                # Mandal
            page_w - 15*mm - 18*mm - 45*mm - 22*mm - 25*mm,         # Name (flex)
            22 * mm,                                                # Dhoran
            25 * mm,                                                # Score
        ]
    else:
        col_w = [page_w / n_cols] * n_cols

    table = Table(rows, colWidths=col_w, repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2c3e50")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),    # Sr. No.
        ("ALIGN", (-1, 0), (-1, -1), "CENTER"),  # Score
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.white, colors.HexColor("#f5f7fa")]),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    if n_cols == 6:
        style.append(("ALIGN", (1, 0), (1, -1), "CENTER"))   # Kshetra
        style.append(("ALIGN", (4, 0), (4, -1), "CENTER"))   # Dhoran
    table.setStyle(TableStyle(style))
    story.append(table)
    doc.build(story)
    print(f"  ✓  {out_path} ({len(df)} rows)")


# ---------- MAIN ----------

def main():
    if len(sys.argv) < 2:
        print("Usage: python split_responses.py <filename.xlsx>")
        sys.exit(1)

    input_file = Path(sys.argv[1])
    if not input_file.exists():
        print(f"ERROR: '{input_file}' not found.")
        sys.exit(1)

    print("Setting up font ...")
    ensure_font()

    print(f"Reading {input_file.name} ...")
    df = pd.read_excel(input_file, sheet_name=SHEET_NAME)
    print(f"  Loaded {len(df)} total responses")

    chapter = extract_chapter_label(input_file)
    title_base = f"{TITLE_PREFIX} : {chapter}"

    Path(OUTPUT_DIR).mkdir(exist_ok=True)

    print("Building PDFs ...")
    render_pdf(build_balika_df(df),
               Path(OUTPUT_DIR) / f"{TITLE_PREFIX} {chapter} (Balika).pdf",
               title_base)
    render_pdf(build_nd_df(df, COL_ND1_MANDAL, COL_ND1_NAME, COL_ND1_STD, "ND1"),
               Path(OUTPUT_DIR) / f"{TITLE_PREFIX} {chapter} (ND1).pdf",
               f"{title_base} : ND 1")
    render_pdf(build_nd_df(df, COL_ND2_MANDAL, COL_ND2_NAME, COL_ND2_STD, "ND2"),
               Path(OUTPUT_DIR) / f"{TITLE_PREFIX} {chapter} (ND2).pdf",
               f"{title_base} : ND 2")

    print("\nDone. Files are in ./output/")


if __name__ == "__main__":
    main()
