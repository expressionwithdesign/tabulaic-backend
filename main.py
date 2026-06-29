import os
import re
import time
import uuid
import tempfile
import pandas as pd
import anthropic

from fastapi import FastAPI, File, UploadFile, HTTPException, Header
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------

app = FastAPI(title="Tabulaic API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://tabulaic.com", "http://localhost:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

FREE_ROW_LIMIT = 1          # 1 ASIN free, as shown on your site
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
BATCH_SIZE = 30

# ---------------------------------------------------------------------------
# System prompt (your skill, verbatim)
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """You are an expert Amazon Catalog Optimization Agent. Rewrite each title to comply
with the July 27, 2026 Title Update guidelines.

RULES:
- Title: max 75 chars. Structure: Brand + Product Type + Dimensions/Specs + Pack Count.
- Brand must be first. If no brand found, use "FoldCard".
- Remove filler, promotional adjectives, banned chars (!, $, ?, _, {, }, ^).
- Capitalize major words. Use numerals.
- Item Highlights: max 125 chars. Recapture dropped keywords (weights, coatings, use cases,
  "kids", "school supplies", "decorations", "learning aid", "classroom", "homeschool").

OUTPUT FORMAT for each row — no preamble, no extra text, nothing else:
[ASIN or ROW_NUMBER]
- **New Title (X chars):** [title]
- **Item Highlights (Y chars):** [highlights]
---"""


# ---------------------------------------------------------------------------
# File parsing — handles TSV, CSV, XLSX
# ---------------------------------------------------------------------------

def parse_upload(file_bytes: bytes, filename: str) -> pd.DataFrame:
    ext = filename.lower().rsplit(".", 1)[-1]

    with tempfile.NamedTemporaryFile(suffix=f".{ext}", delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name

    if ext in ("tsv", "txt"):
        df = pd.read_csv(tmp_path, sep="\t")
    elif ext == "csv":
        df = pd.read_csv(tmp_path)
    elif ext in ("xlsx", "xls"):
        df = pd.read_excel(tmp_path)
    else:
        raise ValueError(f"Unsupported file type: {ext}")

    os.unlink(tmp_path)

    # Normalise column names — find ASIN and Title columns flexibly
    df.columns = [c.strip() for c in df.columns]
    col_map = {}
    for col in df.columns:
        cl = col.lower()
        if cl in ("asin", "asin1") and "asin" not in col_map:
            col_map["asin"] = col
        if cl in ("title", "item-name", "item_name", "product title", "product_title") and "title" not in col_map:
            col_map["title"] = col

    if "title" not in col_map:
        raise ValueError("Could not find a title column. Expected: 'Title', 'item-name', or 'product title'.")

    df = df.rename(columns={v: k for k, v in col_map.items()})
    df = df.dropna(subset=["title"]).reset_index(drop=True)
    if "asin" not in df.columns:
        df["asin"] = [f"ROW_{i+1}" for i in range(len(df))]

    return df[["asin", "title"]]


# ---------------------------------------------------------------------------
# Claude processing
# ---------------------------------------------------------------------------

def build_batch_prompt(rows: list[dict]) -> str:
    lines = []
    for row in rows:
        lines.append(f"ASIN: {row['asin']}\nTitle: {row['title']}")
    return "\n\n".join(lines)


def parse_claude_response(text: str) -> list[dict]:
    """Parse Claude's structured output into a list of dicts."""
    results = []
    blocks = [b.strip() for b in text.split("---") if b.strip()]

    title_re = re.compile(r"\*\*New Title \((\d+) chars?\):\*\*\s*(.+)", re.IGNORECASE)
    highlight_re = re.compile(r"\*\*Item Highlights \((\d+) chars?\):\*\*\s*(.+)", re.IGNORECASE)

    for block in blocks:
        lines = block.strip().splitlines()
        if not lines:
            continue
        asin = lines[0].strip()
        new_title = ""
        title_chars = 0
        highlights = ""
        highlight_chars = 0

        for line in lines[1:]:
            tm = title_re.search(line)
            hm = highlight_re.search(line)
            if tm:
                title_chars = int(tm.group(1))
                new_title = tm.group(2).strip()
            if hm:
                highlight_chars = int(hm.group(1))
                highlights = hm.group(2).strip()

        if new_title:
            results.append({
                "asin": asin,
                "new_title": new_title,
                "title_chars": len(new_title),   # recount ourselves
                "highlights": highlights,
                "highlight_chars": len(highlights),
            })

    return results


def process_with_claude(df: pd.DataFrame) -> list[dict]:
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    all_results = []

    for i in range(0, len(df), BATCH_SIZE):
        batch = df.iloc[i:i + BATCH_SIZE].to_dict("records")
        prompt = build_batch_prompt(batch)

        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=4096,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": prompt}],
        )

        batch_results = parse_claude_response(response.content[0].text)
        all_results.extend(batch_results)
        print(f"  Processed rows {i+1}–{min(i+BATCH_SIZE, len(df))} of {len(df)}")

        if i + BATCH_SIZE < len(df):
            time.sleep(0.5)  # gentle on rate limits

    return all_results


# ---------------------------------------------------------------------------
# Excel output builder
# ---------------------------------------------------------------------------

def build_excel(results: list[dict], output_path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Optimized Titles"

    thin = Side(border_style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", start_color="1F4E79")
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    alt_fill = PatternFill("solid", start_color="EBF3FB")
    data_font = Font(name="Arial", size=10)
    green_fill = PatternFill("solid", start_color="C6EFCE")
    red_fill = PatternFill("solid", start_color="FFC7CE")
    wrap = Alignment(wrap_text=True, vertical="top")
    center = Alignment(horizontal="center", vertical="top")

    headers = ["ASIN", "New Title", "Title Chars", "Item Highlights", "Highlights Chars"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    ws.row_dimensions[1].height = 30

    for row_idx, r in enumerate(results, 2):
        fill = alt_fill if row_idx % 2 == 0 else None
        values = [r["asin"], r["new_title"], r["title_chars"], r["highlights"], r["highlight_chars"]]
        for col, val in enumerate(values, 1):
            cell = ws.cell(row=row_idx, column=col, value=val)
            cell.font = data_font
            cell.border = border
            if fill:
                cell.fill = fill
            cell.alignment = center if col in (1, 3, 5) else wrap

        # Colour-code char count columns
        tc_cell = ws.cell(row=row_idx, column=3)
        hc_cell = ws.cell(row=row_idx, column=5)
        tc_cell.fill = green_fill if r["title_chars"] <= 75 else red_fill
        hc_cell.fill = green_fill if r["highlight_chars"] <= 125 else red_fill

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 13
    ws.column_dimensions["D"].width = 68
    ws.column_dimensions["E"].width = 17
    ws.freeze_panes = "A2"

    wb.save(output_path)


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.get("/")
def root():
    return {"status": "Tabulaic API running"}


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/process")
async def process_file(
    file: UploadFile = File(...),
    x_row_limit: int = Header(default=None),   # optional: paid tier passes higher limit
):
    """
    Accept a catalog file, optimize titles with Claude, return an Excel file.

    - Free tier: processes only the first 1 ASIN (x_row_limit not set or 0)
    - Paid tiers: pass X-Row-Limit header with the purchased limit (100 / 1000 / 999999)
    """
    if not ANTHROPIC_API_KEY:
        raise HTTPException(status_code=500, detail="API key not configured on server.")

    # Read upload
    file_bytes = await file.read()
    if len(file_bytes) == 0:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")

    # Parse
    try:
        df = parse_upload(file_bytes, file.filename)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))

    total_rows = len(df)

    # Apply row limit
    limit = x_row_limit if x_row_limit and x_row_limit > 0 else FREE_ROW_LIMIT
    df_to_process = df.head(limit)

    print(f"File: {file.filename} | Total rows: {total_rows} | Processing: {len(df_to_process)}")

    # Call Claude
    try:
        results = process_with_claude(df_to_process)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Claude processing error: {str(e)}")

    if not results:
        raise HTTPException(status_code=500, detail="No results returned from optimizer.")

    # Build Excel output
    output_filename = f"tabulaic_optimized_{uuid.uuid4().hex[:8]}.xlsx"
    output_path = f"/tmp/{output_filename}"
    build_excel(results, output_path)

    return FileResponse(
        path=output_path,
        filename="tabulaic_optimized.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.post("/preview")
async def preview_file(file: UploadFile = File(...)):
    """
    Returns row count and a 1-row sample — used by the frontend
    to show the user what was detected before they pay.
    """
    file_bytes = await file.read()
    try:
        df = parse_upload(file_bytes, file.filename)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))

    sample = df.head(1).to_dict("records")
    return JSONResponse({
        "total_rows": len(df),
        "sample_asin": sample[0]["asin"] if sample else "",
        "sample_title": sample[0]["title"] if sample else "",
    })
