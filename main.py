import os
import re
import time
import uuid
import tempfile
import pandas as pd
import anthropic
import stripe

from fastapi import FastAPI, File, UploadFile, HTTPException, Header, Request
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------

app = FastAPI(title="Tabulaic API", version="2.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

FREE_ROW_LIMIT   = 1
BATCH_SIZE       = 30

ANTHROPIC_API_KEY      = os.environ.get("ANTHROPIC_API_KEY")
STRIPE_SECRET_KEY      = os.environ.get("STRIPE_SECRET_KEY")
STRIPE_WEBHOOK_SECRET  = os.environ.get("STRIPE_WEBHOOK_SECRET")

stripe.api_key = STRIPE_SECRET_KEY

# Tier config — price_id → row limit
TIER_LIMITS = {
    "price_1Tnlr9RoZgDHWVTnMoItFrlt": 100,      # Starter Fix   $27
    "price_1TnlsGRoZgDHWVTnYQBwcujz": 1000,     # Full Catalog  $97
    "price_1TnltNRoZgDHWVTnPtrSfee2": 999999,   # Unlimited     $197
}

# In-memory token store: token → row_limit
# (swap for Redis/Postgres in production for persistence across restarts)
TOKEN_STORE: dict[str, int] = {}

# ---------------------------------------------------------------------------
# System prompt
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
# File parsing
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
    results = []
    blocks = [b.strip() for b in text.split("---") if b.strip()]
    title_re = re.compile(r"\*\*New Title \((\d+) chars?\):\*\*\s*(.+)", re.IGNORECASE)
    highlight_re = re.compile(r"\*\*Item Highlights \((\d+) chars?\):\*\*\s*(.+)", re.IGNORECASE)

    for block in blocks:
        lines = block.strip().splitlines()
        if not lines:
            continue
        asin = lines[0].strip()
        new_title = highlights = ""
        for line in lines[1:]:
            tm = title_re.search(line)
            hm = highlight_re.search(line)
            if tm:
                new_title = tm.group(2).strip()
            if hm:
                highlights = hm.group(2).strip()
        if new_title:
            results.append({
                "asin": asin,
                "new_title": new_title,
                "title_chars": len(new_title),
                "highlights": highlights,
                "highlight_chars": len(highlights),
            })
    return results


def process_with_claude(df: pd.DataFrame) -> list[dict]:
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    all_results = []
    for i in range(0, len(df), BATCH_SIZE):
        batch = df.iloc[i:i + BATCH_SIZE].to_dict("records")
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=4096,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": build_batch_prompt(batch)}],
        )
        all_results.extend(parse_claude_response(response.content[0].text))
        print(f"  Processed rows {i+1}–{min(i+BATCH_SIZE, len(df))} of {len(df)}")
        if i + BATCH_SIZE < len(df):
            time.sleep(0.5)
    return all_results

# ---------------------------------------------------------------------------
# Excel builder
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
        for col, val in enumerate([r["asin"], r["new_title"], r["title_chars"], r["highlights"], r["highlight_chars"]], 1):
            cell = ws.cell(row=row_idx, column=col, value=val)
            cell.font = data_font
            cell.border = border
            if fill:
                cell.fill = fill
            cell.alignment = center if col in (1, 3, 5) else wrap
        ws.cell(row=row_idx, column=3).fill = green_fill if r["title_chars"] <= 75 else red_fill
        ws.cell(row=row_idx, column=5).fill = green_fill if r["highlight_chars"] <= 125 else red_fill

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 13
    ws.column_dimensions["D"].width = 68
    ws.column_dimensions["E"].width = 17
    ws.freeze_panes = "A2"
    wb.save(output_path)

# ---------------------------------------------------------------------------
# Routes — health & root
# ---------------------------------------------------------------------------

@app.get("/")
def root():
    return {"status": "Tabulaic API running"}

@app.get("/health")
def health():
    return {"status": "ok"}

# ---------------------------------------------------------------------------
# Stripe: create checkout session
# ---------------------------------------------------------------------------

@app.post("/create-checkout-session")
async def create_checkout_session(request: Request):
    body = await request.json()
    price_id = body.get("price_id")

    if price_id not in TIER_LIMITS:
        raise HTTPException(status_code=400, detail="Invalid price ID.")

    try:
        session = stripe.checkout.Session.create(
            payment_method_types=["card"],
            line_items=[{"price": price_id, "quantity": 1}],
            mode="payment",
            success_url="https://tabulaic.com/success?token={CHECKOUT_SESSION_ID}",
            cancel_url="https://tabulaic.com/#pricing",
            metadata={"price_id": price_id},
        )
        return JSONResponse({"checkout_url": session.url})
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# ---------------------------------------------------------------------------
# Stripe: webhook — fires after payment, issues token
# ---------------------------------------------------------------------------

@app.post("/webhook")
async def stripe_webhook(request: Request):
    payload = await request.body()
    sig_header = request.headers.get("stripe-signature")

    try:
        event = stripe.Webhook.construct_event(payload, sig_header, STRIPE_WEBHOOK_SECRET)
    except stripe.error.SignatureVerificationError:
        raise HTTPException(status_code=400, detail="Invalid webhook signature.")

    if event["type"] == "checkout.session.completed":
        session = event["data"]["object"]
        price_id = session["metadata"].get("price_id")
        row_limit = TIER_LIMITS.get(price_id, FREE_ROW_LIMIT)

        # Use Stripe session ID as the token
        token = session["id"]
        TOKEN_STORE[token] = row_limit
        print(f"Token issued: {token} → {row_limit} rows")

    return JSONResponse({"status": "ok"})

# ---------------------------------------------------------------------------
# Process: validate token, run Claude, return Excel
# ---------------------------------------------------------------------------

@app.post("/process")
async def process_file(
    file: UploadFile = File(...),
    x_token: str = Header(default=None),
):
    if not ANTHROPIC_API_KEY:
        raise HTTPException(status_code=500, detail="API key not configured.")

    # Determine row limit from token
    if x_token and x_token in TOKEN_STORE:
        row_limit = TOKEN_STORE[x_token]
        # Consume token after use (one-time)
        del TOKEN_STORE[x_token]
    else:
        row_limit = FREE_ROW_LIMIT  # free: 1 ASIN

    file_bytes = await file.read()
    if not file_bytes:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")

    try:
        df = parse_upload(file_bytes, file.filename)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))

    df_to_process = df.head(row_limit)
    print(f"File: {file.filename} | Total: {len(df)} | Processing: {len(df_to_process)} | Limit: {row_limit}")

    try:
        results = process_with_claude(df_to_process)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Claude error: {str(e)}")

    if not results:
        raise HTTPException(status_code=500, detail="No results returned.")

    output_path = f"/tmp/tabulaic_{uuid.uuid4().hex[:8]}.xlsx"
    build_excel(results, output_path)

    return FileResponse(
        path=output_path,
        filename="tabulaic_optimized.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ---------------------------------------------------------------------------
# Preview: row count + sample (no token needed)
# ---------------------------------------------------------------------------

@app.post("/preview")
async def preview_file(file: UploadFile = File(...)):
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
