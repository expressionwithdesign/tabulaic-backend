# Tabulaic Backend

FastAPI backend that accepts Amazon catalog files and returns optimized titles + Item Highlights.

## Files
- `main.py` — the full API
- `requirements.txt` — Python dependencies
- `railway.toml` — Railway deployment config

## Deploy to Railway (step by step)

1. Go to https://github.com and create a free account if you don't have one
2. Create a new repository called `tabulaic-backend` (private is fine)
3. Upload all three files: main.py, requirements.txt, railway.toml
4. Go to https://railway.app and sign up with your GitHub account
5. Click "New Project" → "Deploy from GitHub repo" → select tabulaic-backend
6. Once deployed, go to your project → "Variables" tab → add:
      ANTHROPIC_API_KEY = your-key-here
7. Railway will redeploy automatically. Copy your public URL — it looks like:
      https://tabulaic-backend-production.up.railway.app

## API Endpoints

### POST /process
Upload a file, get back an optimized Excel file.

Free tier (1 ASIN):
  curl -X POST https://your-url/process -F "file=@your_catalog.tsv"

Paid tier (e.g. 1000 rows):
  curl -X POST https://your-url/process \
       -H "X-Row-Limit: 1000" \
       -F "file=@your_catalog.tsv"

### POST /preview
Returns row count + sample before processing (used by the frontend).

  curl -X POST https://your-url/preview -F "file=@your_catalog.tsv"

### GET /health
  curl https://your-url/health
  → {"status": "ok"}

## Supported file types
- .tsv / .txt (Amazon Active Listings Report format)
- .csv
- .xlsx / .xls
