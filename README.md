# Forge — The Daily Ledger

A personal daily-discipline tracker built with Streamlit. Logs your workout
split, accountability targets, gratitude check-in, and project/study hours
straight into a local Excel file — one row per day, no cloud, no account.

## Features

- Auto-selected daily workout split (5-day push/pull/legs style rotation + recovery days)
- "Rule of Three" daily accountability checkboxes
- Simple gratitude check-in
- Project hours and study/review hours logging
- Rotating daily quote + technical lesson (27-entry pool)
- Live XP, tier, and streak tracking
- Every save writes to a local `.xlsx` file — your data never leaves your machine

## Setup

```bash
git clone https://github.com/<your-username>/<your-repo>.git
cd <your-repo>
pip install -r requirements.txt
streamlit run app.py
```

The app opens at `http://localhost:8501`.

## Where your data goes

On first save, the app creates `forge_log.xlsx` in the same folder as
`app.py` (the path is editable in the sidebar). This file is intentionally
excluded from git via `.gitignore` — it's personal data, not part of the
app itself. Each day gets one row; saving again on the same day updates
that row instead of creating a duplicate.

## Using it from your phone

Run the app on your computer with:

```bash
streamlit run app.py --server.address 0.0.0.0
```

Then, on your phone (same WiFi network), open
`http://<your-computer's-local-IP>:8501` and add it to your home screen.
Your computer needs to stay on and awake for this to work, since the
Excel file lives on it — that's the trade-off for keeping your data local
instead of in the cloud.

## Project structure

```
.
├── app.py              # the whole app
├── requirements.txt    # streamlit, pandas, openpyxl
├── .gitignore          # excludes your personal forge_log.xlsx
└── README.md
```

## Requirements

- Python 3.9+
- streamlit, pandas, openpyxl (see requirements.txt)
