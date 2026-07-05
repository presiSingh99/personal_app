"""
FORGE — The Daily Ledger (Streamlit edition)
Run locally with:  streamlit run app.py
Every entry auto-saves to a local Excel workbook (default: forge_log.xlsx,
same folder as this script) on every change — no save button needed.
One row per calendar day — updating today's inputs updates today's row.
"""

import os
import datetime as dt

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Forge — Daily Ledger", page_icon="⚒", layout="centered")

DEFAULT_LOG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "forge_log.xlsx")

COLUMNS = [
    "Date", "Weekday", "Workout Split", "Workout Done",
    "Target 1 - Code/Review", "Target 2 - Deep Work", "Target 3 - Water/Nutrition",
    "Gratitude Reflected", "Project Hours", "Study/Review Hours", "Journal",
    "Daily XP %", "Streak",
]

SPLITS = {
    0: ("Active Recovery", "Incline walk 30-45min, mobility, structural rest"),
    1: ("Back Focus", "Lat Pulldown 4x10, Seated Cable Row 4x10, Chest Supported Row 3x10, Single Arm DB Row 3x10, Face Pulls 3x15, Shrugs 3x12"),
    2: ("Chest & Lateral Delts", "Incline DB Press 4x8-10, Flat Press 4x10, Incline Machine Press 3x10, Cable Flys 3x12, Pushups 2xfailure, Lateral Raises 4x15"),
    3: ("Legs Engine", "Squat/Hack Squat 4x8-10, RDL 4x10, Leg Press 3x12, Ham Curl 3x12, Leg Extension 3x12, Calf Raises 4x15, Abs Circuit 10min"),
    4: ("Aesthetic Symmetry", "Incline Press 3x10, Lat Pulldown/Pullups 3x10, Cable Row 3x12, Shoulder Press 3x10, Lateral Raises 5x15, Rear Delt Flys 3x15, Bicep Curl 3x12, Tricep Pushdown 3x12"),
    5: ("Arms & Shoulder Caps", "EZ Bar Curl 4x10, Hammer Curl 3x12, Preacher Curl 3x12, Rope Pushdown 4x12, Overhead Tricep Ext 3x12, Lateral Raises 5x15, Rear Delt Flys 3x15"),
    6: ("Active Recovery", "Incline walk 30-45min, mobility, structural rest"),
}
SPLIT_BY_PY_WEEKDAY = {0: SPLITS[1], 1: SPLITS[2], 2: SPLITS[3], 3: SPLITS[4], 4: SPLITS[5], 5: SPLITS[6], 6: SPLITS[0]}

INTEL = [
    ("We suffer more often in imagination than in reality.", "Seneca", "Idempotency",
     "Design every pipeline task so running it twice produces the same result as running it once. Retries become free instead of dangerous."),
    ("You have power over your mind, not outside events.", "Marcus Aurelius", "Parquet vs CSV",
     "Parquet is columnar and typed; analytical scans read only needed columns, cutting I/O 10-100x. Convert at the bronze layer."),
    ("The impediment to action advances action.", "Marcus Aurelius", "Vector Index Params",
     "In HNSW, M controls graph connectivity; efSearch controls query-time recall vs latency. Tune efSearch first since it needs no rebuild."),
    ("He who has a why to live can bear almost any how.", "Nietzsche", "Partitioning Strategy",
     "Partition by the column queried most, usually event date. Keep files 100MB-1GB; compact small files on a schedule."),
    ("We waste a great deal of the time we have.", "Seneca", "RAG Chunking",
     "Chunk by semantic boundaries, not fixed character counts, with 10-20% overlap. Retrieval quality is decided at ingestion time."),
    ("Discipline is the bridge between goals and accomplishment.", "Jim Rohn", "Slowly Changing Dimensions",
     "SCD Type 2 closes the old row and inserts a new one to preserve history. Always keep a current_flag and surrogate key."),
    ("Be one who does good, not one who argues about it.", "Marcus Aurelius", "Embedding Drift",
     "Changing embedding models invalidates old vectors. Re-embed behind a shadow index and cut over atomically — never mix versions."),
    ("Say to yourself what you would be, then do what you must.", "Epictetus", "Backfill Safely",
     "Backfills should be bounded, resumable, and rate-limited. Checkpoint progress; a backfill that can't pause is an outage waiting to happen."),
    ("The obstacle is the path.", "Zen proverb", "Prompt Caching",
     "Put static content (instructions, schemas, examples) first, variable content last, for cache-friendly, cheaper, faster prompts."),
    ("Real problems worry us less than our anxieties about them.", "Epictetus", "Schema Evolution",
     "Additive changes are safe; renames and type changes break things. Enforce a schema registry like an API version bump."),
    ("No man is free who is not master of himself.", "Epictetus", "Evaluation Before Scale",
     "Build a small golden dataset before tuning anything. A model change without an eval harness is a vibe, not an improvement."),
    ("Amateurs wait for inspiration; the rest of us get to work.", "Stephen King", "Late-Arriving Data",
     "Use watermarking with an allowed lateness window and reprocess recent partitions on a rolling basis."),
    ("Hard choices, easy life. Easy choices, hard life.", "Jerzy Gregorek", "Hybrid Search",
     "Combine BM25 keyword search with vector similarity via reciprocal rank fusion — reliably beats either retriever alone."),
    ("Disregard what lies beyond your control.", "Epictetus", "Orchestration Idioms",
     "DAG tasks should be small and parameterized by execution date, never by now() — or they can't be replayed."),
    ("Truth is high, but higher still is truthful living.", "Guru Nanak", "Data Contracts",
     "A data contract is schema, semantics, SLA, and ownership enforced in CI, not just documentation."),
    ("It does not matter how slowly you go, as long as you do not stop.", "Confucius", "Change Data Capture",
     "Log-based CDC streams row-level changes from the write-ahead log with near-zero load on the source."),
    ("Knowing others is intelligence; knowing yourself is wisdom.", "Lao Tzu", "Cross-Encoder Reranking",
     "Retrieve cheaply with bi-encoders, then rerank the top candidates with a cross-encoder for real accuracy gains."),
    ("I fear the man who has practiced one kick 10,000 times.", "Bruce Lee", "Medallion Architecture",
     "Bronze holds raw data, silver holds cleaned data, gold holds business marts — never skip a layer."),
    ("We are what we repeatedly do.", "Will Durant, on Aristotle", "Temperature & Top-p",
     "Run temperature near 0 for extraction tasks, higher for ideation. Tune one knob at a time."),
    ("No man ever steps in the same river twice.", "Heraclitus", "Streaming vs Batch",
     "Most 'real-time' asks are satisfied by 5-minute micro-batches at a fraction of streaming's operational cost."),
    ("There is more than one path to the top of the mountain.", "Miyamoto Musashi", "Structured Outputs",
     "Constrain LLM output with JSON schema or tool-call mode, and validate the response before trusting it downstream."),
    ("The superior man is modest in speech, exceeds in action.", "Confucius", "Dead Letter Queues",
     "Route bad records to a DLQ with payload, error, and timestamp instead of crashing or silently dropping data."),
    ("He who conquers himself is the mightiest warrior.", "Confucius", "Pipeline Observability",
     "Monitor freshness, volume, schema, and distribution on every table — alert on deltas, not absolutes."),
    ("Nature does not hurry, yet everything is accomplished.", "Lao Tzu", "Context Window Budgeting",
     "Reserve fixed token slots for system prompt, retrieved context, and history — evict oldest turns when over budget."),
    ("Prepare for war to preserve peace.", "Sun Tzu (paraphrase)", "Disaster Recovery for Data",
     "Define RPO and RTO, then rehearse a restore quarterly. Table-format time travel is not a substitute for DR."),
    ("The mind is everything; what you think, you become.", "attributed to the Buddha", "Embedding Quantization",
     "int8 or binary quantization shrinks vector storage 4-32x; rescore a shortlist with full precision for near-full quality."),
    ("Do every act as if it were your last.", "Marcus Aurelius", "Star Schema vs One Big Table",
     "Model dimensionally in silver, then denormalize into wide gold tables per use case for fast, simple queries."),
]


# ---------------------------------------------------------------------------
# Excel persistence
# ---------------------------------------------------------------------------
def load_log(path: str) -> pd.DataFrame:
    if os.path.exists(path):
        try:
            df = pd.read_excel(path, dtype={"Date": str})
            for col in COLUMNS:
                if col not in df.columns:
                    df[col] = None
            return df[COLUMNS]
        except Exception as e:
            st.error(f"Could not read existing log ({e}). A fresh one will be created on save.")
    return pd.DataFrame(columns=COLUMNS)


def compute_streak(df: pd.DataFrame, today_key: str, today_xp: int) -> tuple[int, int]:
    hist = df[df["Date"] != today_key].copy()
    hist = hist.sort_values("Date")
    xp_by_date = dict(zip(hist["Date"], hist["Daily XP %"]))
    xp_by_date[today_key] = today_xp

    dates_sorted = sorted(xp_by_date.keys())
    streaks, run = [], 0
    for d in dates_sorted:
        if xp_by_date.get(d, 0) >= 100:
            run += 1
        else:
            if run:
                streaks.append(run)
            run = 0
    if run:
        streaks.append(run)

    current = 0
    for d in reversed(dates_sorted):
        if xp_by_date.get(d, 0) >= 100:
            current += 1
        else:
            break
    best = max(streaks) if streaks else 0
    best = max(best, current)
    return current, best


def write_workbook(path: str, df: pd.DataFrame):
    wb = Workbook()
    ws = wb.active
    ws.title = "Forge Log"

    header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", start_color="20242E", end_color="20242E")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.append(COLUMNS)
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center

    for _, r in df.iterrows():
        ws.append([r[c] for c in COLUMNS])

    widths = [12, 11, 20, 13, 22, 20, 24, 18, 13, 18, 34, 11, 8]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    for row_cells in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row_cells:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.freeze_panes = "A2"
    wb.save(path)


def save_entry(path: str, row: dict):
    df = load_log(path)
    today_key = row["Date"]
    df = df[df["Date"] != today_key]
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    df = df.sort_values("Date").reset_index(drop=True)
    write_workbook(path, df)
    return df


def delete_today(path: str, today_key: str):
    df = load_log(path)
    df = df[df["Date"] != today_key].reset_index(drop=True)
    write_workbook(path, df)
    return df


# ---------------------------------------------------------------------------
# Style — dark minimalist, violet accent, activity ring, micro-interactions
# ---------------------------------------------------------------------------
def inject_style():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@600;700;800&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');

        :root{
            --ink:#0a0a0d; --card:#151519; --card-hi:#1c1c22;
            --text:#f5f5f7; --muted:#93939e; --line:rgba(255,255,255,.08);
            --accent:#7c6cf0; --accent-2:#c26cf0; --accent-soft:rgba(124,108,240,.14);
            --ok:#34d399; --ok-soft:rgba(52,211,153,.14);
        }

        html, body, [class*="css"] { font-family:'Inter',sans-serif; }
        .stApp { background:var(--ink); color:var(--text); }
        h1,h2,h3,h4 { font-family:'Manrope',sans-serif !important; letter-spacing:-.01em; }

        #MainMenu, footer, header[data-testid="stHeader"] { background:transparent; }
        .block-container { padding-top:1.4rem; padding-bottom:3rem; max-width:640px; }

        .hero { text-align:center; margin-bottom:1rem; animation:fadeDown .5s ease; }
        .hero .mark { font-size:2.2rem; font-weight:800; letter-spacing:-.02em; line-height:1; }
        .hero .mark span{ color:var(--accent); }
        .hero .tag { font-family:'JetBrains Mono',monospace; font-size:.65rem; letter-spacing:.3em;
                     color:var(--muted); text-transform:uppercase; margin-top:.3rem; }
        @keyframes fadeDown{ from{opacity:0; transform:translateY(-8px);} to{opacity:1; transform:none;} }

        div[data-testid="stVerticalBlockBorderWrapper"]{
            background:var(--card) !important; border:1px solid var(--line) !important;
            border-radius:16px !important; transition:transform .18s ease, border-color .18s ease, box-shadow .18s ease;
            animation:rise .45s ease backwards;
        }
        div[data-testid="stVerticalBlockBorderWrapper"]:hover{
            border-color:rgba(124,108,240,.35) !important; box-shadow:0 6px 24px rgba(0,0,0,.28);
        }
        @keyframes rise{ from{opacity:0; transform:translateY(10px);} to{opacity:1; transform:none;} }

        div[data-testid="stTabs"] button[role="tab"]{
            font-family:'Manrope',sans-serif; font-weight:700; font-size:.85rem;
            border-radius:999px !important; padding:.4rem 1rem !important; transition:all .18s ease; color:var(--muted);
        }
        div[data-testid="stTabs"] button[role="tab"]:hover{ color:var(--text); background:var(--card-hi); }
        div[data-testid="stTabs"] button[aria-selected="true"]{ color:#fff !important; background:var(--accent) !important; }
        div[data-testid="stTabs"] [data-baseweb="tab-highlight"]{ display:none; }
        div[data-testid="stTabs"] [data-baseweb="tab-border"]{ background:transparent; }

        label[data-testid="stCheckbox"]{ transition:transform .12s ease; }
        label[data-testid="stCheckbox"]:hover{ transform:translateX(2px); }

        div[data-testid="stButton"] button{
            border-radius:11px !important; font-weight:700 !important;
            transition:transform .1s ease, box-shadow .18s ease !important; border:1px solid var(--line) !important;
        }
        div[data-testid="stButton"] button:hover{ transform:translateY(-1px); box-shadow:0 6px 18px rgba(124,108,240,.18); }
        div[data-testid="stButton"] button:active{ transform:translateY(0) scale(.98); }

        div[data-testid="stMetric"]{
            background:var(--card-hi); border:1px solid var(--line); border-radius:12px;
            padding:.7rem .9rem; transition:transform .15s ease;
        }
        div[data-testid="stMetric"]:hover{ transform:translateY(-2px); }
        div[data-testid="stMetricLabel"]{ font-family:'JetBrains Mono',monospace; text-transform:uppercase;
            font-size:.64rem !important; letter-spacing:.14em; color:var(--muted) !important; }

        div[data-testid="stNumberInput"] input, div[data-testid="stTextArea"] textarea{
            border-radius:10px !important; background:var(--card-hi) !important; border:1px solid var(--line) !important;
        }

        section[data-testid="stSidebar"]{ background:var(--card) !important; border-right:1px solid var(--line); }
        hr{ border-color:var(--line) !important; margin:1.1rem 0 !important; }

        .quote-card{ font-family:'Manrope',sans-serif; font-size:1.05rem; font-weight:700; line-height:1.5; padding:.2rem 0 .1rem; }
        .quote-card .q::before{ content:"“"; color:var(--accent); font-size:1.4rem; }
        .quote-author{ font-family:'JetBrains Mono',monospace; font-size:.76rem; color:var(--muted); margin-top:.2rem; }

        .tier-chip{
            display:inline-block; font-family:'JetBrains Mono',monospace; font-size:.7rem; font-weight:700;
            letter-spacing:.1em; text-transform:uppercase; padding:.3rem .7rem; border-radius:999px;
            background:var(--accent-soft); color:var(--accent); border:1px solid rgba(124,108,240,.3);
        }
        .tier-chip.sage{ background:var(--ok-soft); color:var(--ok); border-color:rgba(52,211,153,.3); }

        .sync-badge{
            display:flex; align-items:center; gap:7px; font-family:'JetBrains Mono',monospace; font-size:.72rem;
            color:var(--ok); background:var(--ok-soft); border:1px solid rgba(52,211,153,.3);
            border-radius:10px; padding:.5rem .7rem; margin-top:.6rem;
        }
        .sync-dot{ width:7px; height:7px; border-radius:50%; background:var(--ok); flex:none;
            animation:pulse 1.6s ease-in-out infinite; }
        @keyframes pulse{ 0%,100%{opacity:1} 50%{opacity:.4} }
        </style>
        """,
        unsafe_allow_html=True,
    )


def ring_svg(pct: int, size: int = 88) -> str:
    r = 38
    circumference = 2 * 3.14159265 * r
    offset = circumference - (pct / 100) * circumference
    return f"""
    <div style="position:relative;width:{size}px;height:{size}px;flex:none;">
      <svg viewBox="0 0 88 88" width="{size}" height="{size}" style="transform:rotate(-90deg);">
        <defs>
          <linearGradient id="ringGrad" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" stop-color="#7c6cf0"/>
            <stop offset="100%" stop-color="#c26cf0"/>
          </linearGradient>
        </defs>
        <circle cx="44" cy="44" r="{r}" fill="none" stroke="#1c1c22" stroke-width="9"/>
        <circle cx="44" cy="44" r="{r}" fill="none" stroke="url(#ringGrad)" stroke-width="9"
                stroke-linecap="round" stroke-dasharray="{circumference:.2f}" stroke-dashoffset="{offset:.2f}"
                style="transition:stroke-dashoffset .6s cubic-bezier(.22,1,.36,1);"/>
      </svg>
      <div style="position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;">
        <div style="font-family:'Manrope',sans-serif;font-weight:800;font-size:19px;color:#f5f5f7;">{pct}%</div>
        <div style="font-size:9px;color:#5c5c66;letter-spacing:.08em;text-transform:uppercase;margin-top:-2px;">XP</div>
      </div>
    </div>
    """


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------
inject_style()

st.markdown(
    """
    <div class="hero">
        <div class="mark">F<span>O</span>RGE</div>
        <div class="tag">The Daily Ledger</div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown("#### ⚙ Settings")
    log_path = st.text_input("Excel log file path", value=DEFAULT_LOG_PATH, label_visibility="collapsed")
    st.caption("Every change auto-saves here — no button needed.")
    if os.path.exists(log_path):
        st.success(f"Log found · {os.path.getsize(log_path)/1024:.1f} KB", icon="📗")
    else:
        st.info("No log yet — created on your first change.", icon="📄")

today = dt.date.today()
today_key = today.isoformat()
py_wd = today.weekday()
split_title, split_detail = SPLIT_BY_PY_WEEKDAY[py_wd]
day_of_year = today.timetuple().tm_yday
quote, author, lesson_title, lesson_body = INTEL[day_of_year % len(INTEL)]

existing_df = load_log(log_path)

if st.session_state.get("_clear_today"):
    st.session_state["_clear_today"] = False
    delete_today(log_path, today_key)
    st.session_state["reset_gen"] = st.session_state.get("reset_gen", 0) + 1
    existing_df = load_log(log_path)

reset_gen = st.session_state.get("reset_gen", 0)
wkey = lambda name: f"{name}_{today_key}_{reset_gen}"

existing_today = existing_df[existing_df["Date"] == today_key]
prefill = existing_today.iloc[0].to_dict() if not existing_today.empty else {}

st.caption(f"📅 {today.strftime('%A, %B %d, %Y')}")

tab_today, tab_journal, tab_intel, tab_stats = st.tabs(["🏠 Today", "✍ Journal", "🧠 Intel", "📊 Stats"])

with tab_today:
    with st.container(border=True):
        st.markdown("##### ⚒ Today's Iron")
        st.markdown(f"**{split_title}**")
        st.caption(split_detail)
        workout_done = st.checkbox("Mark today's session complete", value=bool(prefill.get("Workout Done", False)), key=wkey("workout_done"))

    with st.container(border=True):
        st.markdown("##### 🎯 Rule of Three")
        t1 = st.checkbox("Code / review pipeline architecture layouts", value=bool(prefill.get("Target 1 - Code/Review", False)), key=wkey("t1"))
        t2 = st.checkbox("Deep work window (no distractions)", value=bool(prefill.get("Target 2 - Deep Work", False)), key=wkey("t2"))
        t3 = st.checkbox("Drink 4L of water & clean nutrition", value=bool(prefill.get("Target 3 - Water/Nutrition", False)), key=wkey("t3"))

    with st.container(border=True):
        st.markdown("##### ✦ Nightly Reflection")
        gratitude = st.checkbox("I've thought of 3 things I'm grateful for", value=bool(prefill.get("Gratitude Reflected", False)), key=wkey("gratitude"))

with tab_journal:
    with st.container(border=True):
        st.markdown("##### ⏱ Time Ledger")
        col1, col2 = st.columns(2)
        with col1:
            project_hours = st.number_input(
                "Project hours", min_value=0.0, max_value=24.0, step=0.25,
                value=float(prefill.get("Project Hours", 0.0) or 0.0), key=wkey("project_hours"),
            )
        with col2:
            study_hours = st.number_input(
                "Study / review hours", min_value=0.0, max_value=24.0, step=0.25,
                value=float(prefill.get("Study/Review Hours", 0.0) or 0.0), key=wkey("study_hours"),
            )

    with st.container(border=True):
        st.markdown("##### 📝 What did you do today?")
        journal = st.text_area(
            "Journal", value=str(prefill.get("Journal", "") or ""), height=110,
            placeholder="Shipped the ingestion retry logic, reviewed two PRs, read a chapter on vector indexing…",
            label_visibility="collapsed", max_chars=600, key=wkey("journal"),
        )
        st.caption(f"{len(journal)} / 600")

with tab_intel:
    with st.container(border=True):
        st.markdown("##### 🧠 Mind & Machine")
        st.markdown(f'<div class="quote-card"><span class="q">{quote}</span>”</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="quote-author">— {author}</div>', unsafe_allow_html=True)
        st.markdown("---")
        st.markdown(f"**{lesson_title}**")
        st.caption(lesson_body)

with tab_stats:
    df_display = load_log(log_path)
    if df_display.empty:
        st.info("No entries yet — your first save will appear here automatically.")
    else:
        with st.container(border=True):
            total_project = pd.to_numeric(df_display["Project Hours"], errors="coerce").sum()
            total_study = pd.to_numeric(df_display["Study/Review Hours"], errors="coerce").sum()
            cc1, cc2, cc3 = st.columns(3)
            cc1.metric("Days logged", len(df_display))
            cc2.metric("Project hrs", f"{total_project:.1f}")
            cc3.metric("Study hrs", f"{total_study:.1f}")
        st.dataframe(df_display.sort_values("Date", ascending=False), width='stretch', hide_index=True)

    with st.container(border=True):
        st.markdown("##### 🗓 Day Cycle")
        st.caption(
            "Each calendar day gets its own row automatically — tomorrow starts blank "
            "with no action needed. Use this only if you want to wipe *today's* entry and start over."
        )
        if st.button("🗑 Clear today's entry", width='stretch'):
            st.session_state["_clear_today"] = True
            st.rerun()

# ---- XP + streak ----
tokens = [workout_done, t1, t2, t3, gratitude]
xp = sum(tokens) * 20
tier_label = "Sovereign Sage" if xp >= 100 else ("Warrior-Scholar" if xp >= 40 else "The Seeker")
tier_class = "sage" if xp >= 100 else ""

current_streak, best_streak = compute_streak(existing_df, today_key, xp)

with st.sidebar:
    st.markdown("---")
    st.markdown(ring_svg(xp), unsafe_allow_html=True)
    st.markdown(f'<div style="text-align:center;margin-top:8px;"><span class="tier-chip {tier_class}">{tier_label}</span></div>', unsafe_allow_html=True)
    s1, s2 = st.columns(2)
    s1.metric("Streak", f"{current_streak}d")
    s2.metric("Best", f"{best_streak}d")

# ---- Auto-save on every rerun (no button — feels like a live database) ----
row = {
    "Date": today_key,
    "Weekday": today.strftime("%A"),
    "Workout Split": split_title,
    "Workout Done": workout_done,
    "Target 1 - Code/Review": t1,
    "Target 2 - Deep Work": t2,
    "Target 3 - Water/Nutrition": t3,
    "Gratitude Reflected": gratitude,
    "Project Hours": project_hours,
    "Study/Review Hours": study_hours,
    "Journal": journal,
    "Daily XP %": xp,
    "Streak": current_streak,
}

try:
    save_entry(log_path, row)
    with st.sidebar:
        st.markdown(
            f'<div class="sync-badge"><span class="sync-dot"></span>Synced to Excel · {dt.datetime.now().strftime("%H:%M:%S")}</div>',
            unsafe_allow_html=True,
        )
    if xp >= 100 and not st.session_state.get("_celebrated_today") == today_key:
        st.session_state["_celebrated_today"] = today_key
        st.balloons()
        st.toast("Perfect day — 100% XP!", icon="🔥")
except Exception as e:
    with st.sidebar:
        st.error(f"Auto-save failed: {e}")

if st.session_state.get("_force_reset"):
    st.session_state["_force_reset"] = False
    st.toast("Day committed. Streak updates on your next perfect day.", icon="✅")
