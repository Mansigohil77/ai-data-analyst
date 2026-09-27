
None selected 

Skip to content
Using Gmail with screen readers
in:sent 
1 of 121
Ai data analyst

Mansi Gohil <mansigohil2005@gmail.com>
Attachments
26 Sept 2026, 12:12 (1 day ago)
to Archita

Linkdin description

📊 AI Data Analyst — From Raw Data to Verified Findings

I built a local-first AI Data Analyst that turns raw business data into analysis, evidence-backed findings, verification, and professional reports — without relying on paid AI APIs.

The goal was not to build another chatbot that simply answers questions about a dataset.

Instead, I designed the workflow around:

Data → Deterministic Analysis → Evidence → AI Explanation → Verification → Report

🔍 What it can do
• Import Excel, CSV, JSON, Parquet and other data sources
• Profile datasets with rows, columns, schema and data quality
• Detect missing values, duplicates and outliers
• Clean and transform data through Transform Studio
• Query data using SQL
• Perform Python-based analysis
• Automatically generate dashboards and KPIs
• Explore statistics, distributions and correlations
• Forecast trends and run What-If scenarios
• Ask business questions using local AI
• Review supporting evidence behind findings
• Verify and challenge analytical conclusions
• Export results to Excel, PowerPoint and PDF

🤖 Local AI
The AI layer runs locally using Ollama + Phi-3, keeping the application designed around a local-first workflow rather than sending business data to a paid external AI API.

🧠 The part I focused on most
AI should not simply produce an answer and ask us to trust it.

The application first works with the data and generates deterministic evidence. The local AI then helps explain that evidence, followed by verification and challenge steps.

Import → Clean → Analyze → Visualize → AI → Verify → Forecast → Report

Built with Python, Streamlit, Pandas, DuckDB, statistical analysis tools, local Ollama, and reporting/export libraries.

🔗 GitHub: [add your repository link]

#AI #DataAnalytics #ArtificialIntelligence #Python #Streamlit #Ollama #DataScience #BusinessIntelligence #LocalAI #MachineLearning #GitHub #BuildInPublic


Guithub description


📊 AI Data Analyst
From raw business data to verified findings.
AI Data Analyst is a local-first analytics application designed to take a dataset from import → cleaning → analysis → visualization → AI explanation → verification → reporting.

It is designed around a simple principle:

AI should explain the evidence — not replace the analysis.

Instead of sending a dataset directly to an AI model and asking it to guess the answer, the application combines deterministic data analysis with local AI.

                 AI DATA ANALYST
                       │
                 Import Data
                       ↓
              Profile & Data Quality
                       ↓
              Clean & Transform
                       ↓
             ┌─────────┴─────────┐
             ↓                   ↓
          SQL Analysis      Python Analysis
             └─────────┬─────────┘
                       ↓
                Auto Dashboard
                       ↓
             Statistics & Forecast
                       ↓
              Deterministic Evidence
                       ↓
                 Local AI Layer
                       ↓
             Evidence & Verification
                       ↓
                Challenge Findings
                       ↓
              Report & Export
🚀 What does it do?
AI Data Analyst provides an end-to-end workflow for working with business datasets.

📂 Data Sources
Excel — XLS/XLSX/XLSM

CSV

JSON

Parquet

TSV/TXT

SQL/database workflows

Static inspection of Python and Java files

🧹 Data Quality & Cleaning
Missing-value analysis

Duplicate detection

Outlier analysis

Schema inspection

Data-quality scoring

Cleaning operations

Transform Studio

Undo/redo

Audit history

Clean-data export

🔎 Analysis
SQL Analyst

Python Analysis

Automatic KPIs

Dashboard generation

Exploratory analysis

Statistics

Correlation analysis

Distributions

Forecasting

What-If analysis

🤖 Local AI Analyst
The AI layer uses Ollama with a local model such as Phi-3.

The application can route questions through the analytical workflow and provide AI-assisted explanations while keeping the AI layer local.

🔬 Evidence & Verification
One of the core design principles of this project is evidence-first analysis.

A finding should be traceable back to the underlying data.

The workflow therefore separates:

Calculation → Evidence → AI Explanation → Verification

The application also provides a challenge step to examine alternative explanations instead of treating the first generated answer as automatically correct.

📊 Dashboards & Reporting
The analysis can be converted into professional deliverables:

Excel workbooks

PDF reports

PowerPoint presentations

Charts

KPI summaries

Findings

Evidence

Verification results

Forecast results

Audit/provenance information

🛠️ Technology
Python

Streamlit

Pandas

NumPy

DuckDB

Matplotlib

Statsmodels

OpenPyXL

ReportLab

Python-PPTX

Ollama

Phi-3

🔐 Local-First Design
The application is designed to run locally.

The AI layer uses a local Ollama endpoint rather than requiring a paid cloud AI API.

User Data
   ↓
Local Application
   ↓
Deterministic Analysis
   ↓
Evidence
   ↓
Local Ollama / Phi-3
   ↓
Verified Findings
   ↓
Reports
This architecture is particularly useful when working with business or sensitive datasets where keeping the analytical workflow local is important.

🔄 End-to-End Workflow
Import → Clean → Analyze → Visualize → AI → Verify → Forecast → Report

The objective is not simply to generate an AI answer.

The objective is to produce an answer that can be investigated, supported by evidence, reviewed, and exported as a usable business deliverable.

 2 attachment
  •  Scanned by Gmail
from __future__ import annotations

import ast
import io
import json
import math
import os
import re
import hashlib
import shutil
import sqlite3
import time
from pathlib import Path
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st

# ============================================================
# OPTIONAL DEPENDENCIES
# ============================================================

try:
    import requests
except Exception:
    requests = None

try:
    import plotly.express as px
    import plotly.graph_objects as go
except Exception:
    px = None
    go = None

try:
    import duckdb
except Exception:
    duckdb = None

try:
    from statsmodels.tsa.holtwinters import ExponentialSmoothing
except Exception:
    ExponentialSmoothing = None

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
except Exception:
    plt = None

try:
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table as XLTable
    from openpyxl.worksheet.table import TableStyleInfo
    OPENPYXL_AVAILABLE = True
except Exception:
    OPENPYXL_AVAILABLE = False

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.enums import TA_CENTER
    from reportlab.platypus import (
        SimpleDocTemplate,
        Paragraph,
        Spacer,
        Table,
        TableStyle,
        PageBreak,
        Image as RLImage,
    )
    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.dml.color import RGBColor
    PPTX_AVAILABLE = True
except Exception:
    PPTX_AVAILABLE = False


# ============================================================
# APPLICATION CONFIGURATION
# ============================================================

APP_NAME = "AI Data Analyst"
APP_VERSION = "8.0.0"

OLLAMA_URL = os.getenv(
    "OLLAMA_URL",
    "http://127.0.0.1:11434"
).rstrip("/")

OLLAMA_MODEL = os.getenv(
    "OLLAMA_MODEL",
    "phi3:latest"
)

OLLAMA_TAGS_URL = f"{OLLAMA_URL}/api/tags"
OLLAMA_GENERATE_URL = f"{OLLAMA_URL}/api/generate"

PROJECT_DIR = Path("projects")
PROJECT_DIR.mkdir(parents=True, exist_ok=True)

MAX_UNDO = 20
MAX_CHAT_MESSAGES = 30
MAX_SQL_ROWS = 5000

AI_NUM_CTX = 2048
AI_NUM_PREDICT = 350
AI_TEMPERATURE = 0.2

OLLAMA_CONNECT_TIMEOUT = 8
OLLAMA_READ_TIMEOUT = 240

MAX_UPLOAD_MB = 200


# ============================================================
# STREAMLIT
# ============================================================

st.set_page_config(
    page_title=APP_NAME,
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# DARK PROFESSIONAL UI
# ============================================================

st.markdown(
    """
<style>

:root {
    color-scheme: dark;
    --bg:#07111f;
    --panel:#0d1b2d;
    --panel2:#10243a;
    --border:#263b52;
    --text:#eef4fa;
    --muted:#91a5b8;
    --accent:#55c7ff;
    --good:#48d597;
    --warn:#ffca63;
    --bad:#ff7272;
}

html,
body,
[data-testid="stAppViewContainer"] {
    background:var(--bg) !important;
    color:var(--text) !important;
}

[data-testid="stHeader"] {
    background:rgba(7,17,31,.95) !important;
}

.block-container {
    max-width:1550px;
    padding-top:1.2rem;
    padding-bottom:4rem;
}

section[data-testid="stSidebar"] {
    background:#081522 !important;
    border-right:1px solid var(--border);
}

.hero {
    padding:26px 30px;
    border:1px solid var(--border);
    border-radius:20px;
    background:linear-gradient(135deg,#10263d,#0a1727);
    margin-bottom:20px;
}

.hero h1 {
    margin:0;
    font-size:2.25rem;
    font-weight:800;
    letter-spacing:-.04em;
}

.hero p {
    color:var(--muted);
    margin:.5rem 0 0;
}

.kpi {
    background:linear-gradient(180deg,#10243a,#0c1a2a);
    border:1px solid var(--border);
    border-radius:16px;
    padding:17px;
    min-height:110px;
}

.kpi .label {
    color:var(--muted);
    font-size:.78rem;
}

.kpi .value {
    color:#fff;
    font-size:1.55rem;
    font-weight:800;
    margin-top:5px;
}

.kpi .sub {
    color:var(--muted);
    font-size:.73rem;
    margin-top:4px;
}

.badge {
    display:inline-block;
    padding:5px 10px;
    border-radius:999px;
    border:1px solid var(--border);
    font-size:.72rem;
    margin-right:5px;
    background:#0b1928;
}

.good { color:var(--good); }
.warn { color:var(--warn); }
.bad { color:var(--bad); }

.finding {
    border-left:4px solid var(--accent);
    padding:13px 15px;
    background:#0b1928;
    border-radius:10px;
    margin:9px 0;
}

.finding strong {
    color:#fff;
}

.stepbar {
    display:flex;
    gap:7px;
    flex-wrap:wrap;
    margin:8px 0 18px;
}

.step {
    border:1px solid var(--border);
    border-radius:999px;
    padding:6px 11px;
    color:var(--muted);
    font-size:.7rem;
}

.step.active {
    color:#fff;
    border-color:var(--accent);
    background:#0d2940;
}

.small {
    color:var(--muted);
    font-size:.78rem;
}

.report-card {
    border:1px solid var(--border);
    border-radius:15px;
    padding:15px;
    background:#0c1b2c;
    margin-bottom:12px;
}

[data-testid="stMetric"] {
    background:var(--panel) !important;
    border:1px solid var(--border) !important;
    border-radius:14px !important;
}

button:hover,
.stButton button:hover,
.stDownloadButton button:hover {
    background:#1b2432 !important;
    color:#fff !important;
    border-color:#3b82f6 !important;
}

button:focus,
button:active,
.stButton button:focus,
.stButton button:active {
    background:#202938 !important;
    color:#fff !important;
}

input,
textarea {
    background:#0f1520 !important;
    color:#eef2f7 !important;
    border-color:#293446 !important;
}

input:focus,
textarea:focus {
    background:#111827 !important;
    color:#fff !important;
    border-color:#4f7cff !important;
}

[data-baseweb="select"] > div {
    background:#0f1520 !important;
    color:#eef2f7 !important;
    border-color:#293446 !important;
}

[data-baseweb="popover"] *,
[data-baseweb="menu"] *,
[role="listbox"] *,
[role="option"] * {
    background:#0f1520 !important;
    color:#eef2f7 !important;
}

[role="option"]:hover,
[role="option"][aria-selected="true"] {
    background:#1b2432 !important;
    color:#fff !important;
}

[data-baseweb="tooltip"],
[role="tooltip"] {
    background:#0b111a !important;
    color:#f8fafc !important;
    border:1px solid #293446 !important;
}

[data-testid="stChatInput"] textarea {
    background:#0f1520 !important;
    color:#fff !important;
}

hr {
    border-color:var(--border);
}

</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# SESSION STATE
# ============================================================

DEFAULT_STATE = {
    "datasets": {},
    "active_dataset_key": None,
    "source_name": "",
    "undo_stack": [],
    "redo_stack": [],
    "audit_log": [],
    "cleaning_history": [],
    "chat_history": [],
    "sql_history": [],
    "python_history": [],
    "python_results": [],
    "last_forecast": None,
    "last_forecast_signature": None,
    "forecast_config": {},
    "dashboard_config": {},
    "dashboard_saved": False,
    "findings": [],
    "selected_finding": None,
    "verification_results": [],
    "challenge_results": [],
    "evidence_ledger": [],
    "analysis_notes": [],
    "data_fingerprint": None,
    "project_name": "Untitled Analysis",
    "project_created": None,
    "ollama_models": [],
    "ollama_status": False,
    "selected_model": OLLAMA_MODEL,
    "page": "Executive Overview",
    "demo_loaded": False,
    "report_config": {},
}

for key, value in DEFAULT_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# GENERAL HELPERS
# ============================================================

def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def safe_filename(name: str) -> str:
    name = re.sub(
        r"[^A-Za-z0-9._-]+",
        "_",
        str(name)
    ).strip("._")

    return name or "dataset"


def current_df() -> Optional[pd.DataFrame]:
    key = st.session_state.active_dataset_key

    if not key:
        return None

    return st.session_state.datasets.get(key)


def current_dataset_name() -> str:
    return (
        st.session_state.source_name
        or st.session_state.active_dataset_key
        or "No dataset"
    )


def numeric_columns(df: pd.DataFrame) -> List[str]:
    return list(
        df.select_dtypes(include=[np.number]).columns
    )


def categorical_columns(df: pd.DataFrame) -> List[str]:
    return [
        c
        for c in df.columns
        if not pd.api.types.is_numeric_dtype(df[c])
        and not pd.api.types.is_datetime64_any_dtype(df[c])
    ]


def detect_datetime_candidates(df: pd.DataFrame) -> List[str]:
    result = []

    pattern = re.compile(
        r"date|time|timestamp|month|year|day|created|updated|at$",
        re.I,
    )

    for column in df.columns:

        series = df[column]

        if pd.api.types.is_datetime64_any_dtype(series):
            result.append(column)
            continue

        if pattern.search(str(column)):

            parsed = pd.to_datetime(
                series,
                errors="coerce"
            )

            if parsed.notna().mean() >= .65:
                result.append(column)

    return list(dict.fromkeys(result))


def is_numeric_series(series: pd.Series) -> bool:
    return pd.api.types.is_numeric_dtype(series)


def arrow_safe_df(df: pd.DataFrame) -> pd.DataFrame:

    out = df.copy()

    for column in out.columns:

        if pd.api.types.is_datetime64_any_dtype(
            out[column]
        ):
            out[column] = out[column].astype(str)

        elif pd.api.types.is_object_dtype(
            out[column]
        ):
            out[column] = out[column].map(
                lambda x:
                x
                if isinstance(
                    x,
                    (
                        str,
                        int,
                        float,
                        bool,
                        type(None),
                    ),
                )
                else str(x)
            )

    return out


def dataframe_signature(
    df: pd.DataFrame
) -> str:

    try:

        sample = df.head(250).copy()

        raw = pd.util.hash_pandas_object(
            sample,
            index=True
        ).values.tobytes()

        metadata = json.dumps(
            {
                "shape": list(df.shape),
                "columns": list(
                    map(str, df.columns)
                ),
            },
            default=str,
        ).encode()

        return hashlib.sha256(
            raw + metadata
        ).hexdigest()

    except Exception:

        return hashlib.sha256(
            repr(df.shape).encode()
        ).hexdigest()


def format_value(value: Any) -> str:

    if value is None:
        return "—"

    try:

        if isinstance(value, (float, np.floating)):
            if math.isnan(float(value)):
                return "—"

        if isinstance(value, (int, float, np.number)):

            value = float(value)

            absolute = abs(value)

            if absolute >= 1_000_000_000:
                return f"{value / 1_000_000_000:.2f}B"

            if absolute >= 1_000_000:
                return f"{value / 1_000_000:.2f}M"

            if absolute >= 1_000:
                return f"{value / 1_000:.2f}K"

            return f"{value:,.2f}"

    except Exception:
        pass

    return str(value)


def add_audit(
    action: str,
    details: str = ""
):

    st.session_state.audit_log.append(
        {
            "time": now_text(),
            "action": action,
            "details": details,
        }
    )

    st.session_state.audit_log = (
        st.session_state.audit_log[-300:]
    )


def add_cleaning(
    action: str,
    before_rows: int,
    after_rows: int,
    details: str = ""
):

    st.session_state.cleaning_history.append(
        {
            "time": now_text(),
            "action": action,
            "before_rows": before_rows,
            "after_rows": after_rows,
            "details": details,
        }
    )

    st.session_state.cleaning_history = (
        st.session_state.cleaning_history[-200:]
    )


# ============================================================
# DATASET MANAGEMENT
# ============================================================

def unique_dataset_key(name: str) -> str:

    base = safe_filename(
        Path(str(name)).stem
        or "dataset"
    )

    key = base
    counter = 2

    while key in st.session_state.datasets:

        key = f"{base}_{counter}"
        counter += 1

    return key


def register_dataset(
    df: pd.DataFrame,
    name: str,
    source: str = "import",
    make_active: bool = True,
):

    df = df.copy()

    key = unique_dataset_key(name)

    st.session_state.datasets[key] = df

    if make_active:

        st.session_state.active_dataset_key = key
        st.session_state.source_name = str(name)

        st.session_state.undo_stack = []
        st.session_state.redo_stack = []

        invalidate_analysis(
            "New dataset activated"
        )

    add_audit(
        "Dataset imported",
        f"{name} | {len(df):,} rows × "
        f"{len(df.columns):,} columns | {source}",
    )

    return key


def invalidate_analysis(
    reason: str = ""
):

    st.session_state.last_forecast = None
    st.session_state.last_forecast_signature = None
    st.session_state.forecast_config = {}

    st.session_state.dashboard_saved = False

    st.session_state.findings = []
    st.session_state.evidence_ledger = []
    st.session_state.verification_results = []
    st.session_state.challenge_results = []

    if reason:
        add_audit(
            "Analysis invalidated",
            reason
        )


def commit_dataframe(
    df: pd.DataFrame,
    action: str,
    details: str = "",
):

    current = current_df()

    if current is None:
        return

    st.session_state.undo_stack.append(
        current.copy(deep=True)
    )

    st.session_state.undo_stack = (
        st.session_state.undo_stack[-MAX_UNDO:]
    )

    st.session_state.redo_stack = []

    st.session_state.datasets[
        st.session_state.active_dataset_key
    ] = df.copy()

    add_cleaning(
        action,
        len(current),
        len(df),
        details,
    )

    add_audit(
        action,
        details
    )

    invalidate_analysis(
        f"Data changed: {action}"
    )


def undo_change() -> bool:

    if (
        not st.session_state.undo_stack
        or current_df() is None
    ):
        return False

    st.session_state.redo_stack.append(
        current_df().copy(deep=True)
    )

    st.session_state.datasets[
        st.session_state.active_dataset_key
    ] = st.session_state.undo_stack.pop()

    add_audit(
        "Undo",
        "Restored previous dataset state"
    )

    invalidate_analysis(
        "Undo"
    )

    return True


def redo_change() -> bool:

    if (
        not st.session_state.redo_stack
        or current_df() is None
    ):
        return False

    st.session_state.undo_stack.append(
        current_df().copy(deep=True)
    )

    st.session_state.datasets[
        st.session_state.active_dataset_key
    ] = st.session_state.redo_stack.pop()

    add_audit(
        "Redo",
        "Restored next dataset state"
    )

    invalidate_analysis(
        "Redo"
    )

    return True


# ============================================================
# DEMO DATA
# ============================================================

@st.cache_data
def build_demo_dataset() -> pd.DataFrame:

    rng = np.random.default_rng(42)

    dates = pd.date_range(
        "2024-01-01",
        periods=720,
        freq="D"
    )

    regions = np.array(
        ["North", "South", "East", "West"]
    )

    segments = np.array(
        ["Consumer", "SMB", "Enterprise"]
    )

    channels = np.array(
        ["Online", "Partner", "Direct"]
    )

    region = rng.choice(
        regions,
        len(dates)
    )

    segment = rng.choice(
        segments,
        len(dates),
        p=[.52, .30, .18]
    )

    channel = rng.choice(
        channels,
        len(dates),
        p=[.55, .25, .20]
    )

    season = (
        1
        + .10
        * np.sin(
            np.arange(len(dates)) / 18
        )
    )

    revenue = (
        rng.normal(
            16000,
            3000,
            len(dates)
        )
        * season
    )

    cost = (
        revenue
        * rng.normal(
            .64,
            .07,
            len(dates)
        )
    )

    orders = rng.poisson(
        np.maximum(
            revenue / 120,
            5
        )
    )

    conversion = np.clip(
        rng.normal(
            .073,
            .014,
            len(dates)
        ),
        .01,
        .20,
    )

    rating = np.clip(
        rng.normal(
            4.1,
            .35,
            len(dates)
        ),
        1,
        5,
    )

    df = pd.DataFrame(
        {
            "Date": dates,
            "Region": region,
            "Segment": segment,
            "Channel": channel,
            "Revenue": revenue.round(2),
            "Cost": cost.round(2),
            "Orders": orders,
            "ConversionRate": conversion.round(4),
            "CustomerRating": rating.round(2),
        }
    )

    df["Profit"] = (
        df["Revenue"]
        - df["Cost"]
    ).round(2)

    missing = rng.choice(
        df.index,
        22,
        replace=False
    )

    df.loc[
        missing[:8],
        "Region"
    ] = None

    df.loc[
        missing[8:15],
        "Revenue"
    ] = np.nan

    df.loc[
        missing[15:],
        "CustomerRating"
    ] = np.nan

    outliers = rng.choice(
        df.index,
        7,
        replace=False
    )

    df.loc[
        outliers,
        "Revenue"
    ] *= 4.5

    duplicates = df.iloc[
        [5, 25, 125, 225]
    ].copy()

    df = pd.concat(
        [df, duplicates],
        ignore_index=True
    )

    df = df.drop_duplicates(
        keep="first"
    ).reset_index(
        drop=True
    )

    return df.head(720).copy()


# ============================================================
# IMPORT
# ============================================================

def extract_structured_python(
    text: str
) -> pd.DataFrame:

    tree = ast.parse(text)

    candidates = []

    for node in ast.walk(tree):

        if not isinstance(
            node,
            (ast.Assign, ast.AnnAssign)
        ):
            continue

        try:

            value_node = (
                node.value
                if isinstance(
                    node,
                    ast.AnnAssign
                )
                else node.value
            )

            value = ast.literal_eval(
                value_node
            )

            if (
                isinstance(value, (list, tuple))
                and value
                and all(
                    isinstance(x, dict)
                    for x in value
                )
            ):
                candidates.append(value)

            elif (
                isinstance(value, dict)
                and value
                and all(
                    isinstance(v, (list, tuple))
                    for v in value.values()
                )
            ):
                candidates.append(value)

        except Exception:
            pass

    if not candidates:

        raise ValueError(
            "No safely parseable tabular Python "
            "literal was found. Uploaded Python "
            "is never executed."
        )

    value = candidates[0]

    if isinstance(value, dict):
        return pd.DataFrame(value)

    return pd.DataFrame(value)


def extract_structured_java(
    text: str
) -> pd.DataFrame:

    rows = []

    for match in re.finditer(
        r"\{\s*([^{}]+?)\s*\}",
        text,
        re.S,
    ):

        raw = match.group(1)

        if '"' not in raw:
            continue

        parts = re.split(
            r",\s*",
            raw
        )

        parsed = []

        for part in parts:

            part = part.strip()

            if (
                len(part) >= 2
                and part[0] == '"'
                and part[-1] == '"'
            ):
                parsed.append(
                    part[1:-1]
                )

            else:

                try:

                    parsed.append(
                        float(part)
                        if "." in part
                        else int(part)
                    )

                except Exception:

                    parsed.append(part)

        if len(parsed) >= 2:
            rows.append(parsed)

    if not rows:

        raise ValueError(
            "No safely parseable Java tabular "
            "literal was found."
        )

    width = max(
        len(row)
        for row in rows
    )

    rows = [
        row + [None] * (
            width - len(row)
        )
        for row in rows
    ]

    return pd.DataFrame(
        rows,
        columns=[
            f"Column_{i+1}"
            for i in range(width)
        ],
    )


def read_uploaded_file(
    uploaded
) -> Tuple[pd.DataFrame, str]:

    name = uploaded.name
    extension = Path(name).suffix.lower()

    raw = uploaded.getvalue()

    size_mb = len(raw) / (
        1024 * 1024
    )

    if size_mb > MAX_UPLOAD_MB:

        raise ValueError(
            f"File exceeds the {MAX_UPLOAD_MB} MB limit."
        )

    if extension == ".csv":

        return (
            pd.read_csv(
                io.BytesIO(raw)
            ),
            "CSV",
        )

    if extension in {
        ".xlsx",
        ".xlsm",
    }:

        excel = pd.ExcelFile(
            io.BytesIO(raw)
        )

        frames = []

        for sheet in excel.sheet_names:

            part = pd.read_excel(
                io.BytesIO(raw),
                sheet_name=sheet,
            )

            if len(excel.sheet_names) > 1:

                part.insert(
                    0,
                    "_source_sheet",
                    sheet,
                )

            frames.append(part)

        return (
            pd.concat(
                frames,
                ignore_index=True,
                sort=False,
            ),
            f"Excel ({len(frames)} sheet"
            f"{'s' if len(frames) != 1 else ''})",
        )

    if extension == ".xls":

        return (
            pd.read_excel(
                io.BytesIO(raw)
            ),
            "Excel XLS",
        )

    if extension == ".json":

        obj = json.loads(
            raw.decode(
                "utf-8",
                errors="replace"
            )
        )

        if isinstance(obj, dict):

            for key in (
                "data",
                "rows",
                "records",
                "items",
                "results",
            ):

                if isinstance(
                    obj.get(key),
                    list,
                ):

                    obj = obj[key]
                    break

        if not isinstance(
            obj,
            list
        ):

            raise ValueError(
                "JSON must contain a list of records."
            )

        return (
            pd.json_normalize(obj),
            "JSON",
        )

    if extension == ".parquet":

        return (
            pd.read_parquet(
                io.BytesIO(raw)
            ),
            "Parquet",
        )

    if extension in {
        ".tsv",
        ".txt",
    }:

        try:

            return (
                pd.read_csv(
                    io.BytesIO(raw),
                    sep="\t",
                ),
                "TSV/TXT",
            )

        except Exception:

            return (
                pd.read_csv(
                    io.BytesIO(raw)
                ),
                "Text/CSV",
            )

    if extension == ".py":

        return (
            extract_structured_python(
                raw.decode(
                    "utf-8",
                    errors="replace"
                )
            ),
            "Python static extraction",
        )

    if extension == ".java":

        return (
            extract_structured_java(
                raw.decode(
                    "utf-8",
                    errors="replace"
                )
            ),
            "Java static extraction",
        )

    raise ValueError(
        "Unsupported format."
    )


# ============================================================
# DATA QUALITY
# ============================================================

def quality_report(
    df: pd.DataFrame
) -> Dict[str, Any]:

    rows, columns = df.shape

    missing = int(
        df.isna().sum().sum()
    )

    duplicates = int(
        df.duplicated().sum()
    )

    cells = max(
        rows * max(columns, 1),
        1
    )

    missing_rate = (
        missing / cells
    )

    duplicate_rate = (
        duplicates / max(rows, 1)
    )

    constants = [
        c
        for c in df.columns
        if df[c].nunique(
            dropna=False
        ) <= 1
    ]

    score = max(
        0,
        min(
            100,
            100
            - missing_rate * 65
            - duplicate_rate * 25
            - len(constants) * 2,
        ),
    )

    return {
        "rows": rows,
        "columns": columns,
        "missing_cells": missing,
        "missing_rate": missing_rate,
        "duplicate_rows": duplicates,
        "duplicate_rate": duplicate_rate,
        "constant_columns": constants,
        "score": round(score),
    }


def schema_report(
    df: pd.DataFrame
) -> pd.DataFrame:

    rows = []

    for column in df.columns:

        series = df[column]

        if is_numeric_series(series):
            inferred = "numeric"

        elif pd.api.types.is_datetime64_any_dtype(
            series
        ):
            inferred = "datetime"

        else:
            inferred = "categorical/text"

        rows.append(
            {
                "Column": column,
                "Pandas dtype": str(
                    series.dtype
                ),
                "Inferred type": inferred,
                "Non-null": int(
                    series.notna().sum()
                ),
                "Unique": int(
                    series.nunique(
                        dropna=True
                    )
                ),
                "Missing": int(
                    series.isna().sum()
                ),
            }
        )

    return pd.DataFrame(rows)


def missing_report(
    df: pd.DataFrame
) -> pd.DataFrame:

    rows = []

    for column in df.columns:

        count = int(
            df[column].isna().sum()
        )

        if count:

            rows.append(
                {
                    "Column": column,
                    "Missing": count,
                    "Missing %": round(
                        count / len(df) * 100,
                        2,
                    ),
                }
            )

    return pd.DataFrame(rows)


def outlier_report(
    df: pd.DataFrame
) -> pd.DataFrame:

    rows = []

    for column in numeric_columns(df):

        series = pd.to_numeric(
            df[column],
            errors="coerce"
        ).dropna()

        if len(series) < 5:
            continue

        q1 = series.quantile(.25)
        q3 = series.quantile(.75)

        iqr = q3 - q1

        if iqr == 0:

            lower = q1
            upper = q3
            count = 0

        else:

            lower = q1 - 1.5 * iqr
            upper = q3 + 1.5 * iqr

            count = int(
                (
                    (series < lower)
                    | (series > upper)
                ).sum()
            )

        rows.append(
            {
                "Column": column,
                "Outliers": count,
                "Lower fence": lower,
                "Upper fence": upper,
            }
        )

    return pd.DataFrame(rows)


# ============================================================
# CLEANING
# ============================================================

def convert_column_type(
    df: pd.DataFrame,
    column: str,
    target: str,
):

    out = df.copy()

    if target == "Numeric":

        out[column] = pd.to_numeric(
            out[column],
            errors="coerce"
        )

    elif target == "Date/Time":

        out[column] = pd.to_datetime(
            out[column],
            errors="coerce"
        )

    elif target == "Text":

        out[column] = out[column].astype(
            "string"
        )

    return out


def clean_text(
    df: pd.DataFrame,
    column: str,
    operation: str,
):

    out = df.copy()

    series = out[column].astype("string")

    if operation == "Trim whitespace":

        out[column] = series.str.strip()

    elif operation == "Lowercase":

        out[column] = series.str.lower()

    elif operation == "Uppercase":

        out[column] = series.str.upper()

    elif operation == "Title Case":

        out[column] = series.str.title()

    elif operation == "Collapse spaces":

        out[column] = (
            series
            .str.replace(
                r"\s+",
                " ",
                regex=True
            )
            .str.strip()
        )

    return out


def apply_transform(
    df: pd.DataFrame,
    column: str,
    operation: str,
    parameter: str = "",
):

    out = df.copy()

    if operation == "Trim whitespace":

        out[column] = (
            out[column]
            .astype("string")
            .str.strip()
        )

    elif operation == "Lowercase":

        out[column] = (
            out[column]
            .astype("string")
            .str.lower()
        )

    elif operation == "Uppercase":

        out[column] = (
            out[column]
            .astype("string")
            .str.upper()
        )

    elif operation == "Title Case":

        out[column] = (
            out[column]
            .astype("string")
            .str.title()
        )

    elif operation == "Fill missing":

        if pd.api.types.is_numeric_dtype(
            out[column]
        ):

            method = parameter or "Median"

            if method == "Median":
                value = out[column].median()

            elif method == "Mean":
                value = out[column].mean()

            else:
                value = 0

            out[column] = out[column].fillna(
                value
            )

        else:

            value = (
                parameter
                if parameter
                else "Unknown"
            )

            out[column] = out[column].fillna(
                value
            )

    elif operation == "Convert numeric":

        out[column] = pd.to_numeric(
            out[column],
            errors="coerce"
        )

    elif operation == "Convert date":

        out[column] = pd.to_datetime(
            out[column],
            errors="coerce"
        )

    elif operation == "Replace value":

        out[column] = (
            out[column]
            .astype("string")
            .replace(
                parameter.split(
                    "=>",
                    1
                )[0].strip()
                if "=>" in parameter
                else parameter,
                parameter.split(
                    "=>",
                    1
                )[1].strip()
                if "=>" in parameter
                else "",
            )
        )

    return out


# ============================================================
# ANALYTICS
# ============================================================

def auto_kpis(
    df: pd.DataFrame
) -> List[Dict[str, Any]]:

    kpis = []

    nums = numeric_columns(df)

    if "Revenue" in nums:

        kpis.append(
            {
                "label": "Revenue",
                "value": df["Revenue"].sum(),
            }
        )

    if "Cost" in nums:

        kpis.append(
            {
                "label": "Cost",
                "value": df["Cost"].sum(),
            }
        )

    if "Profit" in nums:

        kpis.append(
            {
                "label": "Profit",
                "value": df["Profit"].sum(),
            }
        )

    elif (
        "Revenue" in nums
        and "Cost" in nums
    ):

        kpis.append(
            {
                "label": "Profit",
                "value": (
                    df["Revenue"]
                    - df["Cost"]
                ).sum(),
            }
        )

    if "Orders" in nums:

        kpis.append(
            {
                "label": "Orders",
                "value": df["Orders"].sum(),
            }
        )

    for column in nums:

        if len(kpis) >= 6:
            break

        if not any(
            x["label"] == column
            for x in kpis
        ):

            kpis.append(
                {
                    "label": column,
                    "value": df[column].mean(),
                }
            )

    return kpis[:6]


def monthly_series(
    df: pd.DataFrame,
    date_column: str,
    metric: str,
) -> pd.DataFrame:

    temp = pd.DataFrame()

    temp["Date"] = pd.to_datetime(
        df[date_column],
        errors="coerce"
    )

    temp["Value"] = pd.to_numeric(
        df[metric],
        errors="coerce"
    )

    temp = (
        temp
        .dropna()
        .sort_values("Date")
    )

    if temp.empty:
        return pd.DataFrame(
            columns=[
                "Date",
                "Value",
            ]
        )

    temp["Date"] = (
        temp["Date"]
        .dt.to_period("M")
        .dt.to_timestamp()
    )

    return (
        temp
        .groupby(
            "Date",
            as_index=False
        )["Value"]
        .sum()
    )


def forecast_series(
    df: pd.DataFrame,
    date_column: str,
    metric: str,
    periods: int = 6,
):

    series = monthly_series(
        df,
        date_column,
        metric
    )

    if len(series) < 4:

        raise ValueError(
            "At least four monthly observations "
            "are required."
        )

    values = series["Value"].astype(float).values

    future_dates = pd.date_range(
        series["Date"].iloc[-1]
        + pd.offsets.MonthBegin(1),
        periods=periods,
        freq="MS",
    )

    method = "Linear trend"

    if (
        ExponentialSmoothing is not None
        and len(values) >= 8
    ):

        try:

            model = ExponentialSmoothing(
                values,
                trend="add",
                damped_trend=True,
                initialization_method="estimated",
            ).fit(
                optimized=True
            )

            fitted = model.fittedvalues
            predictions = model.forecast(
                periods
            )

            method = (
                "Holt-Winters damped trend"
            )

        except Exception:

            coefficients = np.polyfit(
                np.arange(len(values)),
                values,
                1,
            )

            fitted = np.polyval(
                coefficients,
                np.arange(len(values))
            )

            predictions = np.polyval(
                coefficients,
                np.arange(
                    len(values),
                    len(values) + periods
                ),
            )

            method = "Linear trend fallback"

    else:

        coefficients = np.polyfit(
            np.arange(len(values)),
            values,
            1,
        )

        fitted = np.polyval(
            coefficients,
            np.arange(len(values))
        )

        predictions = np.polyval(
            coefficients,
            np.arange(
                len(values),
                len(values) + periods
            ),
        )

    mae = float(
        np.mean(
            np.abs(
                values - fitted
            )
        )
    )

    rmse = float(
        np.sqrt(
            np.mean(
                (values - fitted) ** 2
            )
        )
    )

    denominator = np.where(
        np.abs(values) < 1e-9,
        np.nan,
        np.abs(values)
    )

    mape = float(
        np.nanmean(
            np.abs(
                (values - fitted)
                / denominator
            )
        )
        * 100
    )

    history = series.copy()
    history["Type"] = "Historical"

    future = pd.DataFrame(
        {
            "Date": future_dates,
            "Value": predictions,
            "Type": "Forecast",
        }
    )

    return (
        pd.concat(
            [history, future],
            ignore_index=True,
        ),
        {
            "method": method,
            "mae": mae,
            "rmse": rmse,
            "mape": mape,
            "history_points": len(values),
            "forecast_points": periods,
        },
    )


def correlation_df(
    df: pd.DataFrame
) -> pd.DataFrame:

    nums = numeric_columns(df)

    if len(nums) < 2:
        return pd.DataFrame()

    return df[nums].corr(
        numeric_only=True
    )


# ============================================================
# EVIDENCE ENGINE
# ============================================================

def build_evidence_ledger(
    df: pd.DataFrame
) -> List[Dict[str, Any]]:

    quality = quality_report(df)

    ledger = []

    def add(
        claim,
        value,
        method,
    ):

        ledger.append(
            {
                "id": f"E{len(ledger)+1:03d}",
                "claim": claim,
                "value": value,
                "source": "Deterministic analytics engine",
                "method": method,
                "status": "OBSERVED",
                "dataset_signature":
                    dataframe_signature(df),
            }
        )

    add(
        "Dataset size",
        f"{len(df):,} rows × "
        f"{len(df.columns):,} columns",
        "shape",
    )

    add(
        "Data quality score",
        f"{quality['score']}/100",
        "missingness + duplicates + constants",
    )

    add(
        "Missing cells",
        f"{quality['missing_cells']:,}",
        "null count",
    )

    add(
        "Exact duplicate rows",
        f"{quality['duplicate_rows']:,}",
        "duplicate detection",
    )

    nums = numeric_columns(df)

    for column in nums[:8]:

        series = pd.to_numeric(
            df[column],
            errors="coerce"
        ).dropna()

        if len(series):

            add(
                f"{column} total",
                format_value(
                    series.sum()
                ),
                f"sum({column})",
            )

    cats = categorical_columns(df)

    for column in cats[:4]:

        groups = (
            df.groupby(
                column,
                dropna=False
            )[nums[0]]
            .sum()
            .sort_values(
                ascending=False
            )
            if nums
            else pd.Series()
        )

        if not groups.empty:

            add(
                f"Top {column} contribution",
                str(groups.index[0]),
                f"groupby({column}) sum",
            )

            add(
                f"Lowest {column} contribution",
                str(groups.index[-1]),
                f"groupby({column}) sum",
            )

    outliers = outlier_report(df)

    for _, row in (
        outliers[
            outliers["Outliers"] > 0
        ]
        .head(4)
        .iterrows()
    ):

        add(
            f"Potential outliers: "
            f"{row['Column']}",
            f"{int(row['Outliers'])} outliers",
            "IQR = Q3-Q1; "
            "bounds = Q1 ± 1.5×IQR",
        )

    dates = detect_datetime_candidates(df)

    if dates and nums:

        monthly = monthly_series(
            df,
            dates[0],
            nums[0]
        )

        if len(monthly) >= 2:

            previous = monthly.iloc[-2]["Value"]
            latest = monthly.iloc[-1]["Value"]

            movement = (
                (latest - previous)
                / abs(previous)
                * 100
                if previous
                else 0
            )

            add(
                "Latest period movement",
                f"{movement:.2f}%",
                "(latest-previous)/abs(previous)",
            )

    return ledger


def detective_findings(
    df: pd.DataFrame
) -> List[Dict[str, str]]:

    quality = quality_report(df)

    findings = []

    if quality["missing_cells"]:

        findings.append(
            {
                "severity": "warning",
                "title":
                    "Missing values detected",
                "text":
                    f"{quality['missing_cells']:,} "
                    "missing cells were detected.",
            }
        )

    if quality["duplicate_rows"]:

        findings.append(
            {
                "severity": "warning",
                "title":
                    "Duplicate rows detected",
                "text":
                    f"{quality['duplicate_rows']:,} "
                    "exact duplicate rows were detected.",
            }
        )

    for column in quality[
        "constant_columns"
    ]:

        findings.append(
            {
                "severity": "info",
                "title":
                    "Constant column",
                "text":
                    f"{column} contains one unique value.",
            }
        )

    outliers = outlier_report(df)

    for _, row in (
        outliers[
            outliers["Outliers"] > 0
        ]
        .head(8)
        .iterrows()
    ):

        findings.append(
            {
                "severity": "warning",
                "title":
                    f"Potential outliers: "
                    f"{row['Column']}",
                "text":
                    f"{int(row['Outliers'])} "
                    "IQR outliers detected.",
            }
        )

    return findings[:20]


# ============================================================
# SQL
# ============================================================

def validate_sql(
    query: str
):

    query = query.strip()

    if not re.match(
        r"^(SELECT|WITH)\b",
        query,
        re.I,
    ):

        raise ValueError(
            "Only SELECT/WITH queries are allowed."
        )

    if (
        ";"
        in query
        or "--"
        in query
        or "/*"
        in query
    ):

        raise ValueError(
            "Multiple statements/comments are blocked."
        )

    forbidden = (
        r"\b("
        r"INSERT|UPDATE|DELETE|DROP|ALTER|CREATE|"
        r"ATTACH|DETACH|COPY|EXPORT|INSTALL|LOAD|"
        r"PRAGMA|READ_CSV|READ_PARQUET|HTTPFS|"
        r"GLOB|SQLITE_SCAN|POSTGRES_SCAN|"
        r"MYSQL_SCAN"
        r")\b"
    )

    if re.search(
        forbidden,
        query,
        re.I,
    ):

        raise ValueError(
            "Query contains a blocked operation."
        )

    return query


def execute_sql(
    df: pd.DataFrame,
    query: str
) -> pd.DataFrame:

    if duckdb is None:

        raise RuntimeError(
            "DuckDB is not installed."
        )

    query = validate_sql(query)

    connection = duckdb.connect(
        ":memory:"
    )

    try:

        connection.execute(
            "SET enable_external_access=false"
        )

        connection.register(
            "data",
            df
        )

        result = connection.execute(
            query
        ).fetchdf()

        return result.head(
            MAX_SQL_ROWS
        )

    finally:

        connection.close()


# ============================================================
# OLLAMA
# ============================================================

def ollama_health_check() -> Tuple[bool, str]:

    if requests is None:

        return (
            False,
            "requests package missing"
        )

    try:

        response = requests.get(
            OLLAMA_TAGS_URL,
            timeout=8,
        )

        return (
            response.ok,
            f"HTTP {response.status_code}",
        )

    except Exception as exc:

        return (
            False,
            str(exc)[:120],
        )


def get_ollama_models() -> List[str]:

    if requests is None:
        return []

    try:

        response = requests.get(
            OLLAMA_TAGS_URL,
            timeout=8,
        )

        response.raise_for_status()

        return [
            model.get("name")
            for model in response.json().get(
                "models",
                []
            )
            if model.get("name")
        ]

    except Exception:

        return []


def is_general_question(
    question: str
) -> bool:

    q = question.lower().strip()

    data_words = [
        "dataset",
        "data",
        "revenue",
        "profit",
        "cost",
        "sales",
        "orders",
        "rows",
        "columns",
        "missing",
        "duplicate",
        "outlier",
        "forecast",
        "trend",
        "region",
        "segment",
        "metric",
        "quality",
        "correlation",
        "calculate",
        "analysis",
    ]

    return not any(
        word in q
        for word in data_words
    )


def compact_evidence(
    df: pd.DataFrame
) -> Dict[str, Any]:

    quality = quality_report(df)

    evidence = build_evidence_ledger(
        df
    )

    return {
        "dataset": current_dataset_name(),
        "rows": len(df),
        "columns": len(df.columns),
        "quality": quality,
        "evidence": evidence[:18],
    }


def build_ai_prompt(
    question: str,
    df: pd.DataFrame,
) -> str:

    evidence = compact_evidence(
        df
    )

    return f"""
You are a local AI Data Analyst.

Use ONLY the supplied deterministic evidence.

IMPORTANT:
- Never invent numbers.
- Never invent columns.
- Never invent correlations.
- Never invent causes.
- Never claim causation unless the evidence proves it.
- Distinguish measured facts from hypotheses.
- If the evidence is insufficient, say so.
- Do not mention information that is not in the evidence.

DATASET:
{json.dumps(evidence, default=str)[:14000]}

USER QUESTION:
{question}

Answer in this structure:

Direct answer

Evidence
- cite exact values from the evidence when available

Interpretation
- explain what the evidence means

Hypothesis
- only if supported by the available evidence

Next verification
- state what should be checked next
"""


def build_general_prompt(
    question: str
) -> str:

    return f"""
You are a helpful local AI assistant.

Answer the user's question clearly and naturally.

Do not mention datasets.
Do not pretend the question came from a dataset.
Do not invent facts.

USER QUESTION:
{question}
"""


def stream_ollama_answer(
    prompt: str,
    model: str,
):

    if requests is None:

        raise RuntimeError(
            "requests package is not installed."
        )

    payload = {
        "model": model,
        "prompt": prompt,
        "stream": True,
        "keep_alive": "10m",
        "options": {
            "temperature":
                AI_TEMPERATURE,
            "top_p": .9,
            "num_ctx":
                AI_NUM_CTX,
            "num_predict":
                AI_NUM_PREDICT,
            "repeat_penalty":
                1.05,
        },
    }

    with requests.post(
        OLLAMA_GENERATE_URL,
        json=payload,
        stream=True,
        timeout=(
            OLLAMA_CONNECT_TIMEOUT,
            OLLAMA_READ_TIMEOUT,
        ),
    ) as response:

        response.raise_for_status()

        started = time.time()

        for line in response.iter_lines():

            if (
                time.time() - started
                > OLLAMA_READ_TIMEOUT
            ):
                raise TimeoutError(
                    "Ollama generation timed out."
                )

            if not line:
                continue

            try:

                payload = json.loads(
                    line.decode(
                        "utf-8"
                    )
                )

                token = payload.get(
                    "response",
                    ""
                )

                if token:
                    yield token

                if payload.get(
                    "done"
                ):
                    break

            except Exception:
                continue


def deterministic_ai_fallback(
    question: str,
    df: pd.DataFrame,
) -> str:

    q = question.lower()

    quality = quality_report(
        df
    )

    if "missing" in q:

        return (
            f"The dataset contains "
            f"{quality['missing_cells']:,} "
            "missing cells. Review the "
            "Data Quality page to identify "
            "the affected columns."
        )

    if "duplicate" in q:

        return (
            f"The dataset contains "
            f"{quality['duplicate_rows']:,} "
            "exact duplicate rows."
        )

    if "quality" in q:

        return (
            f"The current deterministic "
            f"data-quality score is "
            f"{quality['score']}/100."
        )

    nums = numeric_columns(df)

    if nums:

        column = nums[0]

        series = pd.to_numeric(
            df[column],
            errors="coerce"
        ).dropna()

        return (
            f"Deterministic fallback: "
            f"{column} has "
            f"{len(series):,} usable values, "
            f"mean {format_value(series.mean())}, "
            f"median {format_value(series.median())}, "
            f"minimum {format_value(series.min())}, "
            f"maximum {format_value(series.max())}."
        )

    return (
        f"The dataset contains "
        f"{len(df):,} rows and "
        f"{len(df.columns):,} columns."
    )


# ============================================================
# VERIFICATION
# ============================================================

def verify_finding(
    df: pd.DataFrame,
    finding: Dict[str, Any],
) -> Dict[str, Any]:

    claim = str(
        finding.get("claim")
        or finding.get("text")
        or ""
    )

    result = {
        "time": now_text(),
        "finding_id":
            finding.get("id"),
        "claim": claim,
        "status": "INCONCLUSIVE",
        "summary": "",
    }

    nums = numeric_columns(df)

    metric = next(
        (
            column
            for column in nums
            if re.search(
                rf"\b{re.escape(str(column))}\b",
                claim,
                re.I,
            )
        ),
        None,
    )

    if metric:

        series = pd.to_numeric(
            df[metric],
            errors="coerce"
        ).dropna()

        result["status"] = "VERIFIED"

        result["summary"] = (
            f"Recalculated {metric}: "
            f"{len(series):,} usable values; "
            f"total {format_value(series.sum())}; "
            f"mean {format_value(series.mean())}; "
            f"median {format_value(series.median())}."
        )

    elif "missing" in claim.lower():

        count = int(
            df.isna().sum().sum()
        )

        result["status"] = (
            "VERIFIED"
            if str(count) in claim
            else "INCONCLUSIVE"
        )

        result["summary"] = (
            f"Recalculated missing cells: "
            f"{count:,}."
        )

    else:

        result["summary"] = (
            "The deterministic engine could "
            "not map the natural-language claim "
            "to a direct calculation."
        )

    return result


def challenge_finding(
    df: pd.DataFrame,
    finding: Dict[str, Any],
) -> Dict[str, Any]:

    claim = str(
        finding.get("claim")
        or finding.get("text")
        or ""
    )

    result = {
        "time": now_text(),
        "finding_id":
            finding.get("id"),
        "claim": claim,
        "status": "NOT DISPROVED",
        "summary": "",
        "alternatives": [],
    }

    nums = numeric_columns(df)
    cats = categorical_columns(df)

    metric = next(
        (
            column
            for column in nums
            if re.search(
                rf"\b{re.escape(str(column))}\b",
                claim,
                re.I,
            )
        ),
        None,
    )

    if metric and cats:

        overall = pd.to_numeric(
            df[metric],
            errors="coerce"
        ).mean()

        for dimension in cats[:6]:

            grouped = (
                df.groupby(
                    dimension,
                    dropna=False
                )[metric]
                .mean()
                .dropna()
                .sort_values()
            )

            if len(grouped) >= 2:

                low = grouped.iloc[0]
                high = grouped.iloc[-1]

                spread = (
                    (
                        high - low
                    )
                    / abs(overall)
                    * 100
                    if overall
                    else 0
                )

                if abs(spread) >= 10:

                    result[
                        "alternatives"
                    ].append(
                        {
                            "dimension":
                                dimension,
                            "low":
                                str(grouped.index[0]),
                            "high":
                                str(grouped.index[-1]),
                            "spread_pct":
                                float(spread),
                        }
                    )

    if result["alternatives"]:

        result["status"] = (
            "ALTERNATIVE EVIDENCE FOUND"
        )

        result["summary"] = (
            "The challenge scan found "
            "segment-level differences that "
            "could provide an alternative "
            "explanation. The original finding "
            "is not automatically disproved."
        )

    else:

        result["summary"] = (
            "No strong contradiction was "
            "detected by the automatic challenge "
            "scan. This does not establish causation."
        )

    return result


# ============================================================
# PROJECT SAVE / RESUME
# ============================================================

def project_payload() -> Dict[str, Any]:

    df = current_df()

    return {
        "version": APP_VERSION,
        "project_name":
            st.session_state.project_name,
        "project_created":
            st.session_state.project_created
            or now_text(),
        "source_name":
            st.session_state.source_name,
        "audit_log":
            st.session_state.audit_log,
        "chat_history":
            st.session_state.chat_history,
        "sql_history":
            st.session_state.sql_history,
        "python_history":
            st.session_state.python_history,
        "python_results":
            st.session_state.python_results,
        "cleaning_history":
            st.session_state.cleaning_history,
        "dashboard_config":
            st.session_state.dashboard_config,
        "forecast_config":
            st.session_state.forecast_config,
        "findings":
            st.session_state.findings,
        "verification_results":
            st.session_state.verification_results,
        "challenge_results":
            st.session_state.challenge_results,
        "evidence_ledger":
            st.session_state.evidence_ledger,
        "analysis_notes":
            st.session_state.analysis_notes,
        "report_config":
            st.session_state.report_config,
        "dataset_signature":
            dataframe_signature(df)
            if df is not None
            else None,
    }


def save_project_bytes(
    name: str
) -> Dict[str, bytes]:

    df = current_df()

    if df is None:

        raise ValueError(
            "No active dataset to save."
        )

    payload = project_payload()

    payload["project_name"] = name

    stem = safe_filename(name)

    return {
        f"{stem}.csv":
            df.to_csv(
                index=False
            ).encode(
                "utf-8-sig"
            ),

        f"{stem}.json":
            json.dumps(
                payload,
                indent=2,
                default=str,
            ).encode(
                "utf-8"
            ),
    }


def persist_project_to_disk(
    name: str
):

    files = save_project_bytes(
        name
    )

    stem = safe_filename(
        name
    )

    (
        PROJECT_DIR
        / f"{stem}.csv"
    ).write_bytes(
        files[f"{stem}.csv"]
    )

    (
        PROJECT_DIR
        / f"{stem}.json"
    ).write_bytes(
        files[f"{stem}.json"]
    )

    st.session_state.project_name = name
    st.session_state.project_created = now_text()

    add_audit(
        "Project saved",
        name
    )


def normalize_project_entry(
    entry
):

    if isinstance(
        entry,
        Path
    ):
        return entry

    if isinstance(
        entry,
        str
    ):
        return PROJECT_DIR / entry

    if isinstance(
        entry,
        dict
    ):

        path = entry.get(
            "path"
        ) or entry.get(
            "file"
        )

        if path:
            return Path(path)

    if isinstance(
        entry,
        (list, tuple)
    ):

        for value in entry:

            if isinstance(
                value,
                str
            ):
                return PROJECT_DIR / value

    return None


def project_list():

    return sorted(
        PROJECT_DIR.glob(
            "*.json"
        ),
        key=lambda path:
            path.stat().st_mtime,
        reverse=True,
    )


def load_project_from_disk(
    meta_path: Path
):

    metadata = json.loads(
        meta_path.read_text(
            encoding="utf-8"
        )
    )

    project_name = metadata.get(
        "project_name",
        meta_path.stem
    )

    csv_path = (
        PROJECT_DIR
        / f"{safe_filename(project_name)}.csv"
    )

    if not csv_path.exists():

        csv_path = (
            PROJECT_DIR
            / f"{meta_path.stem}.csv"
        )

    if not csv_path.exists():

        raise FileNotFoundError(
            "Project CSV file was not found."
        )

    df = pd.read_csv(
        csv_path
    )

    for column in detect_datetime_candidates(
        df
    ):

        converted = pd.to_datetime(
            df[column],
            errors="coerce"
        )

        if converted.notna().mean() >= .8:

            df[column] = converted

    register_dataset(
        df,
        metadata.get(
            "source_name",
            project_name
        ),
        "project resume",
        True,
    )

    restore_keys = [
        "audit_log",
        "chat_history",
        "sql_history",
        "python_history",
        "python_results",
        "cleaning_history",
        "dashboard_config",
        "forecast_config",
        "findings",
        "verification_results",
        "challenge_results",
        "evidence_ledger",
        "analysis_notes",
        "report_config",
    ]

    for key in restore_keys:

        if key in metadata:

            value = metadata[key]

            if not isinstance(
                value,
                (
                    dict,
                    list,
                ),
            ) and key not in {
                "report_config",
                "dashboard_config",
                "forecast_config",
            }:
                value = []

            st.session_state[key] = value

    st.session_state.project_name = project_name

    st.session_state.project_created = (
        metadata.get(
            "project_created"
        )
    )

    add_audit(
        "Project resumed",
        project_name
    )


# ============================================================
# REPORT STUDIO CONFIG
# ============================================================

def default_report_config(
    df: pd.DataFrame
) -> Dict[str, Any]:

    nums = numeric_columns(df)
    dates = detect_datetime_candidates(df)

    metric = (
        "Revenue"
        if "Revenue" in nums
        else nums[0]
        if nums
        else None
    )

    date_column = (
        dates[0]
        if dates
        else None
    )

    ledger = build_evidence_ledger(
        df
    )

    return {
        "title":
            "AI Data Analyst — Executive Intelligence Report",

        "subtitle":
            "Evidence-first analysis • "
            "Deterministic calculations • "
            "Local AI",

        "sections": [
            "Executive Summary",
            "KPI Snapshot",
            "Dashboard Charts",
            "Forecast",
            "What-If",
            "Evidence",
            "Verification",
            "Challenge",
            "Data Quality",
            "AI Conversation",
            "Audit / Provenance",
        ],

        "charts": [
            "trend",
            "segment",
        ],

        "metric": metric,

        "date": date_column,

        "evidence_ids": [
            item["id"]
            for item in ledger
        ],

        "forecast_details": [
            "summary",
            "metrics",
            "chart",
            "table",
        ],

        "what_if_change": 10,
    }


def get_report_config(
    df: pd.DataFrame
) -> Dict[str, Any]:

    current = st.session_state.get(
        "report_config"
    ) or {}

    defaults = default_report_config(
        df
    )

    defaults.update(
        {
            key: value
            for key, value in current.items()
            if value is not None
        }
    )

    return defaults


# ============================================================
# REPORT CHART GENERATION
# ============================================================

def report_chart_png(
    df: pd.DataFrame,
    chart_type: str,
    metric: Optional[str],
    date_column: Optional[str],
    dimension: Optional[str] = None,
    forecast: Optional[pd.DataFrame] = None,
    what_if_change: float = 10,
) -> Optional[bytes]:

    if plt is None:
        return None

    if metric is None:
        return None

    figure, axis = plt.subplots(
        figsize=(11, 5.8),
        dpi=160,
    )

    try:

        if chart_type == "trend":

            if not date_column:
                raise ValueError(
                    "No date column"
                )

            series = monthly_series(
                df,
                date_column,
                metric,
            )

            if series.empty:
                raise ValueError(
                    "No time-series data"
                )

            axis.plot(
                series["Date"],
                series["Value"],
                marker="o",
                linewidth=2,
                label="Actual",
            )

            axis.set_title(
                f"{metric} — Monthly Trend",
                loc="left",
                fontsize=16,
                fontweight="bold",
            )

            axis.set_ylabel(
                metric
            )

            axis.grid(
                alpha=.18
            )

            axis.legend(
                frameon=False
            )

            figure.autofmt_xdate()

        elif chart_type == "segment":

            if not dimension:
                cats = categorical_columns(
                    df
                )

                if not cats:
                    raise ValueError(
                        "No dimension"
                    )

                dimension = cats[0]

            grouped = (
                df.groupby(
                    dimension,
                    dropna=False
                )[metric]
                .sum()
                .sort_values(
                    ascending=False
                )
                .head(12)
            )

            axis.bar(
                [
                    str(x)
                    for x in grouped.index
                ],
                grouped.values,
            )

            axis.set_title(
                f"{metric} by {dimension}",
                loc="left",
                fontsize=16,
                fontweight="bold",
            )

            axis.set_ylabel(
                f"Total {metric}"
            )

            axis.tick_params(
                axis="x",
                rotation=35
            )

            axis.grid(
                axis="y",
                alpha=.18
            )

        elif chart_type == "distribution":

            values = pd.to_numeric(
                df[metric],
                errors="coerce"
            ).dropna()

            axis.hist(
                values,
                bins=24
            )

            axis.set_title(
                f"{metric} — Distribution",
                loc="left",
                fontsize=16,
                fontweight="bold",
            )

            axis.set_xlabel(
                metric
            )

            axis.set_ylabel(
                "Records"
            )

            axis.grid(
                alpha=.18
            )

        elif chart_type == "correlation":

            correlation = correlation_df(
                df
            )

            if correlation.empty:
                raise ValueError(
                    "Not enough numeric columns"
                )

            image = axis.imshow(
                correlation.values,
                aspect="auto"
            )

            axis.set_xticks(
                range(
                    len(correlation.columns)
                )
            )

            axis.set_xticklabels(
                correlation.columns,
                rotation=45,
                ha="right",
                fontsize=8,
            )

            axis.set_yticks(
                range(
                    len(correlation.index)
                )
            )

            axis.set_yticklabels(
                correlation.index,
                fontsize=8,
            )

            axis.set_title(
                "Numeric Correlation Matrix",
                loc="left",
                fontsize=16,
                fontweight="bold",
            )

            figure.colorbar(
                image,
                ax=axis,
                fraction=.035,
                pad=.03,
            )

        elif chart_type == "forecast":

            if forecast is None:
                raise ValueError(
                    "Forecast unavailable"
                )

            for type_name, part in forecast.groupby(
                "Type"
            ):

                axis.plot(
                    part["Date"],
                    part["Value"],
                    marker="o",
                    linewidth=2,
                    label=type_name,
                )

            axis.set_title(
                f"{metric} — Historical vs Forecast",
                loc="left",
                fontsize=16,
                fontweight="bold",
            )

            axis.grid(
                alpha=.18
            )

            axis.legend(
                frameon=False
            )

            figure.autofmt_xdate()

        elif chart_type == "what_if":

            if forecast is None:
                raise ValueError(
                    "Forecast unavailable"
                )

            forecast_only = forecast[
                forecast["Type"]
                == "Forecast"
            ]

            if forecast_only.empty:
                raise ValueError(
                    "Forecast unavailable"
                )

            baseline = float(
                forecast_only[
                    "Value"
                ].sum()
            )

            scenario = (
                baseline
                * (
                    1
                    + what_if_change / 100
                )
            )

            axis.bar(
                [
                    "Baseline",
                    "Scenario",
                ],
                [
                    baseline,
                    scenario,
                ],
            )

            axis.set_title(
                f"Forecast What-If "
                f"({what_if_change:+.0f}%)",
                loc="left",
                fontsize=16,
                fontweight="bold",
            )

            axis.set_ylabel(
                metric
            )

            axis.grid(
                axis="y",
                alpha=.18
            )

        else:

            raise ValueError(
                "Unknown chart type"
            )

        figure.tight_layout()

        buffer = io.BytesIO()

        figure.savefig(
            buffer,
            format="png",
            bbox_inches="tight",
        )

        return buffer.getvalue()

    except Exception:

        return None

    finally:

        plt.close(
            figure
        )


def report_chart_images(
    df: pd.DataFrame,
    config: Dict[str, Any],
) -> Dict[str, bytes]:

    images = {}

    forecast = None

    if st.session_state.last_forecast:

        forecast = pd.DataFrame(
            st.session_state.last_forecast
        )

    dashboard = (
        st.session_state.dashboard_config
        or {}
    )

    dimension = dashboard.get(
        "dimension"
    )

    if not dimension:

        cats = categorical_columns(
            df
        )

        dimension = (
            cats[0]
            if cats
            else None
        )

    for chart in config.get(
        "charts",
        [],
    ):

        image = report_chart_png(
            df,
            chart,
            config.get(
                "metric"
            ),
            config.get(
                "date"
            ),
            dimension,
            forecast,
            config.get(
                "what_if_change",
                10,
            ),
        )

        if image:

            images[chart] = image

    return images


# ============================================================
# EXCEL EXPORT
# ============================================================

def dataframe_to_excel(
    df: pd.DataFrame
) -> bytes:

    if not OPENPYXL_AVAILABLE:

        raise RuntimeError(
            "openpyxl is required for Excel export."
        )

    sheets = {
        "Clean Data":
            df,

        "KPIs":
            pd.DataFrame(
                auto_kpis(df)
            ),

        "Schema":
            schema_report(df),

        "Missing":
            missing_report(df),

        "Outliers":
            outlier_report(df),

        "Correlation":
            correlation_df(df),

        "Findings":
            pd.DataFrame(
                st.session_state.findings
                or detective_findings(df)
            ),

        "Evidence":
            pd.DataFrame(
                build_evidence_ledger(df)
            ),

        "Verification":
            pd.DataFrame(
                st.session_state.verification_results
            ),

        "Challenge":
            pd.DataFrame(
                st.session_state.challenge_results
            ),

        "SQL History":
            pd.DataFrame(
                st.session_state.sql_history
            ),

        "Cleaning":
            pd.DataFrame(
                st.session_state.cleaning_history
            ),

        "AI Conversation":
            pd.DataFrame(
                st.session_state.chat_history
            ),

        "Audit":
            pd.DataFrame(
                st.session_state.audit_log
            ),
    }

    if st.session_state.last_forecast:

        forecast_df = pd.DataFrame(
            st.session_state.last_forecast
        )

        sheets["Forecast"] = forecast_df

        sheets[
            "Forecast Metrics"
        ] = pd.DataFrame(
            [
                st.session_state
                .forecast_config
                .get(
                    "metrics",
                    {}
                )
            ]
        )

    if st.session_state.report_config:

        sheets[
            "Report Studio"
        ] = pd.DataFrame(
            [
                {
                    "Report Title":
                        st.session_state
                        .report_config
                        .get(
                            "title",
                            "",
                        ),
                    "Sections":
                        ", ".join(
                            st.session_state
                            .report_config
                            .get(
                                "sections",
                                [],
                            )
                        ),
                    "Charts":
                        ", ".join(
                            st.session_state
                            .report_config
                            .get(
                                "charts",
                                [],
                            )
                        ),
                    "Evidence IDs":
                        ", ".join(
                            st.session_state
                            .report_config
                            .get(
                                "evidence_ids",
                                [],
                            )
                        ),
                }
            ]
        )

    buffer = io.BytesIO()

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl"
    ) as writer:

        for sheet_name, data in sheets.items():

            if not isinstance(
                data,
                pd.DataFrame
            ):
                data = pd.DataFrame(
                    data
                )

            if data.empty:

                data = pd.DataFrame(
                    {
                        "Status":
                            [
                                "No records"
                            ]
                    }
                )

            safe_sheet = re.sub(
                r"[^A-Za-z0-9 ]",
                "",
                sheet_name
            )[:31]

            data.to_excel(
                writer,
                index=False,
                sheet_name=safe_sheet,
            )

        workbook = writer.book

        header_fill = PatternFill(
            "solid",
            fgColor="16324F"
        )

        header_font = Font(
            bold=True,
            color="FFFFFF"
        )

        border = Border(
            bottom=Side(
                style="thin",
                color="D9E2EC"
            )
        )

        for worksheet in workbook.worksheets:

            worksheet.freeze_panes = "A2"

            worksheet.auto_filter.ref = (
                worksheet.dimensions
            )

            for cell in worksheet[1]:

                cell.fill = header_fill
                cell.font = header_font

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                )

                cell.border = border

            worksheet.row_dimensions[
                1
            ].height = 24

            for column_index in range(
                1,
                worksheet.max_column + 1
            ):

                letter = get_column_letter(
                    column_index
                )

                lengths = []

                for row in range(
                    1,
                    min(
                        worksheet.max_row,
                        150
                    ) + 1,
                ):

                    value = worksheet.cell(
                        row,
                        column_index
                    ).value

                    lengths.append(
                        len(
                            str(
                                value
                                if value is not None
                                else ""
                            )
                        )
                    )

                width = min(
                    50,
                    max(
                        12,
                        max(
                            lengths
                            or [12]
                        ) + 2
                    ),
                )

                worksheet.column_dimensions[
                    letter
                ].width = width

            try:

                reference = (
                    f"A1:"
                    f"{get_column_letter(worksheet.max_column)}"
                    f"{worksheet.max_row}"
                )

                table_name = (
                    "T"
                    + re.sub(
                        r"[^A-Za-z0-9]",
                        "",
                        safe_sheet
                    )[:20]
                )

                table = XLTable(
                    displayName=table_name,
                    ref=reference,
                )

                table.tableStyleInfo = (
                    TableStyleInfo(
                        name="TableStyleMedium2",
                        showRowStripes=True,
                        showColumnStripes=False,
                    )
                )

                worksheet.add_table(
                    table
                )

            except Exception:
                pass

    return buffer.getvalue()


# ============================================================
# PDF EXPORT
# ============================================================

def pdf_style(
    name: str,
    size: float,
    color_hex: str = "#243447",
    bold: bool = False,
    leading: Optional[float] = None,
):

    return ParagraphStyle(
        name=name,
        fontName=(
            "Helvetica-Bold"
            if bold
            else "Helvetica"
        ),
        fontSize=size,
        leading=(
            leading
            or size * 1.35
        ),
        textColor=colors.HexColor(
            color_hex
        ),
        spaceAfter=6,
    )


def pdf_escape(
    value: Any
) -> str:

    import html

    return html.escape(
        str(value)
    ).replace(
        "\n",
        "<br/>"
    )


def pdf_table(
    rows,
    widths=None,
):

    converted = []

    body_style = pdf_style(
        "table_body",
        8.5,
    )

    header_style = pdf_style(
        "table_header",
        8.5,
        "#FFFFFF",
        True,
    )

    for row_index, row in enumerate(
        rows
    ):

        converted.append(
            [
                Paragraph(
                    pdf_escape(value),
                    header_style
                    if row_index == 0
                    else body_style,
                )
                for value in row
            ]
        )

    table = Table(
        converted,
        colWidths=widths,
        repeatRows=1,
        hAlign="LEFT",
    )

    table.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.HexColor(
                        "#16324F"
                    ),
                ),
                (
                    "TEXTCOLOR",
                    (0, 0),
                    (-1, 0),
                    colors.white,
                ),
                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    .35,
                    colors.HexColor(
                        "#D9E2EC"
                    ),
                ),
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "TOP",
                ),
                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    7,
                ),
                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    7,
                ),
                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    6,
                ),
                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    6,
                ),
            ]
        )
    )

    return table


def pdf_footer(
    canvas,
    document
):

    canvas.saveState()

    canvas.setFont(
        "Helvetica",
        8
    )

    canvas.setFillColor(
        colors.HexColor(
            "#718096"
        )
    )

    canvas.drawString(
        34,
        18,
        "AI Data Analyst • "
        "Evidence-first • Local-first"
    )

    canvas.drawRightString(
        A4[0] - 34,
        18,
        f"Page {document.page}"
    )

    canvas.restoreState()


def make_pdf(
    df: pd.DataFrame
) -> bytes:

    if not REPORTLAB_AVAILABLE:

        raise RuntimeError(
            "reportlab is required for PDF export."
        )

    config = get_report_config(
        df
    )

    images = report_chart_images(
        df,
        config
    )

    quality = quality_report(
        df
    )

    kpis = auto_kpis(
        df
    )

    ledger = build_evidence_ledger(
        df
    )

    selected_ids = set(
        config.get(
            "evidence_ids",
            []
        )
    )

    selected_ledger = [
        item
        for item in ledger
        if item["id"] in selected_ids
    ]

    if not selected_ledger:
        selected_ledger = ledger

    title_style = pdf_style(
        "title",
        25,
        "#16324F",
        True,
    )

    h1 = pdf_style(
        "heading1",
        16,
        "#16324F",
        True,
    )

    h2 = pdf_style(
        "heading2",
        11,
        "#2B5D7D",
        True,
    )

    body = pdf_style(
        "body",
        9.5,
    )

    muted = pdf_style(
        "muted",
        8.5,
        "#718096",
    )

    buffer = io.BytesIO()

    document = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=34,
        leftMargin=34,
        topMargin=42,
        bottomMargin=34,
        title=config["title"],
        author="AI Data Analyst",
    )

    story = []

    # --------------------------------------------------------
    # COVER
    # --------------------------------------------------------

    story.extend(
        [
            Spacer(1, 65),

            Paragraph(
                "AI DATA ANALYST",
                pdf_style(
                    "cover_kicker",
                    10,
                    "#2B5D7D",
                    True,
                ),
            ),

            Spacer(1, 8),

            Paragraph(
                pdf_escape(
                    config["title"]
                ),
                title_style,
            ),

            Spacer(1, 10),

            Paragraph(
                pdf_escape(
                    config["subtitle"]
                ),
                body,
            ),

            Spacer(1, 25),

            pdf_table(
                [
                    [
                        "DATASET",
                        "RECORDS",
                        "FIELDS",
                        "QUALITY",
                    ],
                    [
                        current_dataset_name(),
                        f"{len(df):,}",
                        f"{len(df.columns):,}",
                        f"{quality['score']}/100",
                    ],
                ],
                [
                    125,
                    100,
                    100,
                    100,
                ],
            ),

            Spacer(1, 22),

            Paragraph(
                f"Generated {now_text()}",
                muted,
            ),

            Spacer(1, 8),

            Paragraph(
                "Evidence-first • "
                "Verified calculations • "
                "Local AI interpretation",
                muted,
            ),

            PageBreak(),
        ]
    )

    # --------------------------------------------------------
    # EXECUTIVE SUMMARY
    # --------------------------------------------------------

    if (
        "Executive Summary"
        in config["sections"]
    ):

        story.extend(
            [
                Paragraph(
                    "Executive Summary",
                    h1,
                ),
                Spacer(1, 7),
            ]
        )

        metric = config.get(
            "metric"
        )

        summary = (
            f"The analysis covers "
            f"{len(df):,} records across "
            f"{len(df.columns):,} fields. "
            f"Data quality is "
            f"{quality['score']}/100 with "
            f"{quality['missing_cells']:,} "
            f"missing cells and "
            f"{quality['duplicate_rows']:,} "
            f"exact duplicate rows."
        )

        if metric:

            total = pd.to_numeric(
                df[metric],
                errors="coerce"
            ).sum()

            summary += (
                f" The selected primary metric "
                f"is {metric}, with a total of "
                f"{format_value(total)}."
            )

        story.extend(
            [
                Paragraph(
                    pdf_escape(
                        summary
                    ),
                    body,
                ),
                Spacer(1, 12),
            ]
        )

    # --------------------------------------------------------
    # KPI
    # --------------------------------------------------------

    if (
        "KPI Snapshot"
        in config["sections"]
    ):

        story.extend(
            [
                Paragraph(
                    "Key Performance Indicators",
                    h1,
                ),
                Spacer(1, 7),
            ]
        )

        rows = [
            ["KPI", "Value"]
        ]

        for item in kpis:

            rows.append(
                [
                    item["label"],
                    format_value(
                        item["value"]
                    ),
                ]
            )

        story.extend(
            [
                pdf_table(
                    rows,
                    [
                        250,
                        180,
                    ],
                ),
                Spacer(1, 15),
            ]
        )

    # --------------------------------------------------------
    # DASHBOARD CHARTS
    # --------------------------------------------------------

    if (
        "Dashboard Charts"
        in config["sections"]
        and images
    ):

        story.extend(
            [
                Paragraph(
                    "Selected Analytics Visuals",
                    h1,
                ),
                Spacer(1, 7),
            ]
        )

        for chart in config[
            "charts"
        ]:

            if chart not in images:
                continue

            story.append(
                Paragraph(
                    chart.replace(
                        "_",
                        " "
                    ).title(),
                    h2,
                )
            )

            story.append(
                RLImage(
                    io.BytesIO(
                        images[chart]
                    ),
                    width=500,
                    height=250,
                )
            )

            story.append(
                Spacer(1, 12)
            )

    # --------------------------------------------------------
    # FORECAST
    # --------------------------------------------------------

    if (
        "Forecast"
        in config["sections"]
        and st.session_state.last_forecast
    ):

        story.append(
            PageBreak()
        )

        story.extend(
            [
                Paragraph(
                    "Forecast Analysis",
                    h1,
                ),
                Spacer(1, 7),
            ]
        )

        forecast = pd.DataFrame(
            st.session_state.last_forecast
        )

        forecast_metrics = (
            st.session_state
            .forecast_config
            .get(
                "metrics",
                {}
            )
        )

        if (
            "chart"
            in config[
                "forecast_details"
            ]
            and "forecast"
            in images
        ):

            story.append(
                RLImage(
                    io.BytesIO(
                        images[
                            "forecast"
                        ]
                    ),
                    width=500,
                    height=250,
                )
            )

            story.append(
                Spacer(1, 10)
            )

        if (
            "summary"
            in config[
                "forecast_details"
            ]
            or "metrics"
            in config[
                "forecast_details"
            ]
        ):

            metric = config.get(
                "metric",
                "Metric"
            )

            metrics_rows = [
                [
                    "Measure",
                    "Value",
                ],
                [
                    "Method",
                    forecast_metrics.get(
                        "method",
                        "—"
                    ),
                ],
                [
                    "Metric",
                    metric,
                ],
                [
                    "Historical periods",
                    forecast_metrics.get(
                        "history_points",
                        "—"
                    ),
                ],
                [
                    "Forecast horizon",
                    forecast_metrics.get(
                        "forecast_points",
                        "—"
                    ),
                ],
                [
                    "MAE",
                    format_value(
                        forecast_metrics.get(
                            "mae"
                        )
                    ),
                ],
                [
                    "RMSE",
                    format_value(
                        forecast_metrics.get(
                            "rmse"
                        )
                    ),
                ],
                [
                    "MAPE",
                    (
                        f"{forecast_metrics.get('mape', 0):.2f}%"
                    ),
                ],
            ]

            story.extend(
                [
                    pdf_table(
                        metrics_rows,
                        [
                            230,
                            200,
                        ],
                    ),
                    Spacer(1, 12),
                ]
            )

        if (
            "table"
            in config[
                "forecast_details"
            ]
        ):

            rows = [
                [
                    "Date",
                    "Type",
                    "Value",
                ]
            ]

            for row in forecast.itertuples(
                index=False
            ):

                rows.append(
                    [
                        str(row.Date),
                        row.Type,
                        format_value(
                            row.Value
                        ),
                    ]
                )

            story.append(
                pdf_table(
                    rows,
                    [
                        150,
                        120,
                        160,
                    ],
                )
            )

    # --------------------------------------------------------
    # WHAT IF
    # --------------------------------------------------------

    if (
        "What-If"
        in config["sections"]
        and st.session_state.last_forecast
    ):

        forecast = pd.DataFrame(
            st.session_state.last_forecast
        )

        forecast_only = forecast[
            forecast["Type"]
            == "Forecast"
        ]

        if not forecast_only.empty:

            change = float(
                config.get(
                    "what_if_change",
                    10
                )
            )

            baseline = float(
                forecast_only[
                    "Value"
                ].sum()
            )

            scenario = (
                baseline
                * (
                    1
                    + change / 100
                )
            )

            story.extend(
                [
                    Spacer(1, 15),
                    Paragraph(
                        "What-If Scenario",
                        h1,
                    ),
                    Spacer(1, 7),
                ]
            )

            if "what_if" in images:

                story.append(
                    RLImage(
                        io.BytesIO(
                            images[
                                "what_if"
                            ]
                        ),
                        width=500,
                        height=250,
                    )
                )

                story.append(
                    Spacer(1, 10)
                )

            story.append(
                pdf_table(
                    [
                        [
                            "Metric",
                            "Baseline",
                            "Scenario",
                            "Delta",
                        ],
                        [
                            config.get(
                                "metric",
                                "Metric"
                            ),
                            format_value(
                                baseline
                            ),
                            format_value(
                                scenario
                            ),
                            format_value(
                                scenario
                                - baseline
                            ),
                        ],
                    ],
                    [
                        145,
                        110,
                        110,
                        110,
                    ],
                )
            )

    # --------------------------------------------------------
    # EVIDENCE
    # --------------------------------------------------------

    if (
        "Evidence"
        in config["sections"]
    ):

        story.append(
            PageBreak()
        )

        story.extend(
            [
                Paragraph(
                    "Evidence Ledger",
                    h1,
                ),
                Spacer(1, 7),
            ]
        )

        rows = [
            [
                "ID",
                "Claim",
                "Value",
                "Method",
            ]
        ]

        for item in selected_ledger:

            rows.append(
                [
                    item["id"],
                    item["claim"],
                    item["value"],
                    item["method"],
                ]
            )

        story.append(
            pdf_table(
                rows,
                [
                    45,
                    180,
                    95,
                    145,
                ],
            )
        )

    # --------------------------------------------------------
    # VERIFICATION
    # --------------------------------------------------------

    if (
        "Verification"
        in config["sections"]
        and st.session_state.verification_results
    ):

        story.extend(
            [
                Spacer(1, 18),
                Paragraph(
                    "Verification",
                    h1,
                ),
                Spacer(1, 7),
            ]
        )

        rows = [
            [
                "Finding",
                "Status",
                "Summary",
            ]
        ]

        for item in (
            st.session_state
            .verification_results
        ):

            rows.append(
                [
                    item.get(
                        "finding_id",
                        ""
                    ),
                    item.get(
                        "status",
                        ""
                    ),
                    item.get(
                        "summary",
                        ""
                    ),
                ]
            )

        story.append(
            pdf_table(
                rows,
                [
                    65,
                    90,
                    310,
                ],
            )
        )

    # --------------------------------------------------------
    # CHALLENGE
    # --------------------------------------------------------

    if (
        "Challenge"
        in config["sections"]
        and st.session_state.challenge_results
    ):

        story.extend(
            [
                Spacer(1, 18),
                Paragraph(
                    "Challenge & Alternative Explanations",
                    h1,
                ),
                Spacer(1, 7),
            ]
        )

        rows = [
            [
                "Finding",
                "Status",
                "Challenge result",
            ]
        ]

        for item in (
            st.session_state
            .challenge_results
        ):

            rows.append(
                [
                    item.get(
                        "finding_id",
                        ""
                    ),
                    item.get(
                        "status",
                        ""
                    ),
                    item.get(
                        "summary",
                        ""
                    ),
                ]
            )

        story.append(
            pdf_table(
                rows,
                [
                    65,
                    120,
                    280,
                ],
            )
        )

    # --------------------------------------------------------
    # DATA QUALITY
    # --------------------------------------------------------

    if (
        "Data Quality"
        in config["sections"]
    ):

        story.extend(
            [
                Spacer(1, 18),
                Paragraph(
                    "Data Quality",
                    h1,
                ),
                Spacer(1, 7),
                pdf_table(
                    [
                        [
                            "Measure",
                            "Value",
                        ],
                        [
                            "Rows",
                            f"{quality['rows']:,}",
                        ],
                        [
                            "Columns",
                            f"{quality['columns']:,}",
                        ],
                        [
                            "Missing cells",
                            f"{quality['missing_cells']:,}",
                        ],
                        [
                            "Duplicate rows",
                            f"{quality['duplicate_rows']:,}",
                        ],
                        [
                            "Quality score",
                            f"{quality['score']}/100",
                        ],
                    ],
                    [
                        250,
                        180,
                    ],
                ),
            ]
        )

    # --------------------------------------------------------
    # AI
    # --------------------------------------------------------

    if (
        "AI Conversation"
        in config["sections"]
        and st.session_state.chat_history
    ):

        story.append(
            PageBreak()
        )

        story.extend(
            [
                Paragraph(
                    "AI Analyst Conversation",
                    h1,
                ),
                Spacer(1, 7),
            ]
        )

        for message in (
            st.session_state.chat_history
        ):

            role = (
                "User"
                if message.get(
                    "role"
                ) == "user"
                else "Phi-3"
            )

            story.extend(
                [
                    Paragraph(
                        role,
                        h2,
                    ),
                    Paragraph(
                        pdf_escape(
                            message.get(
                                "content",
                                ""
                            )
                        ),
                        body,
                    ),
                    Spacer(1, 7),
                ]
            )

    # --------------------------------------------------------
    # AUDIT / PROVENANCE
    # --------------------------------------------------------

    if (
        "Audit / Provenance"
        in config["sections"]
    ):

        story.append(
            PageBreak()
        )

        story.extend(
            [
                Paragraph(
                    "Audit & Provenance",
                    h1,
                ),
                Spacer(1, 8),
                Paragraph(
                    pdf_escape(
                        f"""
Dataset: {current_dataset_name()}

Rows × fields:
{len(df):,} × {len(df.columns):,}

Ollama endpoint:
{OLLAMA_URL}

Model:
{st.session_state.selected_model}

Numerical evidence:
Generated deterministically

Uploaded Python / Java:
Never executed

SQL:
Read-only

Project persistence:
CSV + JSON

Report generation:
Selected through Report Studio
"""
                    ),
                    body,
                ),
            ]
        )

        if st.session_state.audit_log:

            story.extend(
                [
                    Spacer(1, 12),
                    pdf_table(
                        [
                            [
                                "Time",
                                "Action",
                                "Details",
                            ]
                        ]
                        + [
                            [
                                item.get(
                                    "time",
                                    ""
                                ),
                                item.get(
                                    "action",
                                    ""
                                ),
                                item.get(
                                    "details",
                                    ""
                                ),
                            ]
                            for item in
                            st.session_state
                            .audit_log[-80:]
                        ],
                        [
                            95,
                            130,
                            205,
                        ],
                    ),
                ]
            )

    document.build(
        story,
        onFirstPage=pdf_footer,
        onLaterPages=pdf_footer,
    )

    return buffer.getvalue()


# ============================================================
# POWERPOINT EXPORT
# ============================================================

def make_pptx(
    df: pd.DataFrame
) -> bytes:

    if not PPTX_AVAILABLE:

        raise RuntimeError(
            "python-pptx is required for PowerPoint export."
        )

    config = get_report_config(
        df
    )

    images = report_chart_images(
        df,
        config
    )

    quality = quality_report(
        df
    )

    kpis = auto_kpis(
        df
    )

    presentation = Presentation()

    presentation.slide_width = Inches(
        13.333
    )

    presentation.slide_height = Inches(
        7.5
    )

    NAVY = RGBColor(
        22,
        50,
        79
    )

    BLUE = RGBColor(
        43,
        93,
        125
    )

    DARK = RGBColor(
        36,
        52,
        71
    )

    LIGHT = RGBColor(
        242,
        246,
        249
    )

    WHITE = RGBColor(
        255,
        255,
        255
    )

    BORDER = RGBColor(
        220,
        228,
        235
    )

    def background(
        slide
    ):

        slide.background.fill.solid()

        slide.background.fill.fore_color.rgb = WHITE

    def text(
        slide,
        x,
        y,
        width,
        height,
        value,
        size=14,
        bold=False,
        color=DARK,
    ):

        shape = slide.shapes.add_textbox(
            Inches(x),
            Inches(y),
            Inches(width),
            Inches(height),
        )

        frame = shape.text_frame

        frame.word_wrap = True

        paragraph = (
            frame.paragraphs[0]
        )

        paragraph.text = str(
            value
        )

        paragraph.font.size = Pt(
            size
        )

        paragraph.font.bold = bold

        paragraph.font.color.rgb = (
            color
        )

        return shape

    def slide_title(
        slide,
        title,
        subtitle=None,
    ):

        text(
            slide,
            .65,
            .35,
            12,
            .6,
            title,
            27,
            True,
            NAVY,
        )

        if subtitle:

            text(
                slide,
                .67,
                1.02,
                12,
                .35,
                subtitle,
                10,
                False,
                BLUE,
            )

    def card(
        slide,
        x,
        y,
        width,
        height,
        label,
        value,
    ):

        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(x),
            Inches(y),
            Inches(width),
            Inches(height),
        )

        shape.fill.solid()

        shape.fill.fore_color.rgb = (
            LIGHT
        )

        shape.line.color.rgb = (
            BORDER
        )

        text(
            slide,
            x + .18,
            y + .13,
            width - .35,
            .3,
            label,
            9,
            False,
            BLUE,
        )

        text(
            slide,
            x + .18,
            y + .48,
            width - .35,
            .5,
            value,
            21,
            True,
            NAVY,
        )

    # --------------------------------------------------------
    # COVER
    # --------------------------------------------------------

    slide = presentation.slides.add_slide(
        presentation.slide_layouts[6]
    )

    background(slide)

    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        0,
        0,
        presentation.slide_width,
        Inches(.15),
    )

    bar.fill.solid()
    bar.fill.fore_color.rgb = BLUE
    bar.line.fill.background()

    text(
        slide,
        .75,
        1.35,
        11.5,
        .5,
        "AI DATA ANALYST",
        14,
        True,
        BLUE,
    )

    text(
        slide,
        .75,
        2.0,
        11.5,
        1.2,
        config["title"],
        30,
        True,
        NAVY,
    )

    text(
        slide,
        .78,
        3.45,
        11,
        .7,
        config["subtitle"],
        15,
        False,
        DARK,
    )

    text(
        slide,
        .78,
        6.45,
        11,
        .35,
        f"{current_dataset_name()} • "
        f"Generated {now_text()} • "
        "Local-first / Ollama",
        10,
        False,
        BLUE,
    )

    # --------------------------------------------------------
    # EXECUTIVE SNAPSHOT
    # --------------------------------------------------------

    if (
        "Executive Summary"
        in config["sections"]
        or "KPI Snapshot"
        in config["sections"]
    ):

        slide = presentation.slides.add_slide(
            presentation.slide_layouts[6]
        )

        background(slide)

        slide_title(
            slide,
            "Executive Snapshot",
            current_dataset_name(),
        )

        snapshot = [
            (
                "Records",
                f"{len(df):,}"
            ),
            (
                "Fields",
                f"{len(df.columns):,}"
            ),
            (
                "Quality",
                f"{quality['score']}/100"
            ),
            (
                "Missing",
                f"{quality['missing_cells']:,}"
            ),
        ]

        for index, (
            label,
            value
        ) in enumerate(
            snapshot
        ):

            card(
                slide,
                .65 + index * 3.1,
                1.55,
                2.8,
                1.25,
                label,
                value,
            )

        text(
            slide,
            .75,
            3.25,
            11.7,
            2.6,
            "Evidence-first workflow:\n\n"
            "Deterministic calculations establish "
            "measurable facts.\n\n"
            "Local Phi-3 interprets the evidence.\n\n"
            "Verification recalculates claims.\n\n"
            "Challenge searches for alternative explanations.",
            16,
            False,
            DARK,
        )

    # --------------------------------------------------------
    # KPI SLIDE
    # --------------------------------------------------------

    if (
        "KPI Snapshot"
        in config["sections"]
    ):

        slide = presentation.slides.add_slide(
            presentation.slide_layouts[6]
        )

        background(slide)

        slide_title(
            slide,
            "Key Performance Indicators",
        )

        for index, item in enumerate(
            kpis[:6]
        ):

            card(
                slide,
                .65
                + (index % 3) * 4.15,
                1.45
                + (index // 3) * 2.0,
                3.75,
                1.45,
                item["label"],
                format_value(
                    item["value"]
                ),
            )

    # --------------------------------------------------------
    # CHART SLIDES
    # --------------------------------------------------------

    if (
        "Dashboard Charts"
        in config["sections"]
    ):

        for chart in config[
            "charts"
        ]:

            if chart not in images:
                continue

            slide = presentation.slides.add_slide(
                presentation.slide_layouts[6]
            )

            background(slide)

            slide_title(
                slide,
                chart.replace(
                    "_",
                    " "
                ).title(),
                "Selected from Report Studio",
            )

            slide.shapes.add_picture(
                io.BytesIO(
                    images[chart]
                ),
                Inches(.65),
                Inches(1.35),
                width=Inches(12),
                height=Inches(5.65),
            )

    # --------------------------------------------------------
    # FORECAST
    # --------------------------------------------------------

    if (
        "Forecast"
        in config["sections"]
        and st.session_state.last_forecast
    ):

        slide = presentation.slides.add_slide(
            presentation.slide_layouts[6]
        )

        background(slide)

        slide_title(
            slide,
            "Forecast & Scenario Analysis",
        )

        if "forecast" in images:

            slide.shapes.add_picture(
                io.BytesIO(
                    images[
                        "forecast"
                    ]
                ),
                Inches(.55),
                Inches(1.25),
                width=Inches(8.2),
                height=Inches(4.9),
            )

        metrics = (
            st.session_state
            .forecast_config
            .get(
                "metrics",
                {}
            )
        )

        forecast_text = (
            f"Method\n"
            f"{metrics.get('method', '—')}\n\n"
            f"Historical periods\n"
            f"{metrics.get('history_points', '—')}\n\n"
            f"Forecast horizon\n"
            f"{metrics.get('forecast_points', '—')}\n\n"
            f"MAE\n"
            f"{format_value(metrics.get('mae'))}\n\n"
            f"RMSE\n"
            f"{format_value(metrics.get('rmse'))}\n\n"
            f"MAPE\n"
            f"{metrics.get('mape', 0):.2f}%"
        )

        text(
            slide,
            9.0,
            1.45,
            3.5,
            4.9,
            forecast_text,
            14,
            False,
            DARK,
        )

    # --------------------------------------------------------
    # EVIDENCE
    # --------------------------------------------------------

    if (
        "Evidence"
        in config["sections"]
    ):

        ledger = build_evidence_ledger(
            df
        )

        selected = set(
            config.get(
                "evidence_ids",
                []
            )
        )

        ledger = [
            item
            for item in ledger
            if item["id"] in selected
        ] or ledger

        for start in range(
            0,
            len(ledger),
            7
        ):

            batch = ledger[
                start:start + 7
            ]

            slide = presentation.slides.add_slide(
                presentation.slide_layouts[6]
            )

            background(slide)

            slide_title(
                slide,
                "Evidence Ledger",
                f"Evidence {start+1}–"
                f"{min(start+7, len(ledger))}",
            )

            table_shape = (
                slide.shapes.add_table(
                    len(batch) + 1,
                    4,
                    Inches(.55),
                    Inches(1.35),
                    Inches(12.2),
                    Inches(5.3),
                )
            )

            table = table_shape.table

            headers = [
                "ID",
                "Claim",
                "Value",
                "Method",
            ]

            for column_index, header in enumerate(
                headers
            ):

                table.cell(
                    0,
                    column_index
                ).text = header

            for row_index, item in enumerate(
                batch,
                1
            ):

                values = [
                    item.get(
                        "id",
                        ""
                    ),
                    item.get(
                        "claim",
                        ""
                    ),
                    item.get(
                        "value",
                        ""
                    ),
                    item.get(
                        "method",
                        ""
                    ),
                ]

                for column_index, value in enumerate(
                    values
                ):

                    table.cell(
                        row_index,
                        column_index
                    ).text = str(
                        value
                    )

            for row in table.rows:

                for cell in row.cells:

                    for paragraph in (
                        cell.text_frame
                        .paragraphs
                    ):

                        paragraph.font.size = Pt(
                            8.5
                        )

                        paragraph.font.color.rgb = (
                            DARK
                        )

    # --------------------------------------------------------
    # VERIFICATION
    # --------------------------------------------------------

    if (
        "Verification"
        in config["sections"]
        and st.session_state.verification_results
    ):

        slide = presentation.slides.add_slide(
            presentation.slide_layouts[6]
        )

        background(slide)

        slide_title(
            slide,
            "Verification",
            "Independent deterministic recalculation",
        )

        verification_text = "\n\n".join(
            [
                f"{item.get('status', '')} — "
                f"{item.get('finding_id', '')}\n"
                f"{item.get('summary', '')}"
                for item in
                st.session_state
                .verification_results
            ]
        )

        text(
            slide,
            .7,
            1.45,
            11.9,
            5.2,
            verification_text,
            12,
            False,
            DARK,
        )

    # --------------------------------------------------------
    # CHALLENGE
    # --------------------------------------------------------

    if (
        "Challenge"
        in config["sections"]
        and st.session_state.challenge_results
    ):

        slide = presentation.slides.add_slide(
            presentation.slide_layouts[6]
        )

        background(slide)

        slide_title(
            slide,
            "Challenge & Alternative Explanations",
        )

        challenge_text = "\n\n".join(
            [
                f"{item.get('status', '')} — "
                f"{item.get('finding_id', '')}\n"
                f"{item.get('summary', '')}"
                for item in
                st.session_state
                .challenge_results
            ]
        )

        text(
            slide,
            .7,
            1.45,
            11.9,
            5.2,
            challenge_text,
            12,
            False,
            DARK,
        )

    # --------------------------------------------------------
    # AI TRANSCRIPT
    # --------------------------------------------------------

    if (
        "AI Conversation"
        in config["sections"]
        and st.session_state.chat_history
    ):

        slide = presentation.slides.add_slide(
            presentation.slide_layouts[6]
        )

        background(slide)

        slide_title(
            slide,
            "AI Analyst Conversation",
            "Local Phi-3 transcript",
        )

        transcript = "\n\n".join(
            [
                (
                    "USER"
                    if item.get(
                        "role"
                    ) == "user"
                    else "PHI-3"
                )
                + "\n"
                + str(
                    item.get(
                        "content",
                        ""
                    )
                )
                for item in
                st.session_state
                .chat_history[-8:]
            ]
        )

        text(
            slide,
            .7,
            1.35,
            11.9,
            5.7,
            transcript,
            11,
            False,
            DARK,
        )

    # --------------------------------------------------------
    # PROVENANCE
    # --------------------------------------------------------

    if (
        "Audit / Provenance"
        in config["sections"]
    ):

        slide = presentation.slides.add_slide(
            presentation.slide_layouts[6]
        )

        background(slide)

        slide_title(
            slide,
            "Provenance & Controls",
        )

        provenance = (
            f"Dataset: {current_dataset_name()}\n\n"
            f"Rows × fields: "
            f"{len(df):,} × "
            f"{len(df.columns):,}\n\n"
            f"Model: "
            f"{st.session_state.selected_model}\n\n"
            f"Ollama: {OLLAMA_URL}\n\n"
            "Numerical evidence: deterministic\n\n"
            "Python / Java uploads: never executed\n\n"
            "SQL workspace: read-only\n\n"
            "Persistence: CSV + JSON\n\n"
            "Report contents: selected in Report Studio"
        )

        text(
            slide,
            .75,
            1.45,
            11.8,
            5.3,
            provenance,
            14,
            False,
            DARK,
        )

    buffer = io.BytesIO()

    presentation.save(
        buffer
    )

    return buffer.getvalue()


# ============================================================
# DASHBOARD FILTERS
# ============================================================

def apply_dashboard_filters(
    df: pd.DataFrame,
    filters: Dict[str, Any],
) -> pd.DataFrame:

    result = df.copy()

    for column, values in (
        filters or {}
    ).items():

        if (
            column in result.columns
            and values
        ):

            result = result[
                result[column]
                .astype(str)
                .isin(
                    [
                        str(v)
                        for v in values
                    ]
                )
            ]

    return result


# ============================================================
# UI HELPERS
# ============================================================

def metric_card(
    label,
    value,
    sub=""
):

    st.markdown(
        f"""
<div class="kpi">
<div class="label">{label}</div>
<div class="value">{value}</div>
<div class="sub">{sub}</div>
</div>
""",
        unsafe_allow_html=True,
    )


def render_findings(
    findings
):

    if not findings:

        st.info(
            "No material findings detected."
        )

        return

    for finding in findings:

        st.markdown(
            f"""
<div class="finding">
<strong>{finding.get("title", "Finding")}</strong>
<br>
<span class="small">
{finding.get("text", "")}
</span>
</div>
""",
            unsafe_allow_html=True,
        )


def require_df():

    df = current_df()

    if (
        df is None
        or df.empty
    ):

        st.info(
            "Import a dataset or load the Portfolio Demo."
        )

        return None

    return df


def render_header():

    online = st.session_state.ollama_status

    badge = (
        '<span class="badge good">'
        '● Ollama online'
        '</span>'
        if online
        else
        '<span class="badge warn">'
        '● Ollama offline'
        '</span>'
    )

    st.markdown(
        f"""
<div class="hero">
<h1>📊 AI Data Analyst</h1>
<p>
Turn messy data into verified findings.
Local-first analytics, evidence,
verification, challenge and executive reporting.
</p>
<div style="margin-top:12px">
{badge}
<span class="badge">
{current_dataset_name()}
</span>
<span class="badge">
v{APP_VERSION}
</span>
</div>
</div>
""",
        unsafe_allow_html=True,
    )


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.markdown(
        "## 📊 AI DATA ANALYST"
    )

    st.caption(
        "Evidence-first • Local-first • Ollama"
    )

    pages = [
        "Executive Overview",
        "Data Sources",
        "Data Quality & Cleaning",
        "SQL Analyst",
        "Python Analysis",
        "Auto Dashboard",
        "Explore & Statistics",
        "Forecast & What-If",
        "AI Data Analyst",
        "Evidence & Verification",
        "Reports & Exports",
        "Settings",
    ]

    st.session_state.page = st.radio(
        "Workspace",
        pages,
        index=(
            pages.index(
                st.session_state.page
            )
            if st.session_state.page in pages
            else 0
        ),
    )

    st.divider()

    df = current_df()

    if df is not None:

        st.markdown(
            f"**Active dataset**  \n"
            f"{current_dataset_name()}"
        )

        st.caption(
            f"{len(df):,} rows × "
            f"{len(df.columns):,} columns"
        )

        c1, c2 = st.columns(2)

        with c1:

            if st.button(
                "↶ Undo",
                disabled=not st.session_state.undo_stack,
                use_container_width=True,
            ):

                if undo_change():
                    st.rerun()

        with c2:

            if st.button(
                "↷ Redo",
                disabled=not st.session_state.redo_stack,
                use_container_width=True,
            ):

                if redo_change():
                    st.rerun()

    st.divider()

    online, status = ollama_health_check()

    st.session_state.ollama_status = online

    st.markdown(
        f"**Ollama:** "
        f"{'🟢 Online' if online else '🔴 Offline'}"
    )

    if online:

        models = get_ollama_models()

        st.session_state.ollama_models = models

        if models:

            current_model = (
                st.session_state.selected_model
            )

            if current_model not in models:

                current_model = models[0]

            st.session_state.selected_model = (
                st.selectbox(
                    "Model",
                    models,
                    index=models.index(
                        current_model
                    ),
                )
            )

    else:

        st.caption(
            "Deterministic analytics remain available."
        )


# ============================================================
# HEADER
# ============================================================

render_header()


# ============================================================
# EXECUTIVE OVERVIEW
# ============================================================

if st.session_state.page == "Executive Overview":

    df = require_df()

    if df is not None:

        quality = quality_report(
            df
        )

        columns = st.columns(5)

        metrics = [
            (
                "Rows",
                f"{len(df):,}",
                "records",
            ),
            (
                "Columns",
                f"{len(df.columns):,}",
                "fields",
            ),
            (
                "Missing",
                f"{quality['missing_cells']:,}",
                f"{quality['missing_rate']*100:.2f}% of cells",
            ),
            (
                "Duplicates",
                f"{quality['duplicate_rows']:,}",
                f"{quality['duplicate_rate']*100:.2f}% of rows",
            ),
            (
                "Health",
                f"{quality['score']}/100",
                "quality score",
            ),
        ]

        for column, values in zip(
            columns,
            metrics
        ):

            with column:

                metric_card(
                    *values
                )

        st.subheader(
            "🔎 Data Detective"
        )

        render_findings(
            detective_findings(df)
        )

        st.subheader(
            "📈 Automatic Business Insights"
        )

        kpis = auto_kpis(
            df
        )

        if kpis:

            st.dataframe(
                pd.DataFrame(kpis),
                use_container_width=True,
                hide_index=True,
            )

        st.subheader(
            "Schema"
        )

        st.dataframe(
            schema_report(df),
            use_container_width=True,
            hide_index=True,
        )


# ============================================================
# DATA SOURCES
# ============================================================

elif st.session_state.page == "Data Sources":

    st.subheader(
        "📂 Import Data"
    )

    st.caption(
        "Python and Java files are statically inspected. "
        "Uploaded source code is never executed."
    )

    upload = st.file_uploader(
        "Upload CSV, Excel, JSON, Parquet, TSV/TXT, Python or Java",
        type=[
            "csv",
            "xls",
            "xlsx",
            "xlsm",
            "json",
            "parquet",
            "tsv",
            "txt",
            "py",
            "java",
        ],
    )

    if upload:

        if st.button(
            "Import Dataset",
            type="primary",
            use_container_width=True,
        ):

            try:

                data, source = read_uploaded_file(
                    upload
                )

                register_dataset(
                    data,
                    upload.name,
                    source,
                    True,
                )

                st.success(
                    f"Imported {len(data):,} "
                    f"rows × {len(data.columns):,} columns."
                )

                st.rerun()

            except Exception as exc:

                st.error(
                    str(exc)
                )

    st.divider()

    c1, c2 = st.columns(2)

    with c1:

        if st.button(
            "✨ Load Portfolio Demo",
            use_container_width=True,
        ):

            register_dataset(
                build_demo_dataset(),
                "demo_sales.csv",
                "portfolio demo",
                True,
            )

            st.session_state.demo_loaded = True

            st.success(
                "Portfolio demo loaded."
            )

            st.rerun()

    with c2:

        st.info(
            "Supported: CSV • XLS/XLSX • JSON • "
            "Parquet • TSV • TXT • Python • Java"
        )

    if current_df() is not None:

        st.divider()

        st.subheader(
            "Current Dataset"
        )

        st.dataframe(
            arrow_safe_df(
                current_df().head(500)
            ),
            use_container_width=True,
            hide_index=True,
        )


# ============================================================
# DATA QUALITY & CLEANING
# ============================================================

elif st.session_state.page == "Data Quality & Cleaning":

    df = require_df()

    if df is not None:

        st.subheader(
            "🧹 Data Quality & Cleaning"
        )

        quality = quality_report(
            df
        )

        c1, c2, c3, c4 = st.columns(4)

        with c1:
            metric_card(
                "Quality",
                f"{quality['score']}/100",
                "deterministic",
            )

        with c2:
            metric_card(
                "Missing",
                f"{quality['missing_cells']:,}",
                "cells",
            )

        with c3:
            metric_card(
                "Duplicates",
                f"{quality['duplicate_rows']:,}",
                "rows",
            )

        with c4:
            metric_card(
                "Rows",
                f"{len(df):,}",
                "current",
            )

        tabs = st.tabs(
            [
                "Overview",
                "Missing",
                "Duplicates",
                "Outliers",
                "Transform Studio",
                "Audit",
            ]
        )

        with tabs[0]:

            st.dataframe(
                schema_report(df),
                use_container_width=True,
                hide_index=True,
            )

        with tabs[1]:

            missing = missing_report(
                df
            )

            st.dataframe(
                missing,
                use_container_width=True,
                hide_index=True,
            )

            if not missing.empty:

                column = st.selectbox(
                    "Column",
                    missing[
                        "Column"
                    ].tolist(),
                )

                strategy = st.selectbox(
                    "Fill strategy",
                    [
                        "Median",
                        "Mean",
                        "Zero",
                        "Unknown",
                    ],
                )

                if st.button(
                    "Apply missing-value fix",
                    type="primary",
                ):

                    out = df.copy()

                    if strategy == "Median":

                        value = pd.to_numeric(
                            out[column],
                            errors="coerce"
                        ).median()

                    elif strategy == "Mean":

                        value = pd.to_numeric(
                            out[column],
                            errors="coerce"
                        ).mean()

                    elif strategy == "Zero":

                        value = 0

                    else:

                        value = "Unknown"

                    out[column] = (
                        out[column]
                        .fillna(value)
                    )

                    commit_dataframe(
                        out,
                        "Fill missing values",
                        f"{column} using {strategy}",
                    )

                    st.success(
                        "Missing values updated."
                    )

                    st.rerun()

        with tabs[2]:

            duplicates = int(
                df.duplicated().sum()
            )

            st.metric(
                "Exact duplicate rows",
                duplicates
            )

            if duplicates:

                if st.button(
                    "Remove duplicate rows",
                    type="primary",
                ):

                    out = (
                        df
                        .drop_duplicates()
                        .reset_index(
                            drop=True
                        )
                    )

                    commit_dataframe(
                        out,
                        "Remove duplicates",
                        f"Removed {duplicates:,} rows",
                    )

                    st.success(
                        "Duplicates removed."
                    )

                    st.rerun()

        with tabs[3]:

            st.dataframe(
                outlier_report(df),
                use_container_width=True,
                hide_index=True,
            )

        with tabs[4]:

            st.markdown(
                "### Power Query-style Transform Studio"
            )

            st.caption(
                "This is a local transformation layer inspired by "
                "Power Query workflows. It does not execute Excel M code."
            )

            operation = st.selectbox(
                "Transformation",
                [
                    "Trim whitespace",
                    "Lowercase",
                    "Uppercase",
                    "Title Case",
                    "Fill missing",
                    "Convert numeric",
                    "Convert date",
                    "Replace value",
                ],
            )

            column = st.selectbox(
                "Column",
                list(df.columns)
            )

            parameter = ""

            if operation == "Fill missing":

                parameter = st.selectbox(
                    "Strategy",
                    [
                        "Median",
                        "Mean",
                        "Zero",
                        "Unknown",
                    ],
                )

            elif operation == "Replace value":

                parameter = st.text_input(
                    "Replacement",
                    "Old => New",
                )

            if st.button(
                "Preview transformation"
            ):

                try:

                    preview = apply_transform(
                        df,
                        column,
                        operation,
                        parameter,
                    )

                    st.dataframe(
                        arrow_safe_df(
                            preview.head(100)
                        ),
                        use_container_width=True,
                        hide_index=True,
                    )

                except Exception as exc:

                    st.error(
                        str(exc)
                    )

            if st.button(
                "Apply transformation",
                type="primary",
            ):

                try:

                    transformed = apply_transform(
                        df,
                        column,
                        operation,
                        parameter,
                    )

                    commit_dataframe(
                        transformed,
                        "Transform",
                        f"{operation} on {column}",
                    )

                    st.success(
                        "Transformation applied."
                    )

                    st.rerun()

                except Exception as exc:

                    st.error(
                        str(exc)
                    )

            if st.session_state.cleaning_history:

                st.markdown(
                    "### Transformation history"
                )

                st.dataframe(
                    pd.DataFrame(
                        st.session_state
                        .cleaning_history
                    ),
                    use_container_width=True,
                    hide_index=True,
                )

        with tabs[5]:

            st.dataframe(
                pd.DataFrame(
                    st.session_state.audit_log
                ),
                use_container_width=True,
                hide_index=True,
            )

        st.divider()

        st.download_button(
            "⬇️ Download Clean CSV",
            df.to_csv(
                index=False
            ).encode(
                "utf-8-sig"
            ),
            file_name=(
                f"{safe_filename(current_dataset_name())}"
                "_cleaned.csv"
            ),
            mime="text/csv",
            use_container_width=True,
        )

        st.download_button(
            "⬇️ Download Clean Excel",
            dataframe_to_excel(df),
            file_name=(
                f"{safe_filename(current_dataset_name())}"
                "_cleaned_analysis.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            ),
            use_container_width=True,
        )


# ============================================================
# SQL ANALYST
# ============================================================

elif st.session_state.page == "SQL Analyst":

    df = require_df()

    if df is not None:

        st.subheader(
            "🗄️ SQL Analyst"
        )

        st.caption(
            "Read-only DuckDB workspace. "
            "The active dataframe is available as table `data`."
        )

        query = st.text_area(
            "SQL query",
            value=(
                "SELECT * FROM data LIMIT 20"
            ),
            height=150,
        )

        if st.button(
            "Run SQL",
            type="primary",
        ):

            started = time.time()

            try:

                result = execute_sql(
                    df,
                    query
                )

                elapsed = (
                    time.time()
                    - started
                )

                st.session_state.sql_history.append(
                    {
                        "time":
                            now_text(),
                        "query":
                            query,
                        "rows":
                            len(result),
                        "milliseconds":
                            round(
                                elapsed
                                * 1000,
                                2
                            ),
                    }
                )

                st.dataframe(
                    arrow_safe_df(
                        result
                    ),
                    use_container_width=True,
                    hide_index=True,
                )

                st.caption(
                    f"{len(result):,} rows • "
                    f"{elapsed*1000:.1f} ms"
                )

                st.download_button(
                    "Download SQL Result",
                    result.to_csv(
                        index=False
                    ).encode(
                        "utf-8-sig"
                    ),
                    file_name="sql_result.csv",
                    mime="text/csv",
                )

            except Exception as exc:

                st.error(
                    str(exc)
                )

        if st.session_state.sql_history:

            st.subheader(
                "SQL History"
            )

            st.dataframe(
                pd.DataFrame(
                    st.session_state.sql_history
                ),
                use_container_width=True,
                hide_index=True,
            )


# ============================================================
# PYTHON ANALYSIS
# ============================================================

elif st.session_state.page == "Python Analysis":

    df = require_df()

    if df is not None:

        st.subheader(
            "🐍 Safe Python Analysis"
        )

        st.caption(
            "Built-in pandas analytics only. "
            "Uploaded Python files are never executed."
        )

        operation = st.selectbox(
            "Analysis",
            [
                "Group summary",
                "Top/bottom records",
                "Correlation with a metric",
                "Missingness profile",
                "Date trend summary",
            ],
        )

        nums = numeric_columns(
            df
        )

        cats = categorical_columns(
            df
        )

        dates = detect_datetime_candidates(
            df
        )

        if (
            operation == "Group summary"
            and nums
            and cats
        ):

            metric = st.selectbox(
                "Metric",
                nums
            )

            dimension = st.selectbox(
                "Dimension",
                cats
            )

            if st.button(
                "Run analysis",
                type="primary",
            ):

                result = (
                    df.groupby(
                        dimension,
                        dropna=False
                    )[metric]
                    .agg(
                        [
                            "count",
                            "sum",
                            "mean",
                            "median",
                        ]
                    )
                    .reset_index()
                    .sort_values(
                        "sum",
                        ascending=False
                    )
                )

                st.dataframe(
                    result,
                    use_container_width=True,
                    hide_index=True,
                )

        elif (
            operation
            == "Top/bottom records"
            and nums
        ):

            metric = st.selectbox(
                "Metric",
                nums
            )

            number = st.slider(
                "Rows",
                5,
                50,
                10,
            )

            result = (
                df.sort_values(
                    metric,
                    ascending=False
                )
                .head(number)
            )

            st.dataframe(
                arrow_safe_df(result),
                use_container_width=True,
                hide_index=True,
            )

        elif (
            operation
            == "Correlation with a metric"
            and len(nums) >= 2
        ):

            metric = st.selectbox(
                "Reference metric",
                nums
            )

            result = (
                df[nums]
                .corr()[[metric]]
                .sort_values(
                    metric,
                    ascending=False
                )
                .reset_index()
            )

            st.dataframe(
                result,
                use_container_width=True,
                hide_index=True,
            )

        elif operation == "Missingness profile":

            st.dataframe(
                missing_report(df),
                use_container_width=True,
                hide_index=True,
            )

        elif (
            operation
            == "Date trend summary"
            and dates
            and nums
        ):

            date_column = st.selectbox(
                "Date",
                dates
            )

            metric = st.selectbox(
                "Metric",
                nums
            )

            result = monthly_series(
                df,
                date_column,
                metric,
            )

            st.dataframe(
                result,
                use_container_width=True,
                hide_index=True,
            )

            if (
                px is not None
                and not result.empty
            ):

                st.plotly_chart(
                    px.line(
                        result,
                        x="Date",
                        y="Value",
                        markers=True,
                        title=f"{metric} trend",
                    ),
                    use_container_width=True,
                )

        else:

            st.info(
                "This analysis needs suitable columns."
            )


# ============================================================
# AUTO DASHBOARD
# ============================================================

elif st.session_state.page == "Auto Dashboard":

    df = require_df()

    if df is not None:

        st.subheader(
            "📊 Automatic Dashboard"
        )

        filters = {}

        categorical = [
            column
            for column
            in categorical_columns(df)
            if df[column].nunique(
                dropna=True
            ) <= 30
        ][:5]

        with st.expander(
            "Interactive filters",
            expanded=True,
        ):

            columns = st.columns(
                max(
                    1,
                    len(categorical)
                )
            )

            for index, column in enumerate(
                categorical
            ):

                values = (
                    df[column]
                    .dropna()
                    .astype(str)
                    .drop_duplicates()
                    .tolist()
                )

                selected = columns[
                    index
                ].multiselect(
                    column,
                    values,
                    key=f"dashboard_{column}",
                )

                if selected:

                    filters[column] = selected

        filtered = apply_dashboard_filters(
            df,
            filters
        )

        st.session_state.dashboard_config[
            "filters"
        ] = filters

        kpis = auto_kpis(
            filtered
        )

        columns = st.columns(
            max(
                1,
                min(
                    6,
                    len(kpis)
                )
            )
        )

        for column, item in zip(
            columns,
            kpis
        ):

            with column:

                metric_card(
                    item["label"],
                    format_value(
                        item["value"]
                    ),
                    "filtered view",
                )

        nums = numeric_columns(
            filtered
        )

        cats = categorical_columns(
            filtered
        )

        dates = detect_datetime_candidates(
            filtered
        )

        if nums:

            metric = st.selectbox(
                "Dashboard metric",
                nums
            )

            dimension = (
                st.selectbox(
                    "Breakdown",
                    cats
                )
                if cats
                else None
            )

            if (
                dimension
                and px is not None
            ):

                grouped = (
                    filtered
                    .groupby(
                        dimension,
                        dropna=False
                    )[metric]
                    .sum()
                    .reset_index()
                    .sort_values(
                        metric,
                        ascending=False
                    )
                    .head(30)
                )

                st.plotly_chart(
                    px.bar(
                        grouped,
                        x=dimension,
                        y=metric,
                        title=(
                            f"{metric} by "
                            f"{dimension}"
                        ),
                    ),
                    use_container_width=True,
                )

            if (
                dates
                and px is not None
            ):

                date_column = st.selectbox(
                    "Trend date",
                    dates
                )

                trend = monthly_series(
                    filtered,
                    date_column,
                    metric,
                )

                if not trend.empty:

                    st.plotly_chart(
                        px.line(
                            trend,
                            x="Date",
                            y="Value",
                            markers=True,
                            title=(
                                f"Monthly {metric}"
                            ),
                        ),
                        use_container_width=True,
                    )

            st.session_state.dashboard_config[
                "metric"
            ] = metric

            st.session_state.dashboard_config[
                "dimension"
            ] = dimension

        st.subheader(
            "Drill-down"
        )

        st.dataframe(
            arrow_safe_df(
                filtered.head(500)
            ),
            use_container_width=True,
            hide_index=True,
        )

        if st.button(
            "💾 Save dashboard configuration",
            type="primary",
        ):

            st.session_state.dashboard_saved = True

            st.session_state.dashboard_config[
                "saved_at"
            ] = now_text()

            add_audit(
                "Dashboard saved",
                "Filters and chart configuration saved.",
            )

            st.success(
                "Dashboard configuration saved."
            )


# ============================================================
# EXPLORE
# ============================================================

elif st.session_state.page == "Explore & Statistics":

    df = require_df()

    if df is not None:

        st.subheader(
            "🔬 Explore & Statistics"
        )

        nums = numeric_columns(
            df
        )

        if nums:

            st.dataframe(
                df[nums]
                .describe()
                .T,
                use_container_width=True,
            )

        correlation = correlation_df(
            df
        )

        if not correlation.empty:

            st.subheader(
                "Correlation Matrix"
            )

            if px is not None:

                st.plotly_chart(
                    px.imshow(
                        correlation,
                        text_auto=True,
                        aspect="auto",
                    ),
                    use_container_width=True,
                )

            else:

                st.dataframe(
                    correlation
                )

        column = st.selectbox(
            "Explore column",
            list(df.columns)
        )

        counts = (
            df[column]
            .value_counts(
                dropna=False
            )
            .head(50)
            .rename("count")
            .reset_index()
        )

        st.dataframe(
            counts,
            use_container_width=True,
            hide_index=True,
        )


# ============================================================
# FORECAST
# ============================================================

elif st.session_state.page == "Forecast & What-If":

    df = require_df()

    if df is not None:

        st.subheader(
            "🔮 Forecast & What-If"
        )

        dates = detect_datetime_candidates(
            df
        )

        nums = numeric_columns(
            df
        )

        if not dates or not nums:

            st.warning(
                "A usable date column and "
                "numeric metric are required."
            )

        else:

            c1, c2, c3 = st.columns(3)

            date_column = c1.selectbox(
                "Time field",
                dates
            )

            metric = c2.selectbox(
                "Metric",
                nums
            )

            periods = c3.slider(
                "Forecast months",
                1,
                24,
                6,
            )

            if st.button(
                "Generate forecast",
                type="primary",
            ):

                try:

                    forecast, metrics = forecast_series(
                        df,
                        date_column,
                        metric,
                        periods,
                    )

                    st.session_state.last_forecast = (
                        forecast.to_dict(
                            "records"
                        )
                    )

                    st.session_state.last_forecast_signature = (
                        dataframe_signature(
                            df
                        )
                    )

                    st.session_state.forecast_config = {
                        "date":
                            date_column,
                        "metric":
                            metric,
                        "periods":
                            periods,
                        "metrics":
                            metrics,
                    }

                    add_audit(
                        "Forecast generated",
                        f"{metric} using {metrics['method']}",
                    )

                    st.success(
                        f"Forecast generated using "
                        f"{metrics['method']}."
                    )

                except Exception as exc:

                    st.error(
                        str(exc)
                    )

            if st.session_state.last_forecast:

                forecast = pd.DataFrame(
                    st.session_state.last_forecast
                )

                if px is not None:

                    st.plotly_chart(
                        px.line(
                            forecast,
                            x="Date",
                            y="Value",
                            color="Type",
                            markers=True,
                            title=(
                                f"Historical vs forecast "
                                f"— {metric}"
                            ),
                        ),
                        use_container_width=True,
                    )

                metrics = (
                    st.session_state
                    .forecast_config
                    .get(
                        "metrics",
                        {}
                    )
                )

                columns = st.columns(4)

                cards = [
                    (
                        "MAE",
                        format_value(
                            metrics.get(
                                "mae"
                            )
                        ),
                    ),
                    (
                        "RMSE",
                        format_value(
                            metrics.get(
                                "rmse"
                            )
                        ),
                    ),
                    (
                        "MAPE",
                        f"{metrics.get('mape', 0):.2f}%",
                    ),
                    (
                        "Method",
                        metrics.get(
                            "method",
                            "—"
                        ),
                    ),
                ]

                for column, (
                    label,
                    value
                ) in zip(
                    columns,
                    cards
                ):

                    with column:

                        metric_card(
                            label,
                            value,
                            "forecast metric",
                        )

                st.dataframe(
                    forecast,
                    use_container_width=True,
                    hide_index=True,
                )

                st.subheader(
                    "What-If Scenario"
                )

                change = st.slider(
                    "Change forecast by",
                    -50,
                    100,
                    10,
                    1,
                    format="%d%%",
                )

                forecast_only = forecast[
                    forecast["Type"]
                    == "Forecast"
                ].copy()

                if not forecast_only.empty:

                    forecast_only[
                        "ScenarioValue"
                    ] = (
                        forecast_only[
                            "Value"
                        ]
                        * (
                            1
                            + change / 100
                        )
                    )

                    st.dataframe(
                        forecast_only,
                        use_container_width=True,
                        hide_index=True,
                    )


# ============================================================
# AI DATA ANALYST
# ============================================================

elif st.session_state.page == "AI Data Analyst":

    df = require_df()

    if df is not None:

        st.subheader(
            "🤖 Local AI Data Analyst"
        )

        evidence = build_evidence_ledger(
            df
        )

        st.session_state.evidence_ledger = (
            evidence
        )

        findings = detective_findings(
            df
        )

        c1, c2, c3 = st.columns(3)

        with c1:

            metric_card(
                "Evidence",
                f"{len(evidence)}",
                "deterministic records",
            )

        with c2:

            metric_card(
                "Quality",
                f"{quality_report(df)['score']}/100",
                "verified metric",
            )

        with c3:

            metric_card(
                "Model",
                st.session_state.selected_model,
                "local Ollama",
            )

        st.caption(
            "Data questions are grounded in deterministic evidence. "
            "General questions are routed directly to Phi-3. "
            "Numerical claims are never trusted solely to the LLM."
        )

        with st.expander(
            "Data Detective Evidence",
            expanded=True,
        ):

            render_findings(
                findings
            )

        quick_questions = [
            "Custom question",
            "What are the most important findings?",
            "Why are there missing values?",
            "What are the main data-quality risks?",
            "Find potential anomalies or outliers.",
            "Summarize the main numeric patterns.",
            "What is AI?",
        ]

        selected_quick = st.selectbox(
            "Quick question",
            quick_questions
        )

        if selected_quick == "Custom question":

            question = st.chat_input(
                "Ask your data or ask Phi-3 a general question..."
            )

        else:

            question = selected_quick

            if st.button(
                f"Ask: {selected_quick}",
                type="primary",
            ):

                pass

        if question:

            with st.chat_message(
                "user"
            ):

                st.markdown(
                    question
                )

            with st.chat_message(
                "assistant"
            ):

                placeholder = st.empty()

                answer = ""

                try:

                    if is_general_question(
                        question
                    ):

                        prompt = (
                            build_general_prompt(
                                question
                            )
                        )

                    else:

                        prompt = (
                            build_ai_prompt(
                                question,
                                df,
                            )
                        )

                    placeholder.markdown(
                        "▌"
                    )

                    for token in stream_ollama_answer(
                        prompt,
                        st.session_state.selected_model,
                    ):

                        answer += token

                        placeholder.markdown(
                            answer
                            + "▌"
                        )

                    placeholder.markdown(
                        answer
                    )

                except Exception as exc:

                    if is_general_question(
                        question
                    ):

                        answer = (
                            "Local Ollama did not "
                            "respond. Start Ollama and "
                            "make sure "
                            f"{st.session_state.selected_model} "
                            "is installed."
                        )

                    else:

                        answer = (
                            deterministic_ai_fallback(
                                question,
                                df
                            )
                        )

                    placeholder.markdown(
                        answer
                    )

                    add_audit(
                        "AI fallback",
                        str(exc)[:250],
                    )

                st.session_state.chat_history.append(
                    {
                        "role": "user",
                        "content": question,
                        "time": now_text(),
                    }
                )

                st.session_state.chat_history.append(
                    {
                        "role": "assistant",
                        "content": answer,
                        "time": now_text(),
                    }
                )

                st.session_state.chat_history = (
                    st.session_state
                    .chat_history[
                        -MAX_CHAT_MESSAGES:
                    ]
                )

        if st.session_state.chat_history:

            st.divider()

            st.subheader(
                "Conversation"
            )

            for message in (
                st.session_state.chat_history
            ):

                with st.chat_message(
                    message.get(
                        "role",
                        "assistant"
                    )
                ):

                    st.markdown(
                        message.get(
                            "content",
                            ""
                        )
                    )


# ============================================================
# EVIDENCE & VERIFICATION
# ============================================================

elif st.session_state.page == "Evidence & Verification":

    df = require_df()

    if df is not None:

        st.subheader(
            "🔬 Evidence & Verification"
        )

        ledger = build_evidence_ledger(
            df
        )

        st.session_state.evidence_ledger = (
            ledger
        )

        st.dataframe(
            pd.DataFrame(ledger),
            use_container_width=True,
            hide_index=True,
        )

        st.download_button(
            "Download Evidence JSON",
            json.dumps(
                ledger,
                indent=2,
                default=str,
            ).encode(),
            file_name="evidence_ledger.json",
            mime="application/json",
        )

        st.divider()

        if not st.session_state.findings:

            generated = []

            for item in detective_findings(
                df
            )[:8]:

                generated.append(
                    {
                        "id":
                            hashlib.md5(
                                (
                                    item["title"]
                                    + item["text"]
                                ).encode()
                            ).hexdigest()[:10],

                        "title":
                            item["title"],

                        "claim":
                            item["text"],

                        "text":
                            item["text"],

                        "status":
                            "UNVERIFIED",
                    }
                )

            st.session_state.findings = (
                generated
            )

        st.subheader(
            "Finding Lifecycle"
        )

        for finding in (
            st.session_state.findings
        ):

            with st.container(
                border=True
            ):

                st.markdown(
                    f"**{finding.get('id')} — "
                    f"{finding.get('title')}**"
                )

                st.write(
                    finding.get(
                        "claim",
                        ""
                    )
                )

                st.caption(
                    "Status: "
                    + str(
                        finding.get(
                            "status",
                            "UNVERIFIED"
                        )
                    )
                )

                c1, c2 = st.columns(2)

                with c1:

                    if st.button(
                        "✓ Verify",
                        key=(
                            "verify_"
                            + finding["id"]
                        ),
                    ):

                        result = verify_finding(
                            df,
                            finding
                        )

                        finding[
                            "status"
                        ] = result[
                            "status"
                        ]

                        finding[
                            "verification"
                        ] = result

                        st.session_state.verification_results.append(
                            result
                        )

                        st.success(
                            result[
                                "summary"
                            ]
                        )

                with c2:

                    if st.button(
                        "⚔ Challenge",
                        key=(
                            "challenge_"
                            + finding["id"]
                        ),
                    ):

                        result = challenge_finding(
                            df,
                            finding
                        )

                        finding[
                            "challenge"
                        ] = result

                        st.session_state.challenge_results.append(
                            result
                        )

                        st.info(
                            result[
                                "summary"
                            ]
                        )

        if st.session_state.verification_results:

            st.subheader(
                "Verification History"
            )

            st.dataframe(
                pd.DataFrame(
                    st.session_state
                    .verification_results
                ),
                use_container_width=True,
                hide_index=True,
            )

        if st.session_state.challenge_results:

            st.subheader(
                "Challenge History"
            )

            st.dataframe(
                pd.DataFrame(
                    st.session_state
                    .challenge_results
                ),
                use_container_width=True,
                hide_index=True,
            )


# ============================================================
# REPORT STUDIO
# ============================================================

elif st.session_state.page == "Reports & Exports":

    df = require_df()

    if df is not None:

        st.subheader(
            "📑 Report Studio"
        )

        st.caption(
            "Select exactly what should appear in your executive PDF "
            "and PowerPoint. Excel contains the complete analytical workbook."
        )

        config = get_report_config(
            df
        )

        nums = numeric_columns(
            df
        )

        dates = detect_datetime_candidates(
            df
        )

        cats = categorical_columns(
            df
        )

        c1, c2 = st.columns(2)

        with c1:

            title = st.text_input(
                "Report title",
                config["title"],
            )

            subtitle = st.text_input(
                "Subtitle",
                config["subtitle"],
            )

        with c2:

            metric = (
                st.selectbox(
                    "Primary metric",
                    nums,
                    index=(
                        nums.index(
                            config["metric"]
                        )
                        if config["metric"] in nums
                        else 0
                    ),
                )
                if nums
                else None
            )

            date_column = (
                st.selectbox(
                    "Time column",
                    dates,
                    index=(
                        dates.index(
                            config["date"]
                        )
                        if config["date"] in dates
                        else 0
                    ),
                )
                if dates
                else None
            )

        sections_all = [
            "Executive Summary",
            "KPI Snapshot",
            "Dashboard Charts",
            "Forecast",
            "What-If",
            "Evidence",
            "Verification",
            "Challenge",
            "Data Quality",
            "SQL History",
            "AI Conversation",
            "Transformations",
            "Audit / Provenance",
        ]

        sections = st.multiselect(
            "Report Sections",
            sections_all,
            default=[
                item
                for item in config[
                    "sections"
                ]
                if item in sections_all
            ],
        )

        chart_options = [
            "trend",
            "segment",
            "distribution",
            "correlation",
        ]

        charts = st.multiselect(
            "Dashboard Charts",
            chart_options,
            default=[
                chart
                for chart in config[
                    "charts"
                ]
                if chart in chart_options
            ],
        )

        evidence = build_evidence_ledger(
            df
        )

        evidence_ids = st.multiselect(
            "Evidence to include",
            [
                item["id"]
                for item in evidence
            ],
            default=[
                item["id"]
                for item in evidence
                if item["id"]
                in config.get(
                    "evidence_ids",
                    []
                )
            ],
        )

        forecast_details = st.multiselect(
            "Forecast details",
            [
                "summary",
                "metrics",
                "chart",
                "table",
            ],
            default=[
                item
                for item in config[
                    "forecast_details"
                ]
                if item in [
                    "summary",
                    "metrics",
                    "chart",
                    "table",
                ]
            ],
        )

        what_if_change = st.slider(
            "What-if scenario (%)",
            -50,
            100,
            int(
                config.get(
                    "what_if_change",
                    10
                )
            ),
        )

        st.session_state.report_config = {
            "title": title,
            "subtitle": subtitle,
            "sections": sections,
            "charts": charts,
            "metric": metric,
            "date": date_column,
            "evidence_ids": evidence_ids,
            "forecast_details": forecast_details,
            "what_if_change":
                what_if_change,
        }

        st.divider()

        st.subheader(
            "Report Preview"
        )

        preview = st.columns(4)

        preview_values = [
            (
                "Sections",
                len(sections)
            ),
            (
                "Charts",
                len(charts)
            ),
            (
                "Evidence",
                len(evidence_ids)
            ),
            (
                "Forecast details",
                len(forecast_details)
            ),
        ]

        for column, (
            label,
            value
        ) in zip(
            preview,
            preview_values
        ):

            with column:

                metric_card(
                    label,
                    str(value),
                    "selected",
                )

        images = report_chart_images(
            df,
            st.session_state.report_config
        )

        for chart in charts:

            if chart in images:

                st.image(
                    images[chart],
                    caption=(
                        chart
                        .replace(
                            "_",
                            " "
                        )
                        .title()
                    ),
                    use_container_width=True,
                )

            else:

                st.warning(
                    f"{chart.title()} "
                    "could not be rendered."
                )

        st.divider()

        st.subheader(
            "Export Center"
        )

        c1, c2, c3 = st.columns(3)

        with c1:

            st.download_button(
                "📊 Clean CSV",
                df.to_csv(
                    index=False
                ).encode(
                    "utf-8-sig"
                ),
                file_name=(
                    f"{safe_filename(current_dataset_name())}"
                    "_cleaned.csv"
                ),
                mime="text/csv",
                use_container_width=True,
            )

        with c2:

            try:

                excel_bytes = (
                    dataframe_to_excel(
                        df
                    )
                )

                st.download_button(
                    "📗 Professional Excel",
                    excel_bytes,
                    file_name=(
                        f"{safe_filename(current_dataset_name())}"
                        "_analysis.xlsx"
                    ),
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    ),
                    use_container_width=True,
                )

            except Exception as exc:

                st.error(
                    str(exc)
                )

        with c3:

            try:

                pdf_bytes = make_pdf(
                    df
                )

                st.download_button(
                    "📄 Executive PDF",
                    pdf_bytes,
                    file_name=(
                        f"{safe_filename(current_dataset_name())}"
                        "_executive_report.pdf"
                    ),
                    mime="application/pdf",
                    use_container_width=True,
                )

            except Exception as exc:

                st.error(
                    str(exc)
                )

        c4, c5 = st.columns(2)

        with c4:

            try:

                ppt_bytes = make_pptx(
                    df
                )

                st.download_button(
                    "📽️ Executive PowerPoint",
                    ppt_bytes,
                    file_name=(
                        f"{safe_filename(current_dataset_name())}"
                        "_executive_report.pptx"
                    ),
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.presentationml.presentation"
                    ),
                    use_container_width=True,
                )

            except Exception as exc:

                st.error(
                    str(exc)
                )

        with c5:

            package = {
                "findings":
                    st.session_state.findings,

                "evidence":
                    st.session_state.evidence_ledger,

                "verification":
                    st.session_state.verification_results,

                "challenge":
                    st.session_state.challenge_results,

                "report_config":
                    st.session_state.report_config,
            }

            st.download_button(
                "📋 Findings + Evidence JSON",
                json.dumps(
                    package,
                    indent=2,
                    default=str,
                ).encode(),
                file_name=(
                    "findings_evidence.json"
                ),
                mime="application/json",
                use_container_width=True,
            )

        st.divider()

        st.subheader(
            "Export Design"
        )

        st.info(
            "PDF and PowerPoint use dedicated Matplotlib "
            "report figures instead of relying on Plotly/Kaleido. "
            "This prevents the previous 'Chart image generation "
            "is unavailable' placeholder."
        )


# ============================================================
# SETTINGS / SAVE / RESUME
# ============================================================

else:

    st.subheader(
        "⚙️ Settings & Save / Resume"
    )

    df = current_df()

    if df is not None:

        st.markdown(
            "### Save Complete Analysis Project"
        )

        name = st.text_input(
            "Project name",
            value=st.session_state.project_name,
        )

        c1, c2 = st.columns(2)

        with c1:

            if st.button(
                "💾 Save Project",
                type="primary",
                use_container_width=True,
            ):

                try:

                    persist_project_to_disk(
                        name
                    )

                    st.success(
                        f"Saved {name}."
                    )

                except Exception as exc:

                    st.error(
                        str(exc)
                    )

        with c2:

            try:

                project_files = (
                    save_project_bytes(
                        name
                    )
                )

                stem = safe_filename(
                    name
                )

                st.download_button(
                    "⬇️ Download Project JSON",
                    project_files[
                        f"{stem}.json"
                    ],
                    file_name=(
                        f"{stem}.json"
                    ),
                    mime="application/json",
                    use_container_width=True,
                )

            except Exception as exc:

                st.warning(
                    str(exc)
                )

    st.divider()

    st.markdown(
        "### Resume a Project"
    )

    projects = project_list()

    if projects:

        selected_project = st.selectbox(
            "Saved projects",
            projects,
            format_func=lambda path:
                path.stem,
        )

        if st.button(
            "Open Selected Project",
            type="primary",
        ):

            try:

                load_project_from_disk(
                    selected_project
                )

                st.success(
                    f"Resumed "
                    f"{selected_project.stem}."
                )

                st.rerun()

            except Exception as exc:

                st.error(
                    str(exc)
                )

    else:

        st.info(
            "No saved projects yet."
        )

    st.divider()

    st.markdown(
        "### Runtime"
    )

    st.code(
        f"""
Application: {APP_NAME}
Version: {APP_VERSION}
Ollama URL: {OLLAMA_URL}
Default model: {OLLAMA_MODEL}
Project directory: {PROJECT_DIR.resolve()}
        """
    )

    st.markdown(
        "### Dependency Status"
    )

    dependency_status = {
        "requests":
            requests is not None,
        "plotly":
            px is not None,
        "duckdb":
            duckdb is not None,
        "statsmodels":
            ExponentialSmoothing is not None,
        "matplotlib":
            plt is not None,
        "openpyxl":
            OPENPYXL_AVAILABLE,
        "reportlab":
            REPORTLAB_AVAILABLE,
        "python-pptx":
            PPTX_AVAILABLE,
    }

    st.table(
        pd.DataFrame(
            [
                {
                    "Package":
                        package,
                    "Available":
                        "✓"
                        if available
                        else "✗",
                }
                for package, available
                in dependency_status.items()
            ]
        )
    )

    st.code(
        "pip install streamlit pandas numpy "
        "openpyxl xlrd requests plotly duckdb "
        "statsmodels reportlab python-pptx matplotlib"
    )


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    "---"
)

st.markdown(
    f"""
<div class="small">
AI Data Analyst v{APP_VERSION}
•
Local-first analytics
•
Ollama / Phi-3
•
Deterministic Evidence
→ Local AI
→ Verification
→ Challenge
→ Executive Reports
•
No paid AI API
</div>
""",
    unsafe_allow_html=True,
)
app.py
Displaying app.py.
