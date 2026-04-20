import json
from datetime import date
from pathlib import Path
import re

import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import matplotlib.pyplot as plt
import math
import base64
import io


# =============================
# CONFIG: your local Excel path
# =============================
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
CLEAN_DATA_DIR = BASE_DIR / "clean_data"
REPORT_DATE_PATTERN = re.compile(r"(\d{1,2})-(\d{1,2})-(\d{4})")
SEMESTER_WINDOWS = [
    ((1, 12), (5, 8), "Spring"),
    ((5, 12), (8, 15), "Summer"),
]


def list_clean_files(clean_dir: Path) -> list[Path]:
    candidates = [f for f in clean_dir.glob("*.xlsx") if not f.name.startswith("~$")]
    if not candidates:
        raise FileNotFoundError(
            f"No cleaned Excel file found in {clean_dir}. Run convert_excel.py first."
        )
    return sorted(
        candidates,
        key=lambda p: (extract_report_date(p), p.stat().st_mtime),
        reverse=True,
    )


def extract_report_date(file_path: Path) -> pd.Timestamp:
    match = REPORT_DATE_PATTERN.search(file_path.stem)
    if not match:
        return pd.Timestamp(date.today())
    month, day, year = map(int, match.groups())
    return pd.Timestamp(year=year, month=month, day=day)


# =============================
# Page config
# =============================
st.set_page_config(page_title="HR Dashboard", layout="wide")


def build_clean_file_options(clean_dir: Path):
    files = list_clean_files(clean_dir)
    option_labels = []
    option_to_file: dict[str, Path] = {}
    for f in files:
        report_dt = extract_report_date(f)
        month_label = report_dt.strftime("%B %Y")
        as_of_label = f"{report_dt.strftime('%B')} {report_dt.day}"
        label = f"{month_label} (As of {as_of_label})"
        option_labels.append(label)
        option_to_file[label] = f
    return option_labels, option_to_file


# =============================
# Global CSS (top-level)
# =============================
st.markdown(
    """
<style>
div[data-testid="stMetric"] {
  padding: 8px 10px;
  border-radius: 10px;
}
div[data-testid="stMetricLabel"] > div {
  font-size: 14px !important;
  font-weight: 700 !important;
}
div[data-testid="stMetricValue"] > div {
  font-size: 36px !important;
  font-weight: 800 !important;
  line-height: 1.0 !important;
}
</style>
""",
    unsafe_allow_html=True,
)


# =============================
# Load Excel -> employees DataFrame
# We ONLY use these columns:
# Full name, Gender Code, Birth Date, PD, Employment Type, Last Hire Date, supervisor full name, Job
# =============================
@st.cache_data(show_spinner=False)
def load_employees(excel_path: str, sheet_name=None) -> pd.DataFrame:
    raw = pd.read_excel(excel_path, sheet_name=sheet_name)

    # if sheet_name=None -> dict of sheets
    if isinstance(raw, dict):
        raw = list(raw.values())[0]

    # Clean headers
    raw.columns = (
        pd.Index(raw.columns)
        .map(lambda x: str(x).replace("\n", " ").replace("\r", " ").strip())
    )

    normalized_map = {str(col).strip().lower(): col for col in raw.columns}

    required_cols = [
        "Full name",
        "Gender Code",
        "Birth Date",
        "PD",
        "Employment Type",
        "Last Hire Date",
        "supervisor full name",
        "Job",  # ✅ needed for functional org chart
    ]

    missing = [c for c in required_cols if c.strip().lower() not in normalized_map]
    if missing:
        raise ValueError(
            f"Missing required columns: {missing}\n\nDetected columns: {list(raw.columns)}"
        )

    optional_cols = []
    for optional_name in ["Is Coop", "Employee Group"]:
        key = optional_name.strip().lower()
        if key in normalized_map:
            optional_cols.append(normalized_map[key])

    resolved_cols = [normalized_map[c.strip().lower()] for c in required_cols] + optional_cols
    df = raw[resolved_cols].copy()

    # Rename to internal names used downstream
    df = df.rename(
        columns={
            normalized_map["full name"]: "name",
            normalized_map["gender code"]: "gender",
            normalized_map["birth date"]: "birth_date",
            normalized_map["pd"]: "pd",
            normalized_map["employment type"]: "employee_class",
            normalized_map["last hire date"]: "hire_date",
            normalized_map["supervisor full name"]: "manager",
            normalized_map["job"]: "job",
            **(
                {normalized_map["is coop"]: "is_coop"}
                if "is coop" in normalized_map else {}
            ),
            **(
                {normalized_map["employee group"]: "employee_group"}
                if "employee group" in normalized_map else {}
            ),
        }
    )

    # Parse dates
    df["hire_date"] = pd.to_datetime(df["hire_date"], errors="coerce")
    df["birth_date"] = pd.to_datetime(df["birth_date"], errors="coerce")

    # employee_id: temporary use full name
    df["employee_id"] = df["name"].fillna("").astype(str).str.strip()

    # cleanup text
    df["pd"] = df["pd"].fillna("Unknown").astype(str).str.strip()
    df["manager"] = df["manager"].fillna("Unknown").astype(str).str.strip()
    df["gender"] = df["gender"].fillna("Unknown").astype(str).str.strip()
    df["employee_class"] = df["employee_class"].fillna("Unknown").astype(str).str.strip()
    df["job"] = df["job"].fillna("").astype(str).str.strip()

    employment_type_key = (
        df["employee_class"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
        .str.replace("-", "", regex=False)
        .str.replace(" ", "", regex=False)
    )

    if "is_coop" in df.columns:
        df["is_coop"] = (
            df["is_coop"]
            .fillna(False)
            .astype(str)
            .str.strip()
            .str.lower()
            .isin(["true", "1", "yes"])
        )
    else:
        df["is_coop"] = employment_type_key.eq("coop")

    if "employee_group" in df.columns:
        df["employee_group"] = df["employee_group"].fillna("").astype(str).str.strip()
        df.loc[df["employee_group"] == "", "employee_group"] = np.where(df["is_coop"], "Coop", "Core")
    else:
        df["employee_group"] = np.where(df["is_coop"], "Coop", "Core")

    # placeholder team (not used)
    df["team"] = "N/A"

    return df[
        [
            "employee_id", "name", "pd", "team", "manager",
            "hire_date", "birth_date", "gender", "employee_class",
            "job", "is_coop", "employee_group"
        ]
    ].copy()


@st.cache_data(show_spinner=False)
def image_to_base64(image_path: str) -> str:
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode("utf-8")


def fig_to_base64(fig) -> str:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=180, bbox_inches="tight", facecolor=fig.get_facecolor())
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")


# =============================
# Force page start from top
# =============================
st.markdown(
    """
    <script>
        window.scrollTo(0, 0);
    </script>
    """,
    unsafe_allow_html=True,
)

# =============================
# Reporting month / Header
# =============================
try:
    report_options, report_file_map = build_clean_file_options(CLEAN_DATA_DIR)
except Exception as e:
    st.error(f"Failed to find cleaned Excel files: {e}")
    st.info("Run convert_excel.py first, then make sure clean_data contains at least one valid Excel file.")
    st.stop()

selected_report_label = st.session_state.get("report_month", report_options[0])
if selected_report_label not in report_file_map:
    selected_report_label = report_options[0]

EXCEL_PATH = report_file_map[selected_report_label]
SHEET_NAME = None
as_of = extract_report_date(EXCEL_PATH)
dashboard_title = "MARC HR Dashboard"
report_period_label = as_of.strftime("%B %Y")
report_as_of_label = f"{as_of.strftime('%B')} {as_of.day}, {as_of.year}"

try:
    employees = load_employees(EXCEL_PATH, SHEET_NAME)
except Exception as e:
    st.error(f"Failed to load Excel: {e}")
    st.info("Please make sure the selected file in clean_data contains the required cleaned columns.")
    st.stop()

LOGO_PATH = BASE_DIR / "Midea.png"
logo_b64 = image_to_base64(LOGO_PATH)
ROW1_LEFT_IMAGE_PATH = BASE_DIR / "row1left.png"
MAN_IMAGE_PATH = BASE_DIR / "man.jpg"
WOMAN_IMAGE_PATH = BASE_DIR / "woman.png"

row1_left_b64 = image_to_base64(ROW1_LEFT_IMAGE_PATH)
man_b64 = image_to_base64(MAN_IMAGE_PATH)
woman_b64 = image_to_base64(WOMAN_IMAGE_PATH)

st.markdown(
    """
    <style>
    div[data-testid="stAppViewContainer"] {
        background: linear-gradient(180deg, #f8fcff 0%, #eef7fc 100%);
    }

    section[data-testid="stMain"] {
        background: transparent;
    }

    header[data-testid="stHeader"] {
        height: 0rem;
        background: transparent;
    }

    div[data-testid="stToolbar"] {
        top: 0.25rem;
        right: 0.5rem;
    }

    section[data-testid="stMain"] div.block-container {
        max-width: 100%;
        padding-top: 0 !important;
        padding-bottom: 1.15rem;
        padding-left: 0.55rem !important;
        padding-right: 0.55rem !important;
    }

    hr {
        border: none;
        border-top: 1px solid #cfe5f3;
        margin: 2rem 0 1.6rem 0;
    }

    .section-heading {
        margin: 0 0 14px 0;
        padding-left: 14px;
        border-left: 4px solid #0096db;
        color: #0e3a67;
        font-size: 28px;
        font-weight: 900;
        letter-spacing: 0.01em;
    }

    .header-logo-wrap {
        display:flex;
        justify-content:flex-start;
        align-items:flex-start;
        margin-top: -54px;
        padding-left: 6px;
    }

    .header-logo {
        width: 168px;
        max-width: 100%;
        height: auto;
        display: block;
    }

    .header-wrap {
        margin-top: -14px;
    }

    .header-title {
        text-align: center;
        font-size: 1.95rem;
        font-weight: 850;
        color:#0e3a67;
        line-height: 1.02;
        margin-top: 0px;
        margin-bottom: 0px;
        letter-spacing: -0.03em;
    }

    .header-subtitle {
        text-align:center;
        font-size:11px;
        color:#5d7b94;
        margin-top:0px;
        margin-bottom:0px;
    }

    .month-filter-label {
        text-align:center;
        font-size:11px;
        font-weight:800;
        color:#5d7b94;
        margin-top:2px;
        margin-bottom:4px;
    }

    .month-filter-inline-label {
        font-size: 11px;
        font-weight: 800;
        color: #5d7b94;
        white-space: nowrap;
        text-align: right;
        padding-top: 9px;
    }

    div[data-baseweb="select"] > div {
        border-radius: 10px !important;
        border-color: #cfe5f3 !important;
        min-height: 36px !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

col_logo, col_title, col_right = st.columns([1.8, 4.4, 1.8], gap="small")

with col_logo:
    st.markdown(
        f"""
        <div class="header-logo-wrap">
            <img class="header-logo" src="data:image/png;base64,{logo_b64}" alt="Midea logo" />
        </div>
        """,
        unsafe_allow_html=True,
    )

with col_title:
    st.markdown(
        f"""
        <div class="header-wrap">
            <div class="header-title">{dashboard_title}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    filter_left, filter_mid, filter_right = st.columns([1.15, 1.25, 1.15])
    with filter_mid:
        st.markdown('<div class="month-filter-label">Reporting Month</div>', unsafe_allow_html=True)
        st.selectbox(
            "Reporting Month",
            report_options,
            index=report_options.index(selected_report_label),
            key="report_month",
            label_visibility="collapsed",
        )

with col_right:
    st.empty()

st.markdown("<div style='height: 0px;'></div>", unsafe_allow_html=True)

# =============================
# Helpers
# =============================
def years_between(start: pd.Timestamp, end: pd.Timestamp) -> float:
    if pd.isna(start) or pd.isna(end):
        return np.nan
    return (end - start).days / 365.25


def avg_years_from_dates(series_start_dates: pd.Series, as_of_dt: pd.Timestamp) -> float:
    vals = series_start_dates.apply(lambda d: years_between(d, as_of_dt)).dropna()
    return float(vals.mean()) if len(vals) else float("nan")


def avg_industry_years_from_birth(birth_dates: pd.Series, as_of_dt: pd.Timestamp) -> float:
    ages = birth_dates.apply(lambda d: years_between(d, as_of_dt)).dropna()
    if len(ages) == 0:
        return float("nan")
    industry = (ages - 22).clip(lower=0)
    return float(industry.mean())


def month_range(as_of_dt: pd.Timestamp):
    start = pd.Timestamp(as_of_dt.year, as_of_dt.month, 1)
    end = start + pd.offsets.MonthBegin(1)
    return start, end


def get_semester_label(as_of_dt: pd.Timestamp) -> str | None:
    dt = pd.Timestamp(as_of_dt).normalize()
    for (start_month, start_day), (end_month, end_day), semester_name in SEMESTER_WINDOWS:
        start = pd.Timestamp(year=dt.year, month=start_month, day=start_day)
        end = pd.Timestamp(year=dt.year, month=end_month, day=end_day)
        if start <= dt <= end:
            return f"{dt.year} {semester_name}"
    return None


def new_hires_in_month(df: pd.DataFrame, as_of_dt: pd.Timestamp) -> pd.DataFrame:
    m_start, m_end = month_range(as_of_dt)
    out = df[(df["hire_date"] >= m_start) & (df["hire_date"] < m_end)].copy()
    out = out.sort_values("hire_date")
    return out[["employee_id", "name", "pd", "manager", "hire_date"]]


def service_anniversaries_in_month(df: pd.DataFrame, as_of_dt: pd.Timestamp) -> pd.DataFrame:
    out = df[df["hire_date"].notna()].copy()
    out = out[out["hire_date"].dt.month == as_of_dt.month].copy()
    out["anniv_years"] = as_of_dt.year - out["hire_date"].dt.year
    out = out[out["anniv_years"] >= 1].copy()
    out = out.sort_values(["anniv_years", "hire_date"], ascending=[False, True])
    return out[["employee_id", "name", "pd", "manager", "hire_date", "anniv_years"]]


def is_dark_color(color) -> bool:
    if isinstance(color, str) and color.startswith("#") and len(color) == 7:
        r = int(color[1:3], 16) / 255
        g = int(color[3:5], 16) / 255
        b = int(color[5:7], 16) / 255
    else:
        r, g, b = color[:3]
    luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b
    return luminance < 0.58


# =============================
# Data view (no filters)
# =============================
employees_view = employees.copy()

# Core headcount excludes interns, contractors, and coops
exclude_classes = {"Intern", "Contractor"}
employees_coop = employees_view[employees_view["is_coop"]].copy()
employees_core = employees_view[
    (~employees_view["employee_class"].isin(exclude_classes)) & (~employees_view["is_coop"])
].copy()

hc = int(len(employees_core))
new_hires_mtd = len(new_hires_in_month(employees_core, as_of))
avg_midea = avg_years_from_dates(employees_core["hire_date"], as_of)
avg_industry = avg_industry_years_from_birth(employees_core["birth_date"], as_of)


st.markdown(
    """
    <style>
    .block-card{
      border:1px solid #d6eaf6;
      border-radius:18px;
      padding:14px 14px;
      background:linear-gradient(180deg, #ffffff 0%, #f4fbff 100%);
      box-shadow: 0 14px 28px rgba(14,58,103,.08);
      margin-bottom:12px;
    }
    .tag-container{
      display:flex;
      flex-wrap:wrap;
      gap:10px;
    }
    .tag{
      border:1px solid #d8ebf7;
      background:#eaf7fd;
      border-radius:999px;
      padding:6px 10px;
      font-size:15px;
      font-weight:900;
      color: #1d4f77;
      line-height:1;
      white-space:nowrap;
    }
    .svc-row{
      display:flex;
      align-items:flex-start;
      gap:12px;
      margin-bottom:10px;
    }
    .svc-row:last-child{ margin-bottom:0; }
    .svc-year{
      min-width:82px;
      font-weight:900;
      font-size:15px;
      color: #0e3a67;
      padding-top:2px;
    }
    .section-title{
      font-size:22px;
      font-weight:900;
      color:#0e3a67;
      padding-left:14px;
      border-left:4px solid #0096db;
      margin-bottom:8px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def render_new_hires_card():
    nh = new_hires_in_month(employees_core, as_of)
    st.markdown(f'<div class="section-title">New Hire ({report_period_label})</div>', unsafe_allow_html=True)
    if nh.empty:
        st.markdown(
        f"""
        <div class="block-card">
          <div style="color:#5d7b94; font-weight:900; font-size:15px;">
            No new hires through {report_as_of_label}.
          </div>
        </div>
            """,
            unsafe_allow_html=True,
        )
        return

    nh_names = sorted(nh["name"].astype(str).str.strip().tolist())
    tags_html = "".join([f'<span class="tag">{n}</span>' for n in nh_names])
    st.markdown(
        f"""
        <div class="block-card">
          <div class="tag-container">
            {tags_html}
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_service_anniversary_card():
    an = service_anniversaries_in_month(employees_core, as_of)
    st.markdown(f'<div class="section-title">Service Anniversary ({report_period_label})</div>', unsafe_allow_html=True)
    if an.empty:
        st.markdown(
        f"""
        <div class="block-card">
          <div style="color:#5d7b94; font-weight:900; font-size:15px;">
            No service anniversaries through {report_as_of_label}.
          </div>
        </div>
            """,
            unsafe_allow_html=True,
        )
        return

    an2 = an[["anniv_years", "name"]].copy()
    an2["name"] = an2["name"].astype(str).str.strip()
    grp = (
        an2.groupby("anniv_years")["name"]
        .apply(lambda s: sorted(s.tolist()))
        .sort_index(ascending=False)
    )

    rows_html = ""
    for years, names in grp.items():
        tags_html = "".join([f'<span class="tag">{n}</span>' for n in names])
        label = f"{int(years)} Year" + ("s" if int(years) > 1 else "")
        rows_html += f"""
          <div class="svc-row">
            <div class="svc-year">{label}</div>
            <div class="tag-container">
              {tags_html}
            </div>
          </div>
        """

    st.markdown(
        f"""
        <div class="block-card">
          {rows_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_workforce_snapshot():
    kpi_card = f"""
    <div style="
        width:100%;
        max-width:100%;
        border:1px solid #d6eaf6;
        border-radius:22px;
        background:linear-gradient(145deg, #0d355f 0%, #0f4577 40%, #0a84c7 100%);
        height:304px;
        box-sizing:border-box;
        padding:8px 10px 8px 10px;
        display:flex;
        flex-direction:column;
        box-shadow:none;
        position:relative;
        overflow:hidden;
    ">
      <div style="height:4px; border-radius:999px; background:linear-gradient(90deg, rgba(255,255,255,.9) 0%, rgba(196,233,250,.95) 55%, rgba(120,202,235,.9) 100%); margin-bottom:6px;"></div>
      <div style="position:relative; min-height:68px; margin-bottom:6px;">
        <div style="font-size:22px; font-weight:950; color:white; margin-top:8px; line-height:1.02; max-width:54%;">Workforce Snapshot</div>
        <div style="position:absolute; top:2px; right:82px; font-size:10px; font-weight:700; color:rgba(255,255,255,.82); letter-spacing:.01em; white-space:nowrap;">
          Excludes coops and contractors
        </div>
        <img src="data:image/png;base64,{row1_left_b64}" alt="workforce illustration" style="position:absolute; top:0; right:0; height:72px; width:auto; object-fit:contain; filter:none; z-index:2;" />
      </div>
      <div style="display:grid; grid-template-columns:1fr 1fr; gap:8px; flex:1;">
        <div style="background:rgba(255,255,255,.12); border:1px solid rgba(255,255,255,.18); border-radius:14px; padding:8px 8px; backdrop-filter:blur(8px); display:flex; flex-direction:column; justify-content:center; align-items:center; text-align:center;">
          <div style="font-size:15px; font-weight:900; color:rgba(255,255,255,.84); letter-spacing:.03em; text-transform:uppercase;">Headcount</div>
          <div style="font-size:27px; font-weight:950; color:white; margin-top:4px; line-height:1;">{hc}</div>
        </div>
        <div style="background:rgba(255,255,255,.12); border:1px solid rgba(255,255,255,.18); border-radius:14px; padding:8px 8px; backdrop-filter:blur(8px); display:flex; flex-direction:column; justify-content:center; align-items:center; text-align:center;">
          <div style="font-size:15px; font-weight:900; color:rgba(255,255,255,.84); letter-spacing:.03em; text-transform:uppercase;">New Hires (MTD)</div>
          <div style="font-size:27px; font-weight:950; color:white; margin-top:4px; line-height:1;">{new_hires_mtd}</div>
        </div>
        <div style="background:rgba(255,255,255,.12); border:1px solid rgba(255,255,255,.18); border-radius:14px; padding:8px 8px; backdrop-filter:blur(8px); display:flex; flex-direction:column; justify-content:center; align-items:center; text-align:center;">
          <div style="font-size:15px; font-weight:900; color:rgba(255,255,255,.84); letter-spacing:.03em; text-transform:uppercase;">Midea Tenure</div>
          <div style="font-size:27px; font-weight:950; color:white; margin-top:4px; line-height:1;">{"-" if np.isnan(avg_midea) else f"{avg_midea:.2f}"}</div>
        </div>
        <div style="background:rgba(255,255,255,.12); border:1px solid rgba(255,255,255,.18); border-radius:14px; padding:8px 8px; backdrop-filter:blur(8px); display:flex; flex-direction:column; justify-content:center; align-items:center; text-align:center;">
          <div style="font-size:15px; font-weight:900; color:rgba(255,255,255,.84); letter-spacing:.03em; text-transform:uppercase;">Industry Experience</div>
          <div style="font-size:27px; font-weight:950; color:white; margin-top:4px; line-height:1;">{"-" if np.isnan(avg_industry) else f"{avg_industry:.2f}"}</div>
        </div>
      </div>
    </div>
    """
    components.html(kpi_card, height=316)


def render_gender_distribution():
    gender_dist = (
        employees_core["gender"]
        .fillna("Unknown")
        .astype(str)
        .str.strip()
        .replace({"Female": "F", "Male": "M"})
        .value_counts()
    )
    male_count = int(gender_dist.get("M", 0))
    female_count = int(gender_dist.get("F", 0))
    total_gender = male_count + female_count

    if total_gender > 0:
        male_pct = male_count / total_gender * 100
        female_pct = female_count / total_gender * 100
    else:
        male_pct = 0.0
        female_pct = 0.0
    gender_by_pd = (
        employees_core.assign(
            gender_code=(
                employees_core["gender"]
                .fillna("Unknown")
                .astype(str)
                .str.strip()
                .replace({"Female": "F", "Male": "M"})
            )
        )
        .groupby(["pd", "gender_code"])
        .size()
        .unstack(fill_value=0)
    )
    gender_by_pd = gender_by_pd.reindex(columns=["M", "F"], fill_value=0)
    gender_by_pd = gender_by_pd.loc[gender_by_pd.sum(axis=1).sort_values(ascending=False).index]
    fig_gender, ax_gender = plt.subplots(figsize=(4.95, 3.25), facecolor="#f7fcff")
    ax_gender.set_facecolor("#f7fcff")

    pd_rows = gender_by_pd.index.tolist()
    male_vals = gender_by_pd["M"].values.astype(float) if "M" in gender_by_pd else np.zeros(len(pd_rows))
    female_vals = gender_by_pd["F"].values.astype(float) if "F" in gender_by_pd else np.zeros(len(pd_rows))
    y = np.arange(len(pd_rows))
    max_count = max(1.0, np.max(np.concatenate([male_vals, female_vals])) if len(pd_rows) else 1.0)
    male_plot = -male_vals
    female_plot = female_vals

    bars_m = ax_gender.barh(y, male_plot, color="#0e3a67", height=0.56, label="Male")
    bars_f = ax_gender.barh(y, female_plot, color="#0096db", height=0.56, label="Female")

    for bar, v in zip(bars_m, male_vals):
        if v <= 0:
            continue
        ax_gender.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_y() + bar.get_height() / 2,
            f"{int(v)}",
            ha="center",
            va="center",
            fontsize=9.3,
            fontweight="bold",
            color="white",
        )
    for bar, v in zip(bars_f, female_vals):
        if v <= 0:
            continue
        ax_gender.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_y() + bar.get_height() / 2,
            f"{int(v)}",
            ha="center",
            va="center",
            fontsize=9.3,
            fontweight="bold",
            color="white",
        )

    ax_gender.axvline(0, color="#d7ebf7", linewidth=1.1)
    ax_gender.set_xlim(-(max_count + 2), max_count + 2)
    ax_gender.set_xticks([])
    ax_gender.set_yticks(y)
    ax_gender.set_yticklabels(pd_rows, fontsize=9.8, color="#36536f")
    ax_gender.tick_params(axis="y", length=0, pad=6)
    ax_gender.set_axisbelow(True)
    ax_gender.invert_yaxis()
    ax_gender.spines["top"].set_visible(False)
    ax_gender.spines["right"].set_visible(False)
    ax_gender.spines["left"].set_visible(False)
    ax_gender.spines["bottom"].set_visible(False)
    ax_gender.text(0.25, 1.08, "Male", transform=ax_gender.transAxes, ha="center", va="bottom", fontsize=10.5, fontweight="bold", color="#36536f")
    ax_gender.text(0.75, 1.08, "Female", transform=ax_gender.transAxes, ha="center", va="bottom", fontsize=10.5, fontweight="bold", color="#36536f")
    fig_gender.tight_layout(rect=[0.03, 0.05, 0.99, 0.92])
    gender_b64 = fig_to_base64(fig_gender)
    plt.close(fig_gender)
    gender_card_height = 316

    gender_card = f"""
    <html>
    <head>
    <style>
      html, body {{
        width: 100%;
        max-width: 100%;
      }}
      body {{
        margin: 0;
        background: transparent;
        font-family: "Segoe UI", Arial, sans-serif;
      }}
      .gender-card {{
        width: 100%;
        max-width: 100%;
        border: 1px solid #d6eaf6;
        border-radius: 22px;
        background: linear-gradient(180deg, #ffffff 0%, #f4fbff 100%);
        padding: 8px 12px 8px 12px;
        color: #102f52;
        height: {gender_card_height}px;
        box-sizing: border-box;
        overflow: hidden;
      }}
      .title-row {{
        display:flex;
        justify-content:space-between;
        align-items:flex-start;
        gap:12px;
        margin-bottom: 6px;
      }}
      .gender-title {{
        text-align: left;
        font-size: 22px;
        font-weight: 900;
        color: #0e3a67;
        padding-left: 10px;
        border-left: 4px solid #0096db;
      }}
      .gender-note {{
        font-size: 11px;
        font-weight: 700;
        color: #6a879f;
        white-space: nowrap;
        padding-top: 4px;
      }}
      .gender-layout {{
        display: grid;
        grid-template-columns: 0.36fr 0.64fr;
        gap: 10px;
        align-items: stretch;
        height: calc(100% - 40px);
      }}
      .count-panel {{
        display: flex;
        flex-direction: column;
        gap: 10px;
        justify-content: flex-start;
        padding-top: 6px;
        height: 100%;
      }}
      .count-box {{
        padding: 10px 10px;
        border-radius: 16px;
        background: #f7fcff;
        border: 1px solid #d8ebf7;
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 10px;
        text-align: center;
      }}
      .count-icon {{
        width: 46px;
        height: 68px;
        object-fit: contain;
        flex: 0 0 46px;
      }}
      .count-text {{
        display:flex;
        flex-direction:column;
        align-items:center;
      }}
      .count-label {{
        font-size: 15px;
        font-weight: 900;
        color: #5d7b94;
        text-transform: uppercase;
      }}
      .count-value {{
        margin-top: 4px;
        font-size: 27px;
        font-weight: 950;
        line-height: 1;
      }}
      .count-pct {{
        margin-top: 6px;
        font-size: 13px;
        font-weight: 800;
        line-height: 1;
        color: #5d7b94;
      }}
      .count-box.male .count-value {{ color:#0e3a67; }}
      .count-box.female .count-value {{ color:#0096db; }}
      .chart-wrap {{
        display:flex;
        align-items:flex-start;
        justify-content:center;
        height:100%;
      }}
      .chart-wrap img {{
        width:100%;
        display:block;
      }}
    </style>
    </head>
    <body>
      <div class="gender-card">
        <div class="title-row">
          <div class="gender-title">Gender Distribution</div>
          <div class="gender-note">Excludes coops and contractors</div>
        </div>
        <div class="gender-layout">
          <div class="count-panel">
            <div class="count-box male">
              <img class="count-icon" src="data:image/png;base64,{man_b64}" alt="male icon" />
              <div class="count-text">
                <div class="count-label">Male</div>
                <div class="count-value">{male_count}</div>
                <div class="count-pct">{male_pct:.1f}%</div>
              </div>
            </div>
            <div class="count-box female">
              <img class="count-icon" src="data:image/png;base64,{woman_b64}" alt="female icon" />
              <div class="count-text">
                <div class="count-label">Female</div>
                <div class="count-value">{female_count}</div>
                <div class="count-pct">{female_pct:.1f}%</div>
              </div>
            </div>
          </div>
          <div class="chart-wrap">
            <img src="data:image/png;base64,{gender_b64}" alt="PD gender ratio chart" />
          </div>
        </div>
      </div>
    </body>
    </html>
    """
    components.html(gender_card, height=gender_card_height)

def compute_integer_percentages(counts):
    counts = np.array(counts, dtype=float)
    total = counts.sum()

    raw = counts / total * 100.0
    floors = np.floor(raw).astype(int)
    remainders = raw - floors

    need = 100 - floors.sum()

    if need > 0:
        idx = np.argsort(-remainders)
        floors[idx[:need]] += 1
    elif need < 0:
        idx = np.argsort(remainders)
        floors[idx[:abs(need)]] -= 1

    return floors


def autopct_percent_only(pcts):
    i = {"k": 0}
    def _fmt(_pct):
        val = int(pcts[i["k"]])
        i["k"] += 1
        return f"{val}%"
    return _fmt


def render_pd_analysis():
    pd_dist = (
        employees_core["pd"]
        .fillna("Unknown")
        .astype(str)
        .str.strip()
        .value_counts()
        .sort_values(ascending=False)
    )

    tmp = employees_core.copy()
    tmp["pd"] = tmp["pd"].fillna("Unknown").astype(str).str.strip()
    tmp["employee_class"] = tmp["employee_class"].fillna("Unknown").astype(str)

    pivot = (
        tmp.groupby(["pd", "employee_class"])
        .size()
        .unstack(fill_value=0)
    )
    pivot["__total__"] = pivot.sum(axis=1)
    pivot = pivot.loc[pd_dist.index]

    counts = pd_dist.values
    labels = pd_dist.index.tolist()
    pcts_int = compute_integer_percentages(counts)

    fig_pie, ax_pie = plt.subplots(figsize=(4.35, 4.05))
    fig_pie.patch.set_facecolor("#f7fcff")
    ax_pie.set_facecolor("#f7fcff")
    pie_colors = [
        "#0e3a67",
        "#0096db",
        "#2a78b0",
        "#4db9d6",
        "#7fcfec",
        "#a8dff4",
        "#6e90ab",
        "#d3edf9",
        "#b9d5e7",
    ]

    wedges, texts, autotexts = ax_pie.pie(
        counts,
        labels=labels,
        colors=pie_colors[: len(counts)],
        autopct=autopct_percent_only(pcts_int),
        startangle=90,
        pctdistance=0.72,
        labeldistance=1.12,
        wedgeprops={"edgecolor": "white", "linewidth": 1.5},
        textprops={"fontsize": 10, "color": "#3b5873"},
    )
    ax_pie.axis("equal")

    for t in texts:
        t.set_color("#3b5873")
        t.set_fontsize(10)
    for wedge, t in zip(wedges, autotexts):
        t.set_fontsize(10)
        t.set_weight("bold")
        t.set_color("white" if is_dark_color(wedge.get_facecolor()) else "#0e3a67")

    fig_pie.tight_layout(rect=[0.02, 0.14, 0.98, 0.98])
    pie_b64 = fig_to_base64(fig_pie)
    plt.close(fig_pie)

    col_order = [c for c in ["Local", "Expat", "Unknown"] if c in pivot.columns] + [
        c for c in pivot.columns if c not in {"Local", "Expat", "Unknown", "__total__"}
    ]
    pivot = pivot[col_order]

    fig_bar, ax_bar = plt.subplots(figsize=(5.0, 4.55))
    fig_bar.patch.set_facecolor("#f7fcff")
    ax_bar.set_facecolor("#f7fcff")

    y = np.arange(len(pivot.index))
    left = np.zeros(len(pivot.index))
    stack_colors = {
        "Local": "#0e3a67",
        "Expat": "#0096db",
        "Unknown": "#7fcfec",
    }
    fallback_colors = ["#2a78b0", "#4db9d6", "#a8dff4", "#b5c8d8"]

    for idx, col in enumerate(pivot.columns):
        vals = pivot[col].values
        bar_color = stack_colors.get(col, fallback_colors[idx % len(fallback_colors)])
        bars = ax_bar.barh(
            y,
            vals,
            left=left,
            label=col,
            height=0.92,
            color=bar_color,
            edgecolor="white",
            linewidth=1.0,
        )

        for i, (bar, v) in enumerate(zip(bars, vals)):
            if v > 0:
                if v >= 2:
                    text_color = "white" if is_dark_color(bar_color) else "#102f52"
                    ax_bar.text(
                        left[i] + v / 2,
                        bar.get_y() + bar.get_height() / 2,
                        str(int(v)),
                        ha="center",
                        va="center",
                        fontsize=10.5,
                        fontweight="bold",
                        color=text_color,
                    )
                else:
                    ax_bar.text(
                        left[i] + v + 0.18,
                        bar.get_y() + bar.get_height() / 2,
                        str(int(v)),
                        ha="left",
                        va="center",
                        fontsize=10.5,
                        fontweight="bold",
                        color="#0096db" if col == "Expat" else "#36536f",
                    )
        left += vals

    ax_bar.set_yticks(y)
    ax_bar.set_yticklabels(pivot.index, fontsize=11, color="#36536f")
    ax_bar.tick_params(axis="x", labelsize=10.5, colors="#6b879f")
    ax_bar.grid(axis="x", color="#d7ebf7", linewidth=0.8)
    ax_bar.set_axisbelow(True)
    ax_bar.invert_yaxis()
    ax_bar.spines["top"].set_visible(False)
    ax_bar.spines["right"].set_visible(False)
    ax_bar.spines["left"].set_visible(False)
    ax_bar.spines["bottom"].set_color("#d7ebf7")
    totals = pivot.sum(axis=1).values
    ax_bar.set_xlim(0, totals.max() + 3.2)
    legend = ax_bar.legend(
        loc="lower center",
        bbox_to_anchor=(0.5, 1.02),
        ncol=max(1, len(pivot.columns)),
        fontsize=10.5,
        frameon=False,
    )
    for txt in legend.get_texts():
        txt.set_color("#36536f")

    fig_bar.tight_layout(rect=[0, 0, 1, 0.94])
    bar_b64 = fig_to_base64(fig_bar)
    plt.close(fig_bar)

    pd_card = f"""
    <div style="
        width:100%;
        max-width:100%;
        border:1px solid #d6eaf6;
        border-radius:22px;
        background:linear-gradient(180deg, #ffffff 0%, #f4fbff 100%);
        box-shadow:0 14px 28px rgba(14,58,103,.08);
        padding:10px 12px 22px 12px;
        box-sizing:border-box;
        height:374px;
        min-height:374px;
        overflow:hidden;
        margin-top:0;
    ">
      <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:18px; margin-bottom:8px;">
        <div style="font-size:22px; font-weight:900; color:#0e3a67; padding-left:10px; border-left:4px solid #0096db;">
          PD Analysis
        </div>
        <div style="font-size:11px; font-weight:700; color:#6a879f; white-space:nowrap; padding-top:5px;">
          Excludes coops and contractors
        </div>
      </div>
      <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px; align-items:start;">
        <div>
          <div style="font-size:18px; font-weight:900; color:#0e3a67; margin-bottom:6px; padding-left:10px; border-left:4px solid #25a7de;">PD Distribution</div>
          <img src="data:image/png;base64,{pie_b64}" alt="PD distribution chart" style="width:100%; display:block;" />
        </div>
        <div>
          <div style="font-size:18px; font-weight:900; color:#0e3a67; margin-bottom:6px; padding-left:10px; border-left:4px solid #25a7de;">PD Headcount</div>
          <img src="data:image/png;base64,{bar_b64}" alt="PD headcount mix chart" style="width:100%; display:block;" />
        </div>
      </div>
    </div>
    """
    components.html(pd_card, height=398, scrolling=False)

if False:
    st.markdown('<div class="section-heading">Gender Distribution</div>', unsafe_allow_html=True)

    gender_dist = (
        employees_view["gender"]
        .fillna("Unknown")
        .astype(str)
        .str.strip()
        .replace({"Female": "F", "Male": "M"})
        .value_counts()
    )

    male_count = int(gender_dist.get("M", 0))
    female_count = int(gender_dist.get("F", 0))
    total_gender = male_count + female_count

    if total_gender > 0:
        male_pct = round(male_count / total_gender * 100)
        female_pct = round(female_count / total_gender * 100)
    else:
        male_pct = 0
        female_pct = 0

    # 上半部分：男女图标 + 数字
    icon_col1, icon_col2 = st.columns(2)

    with icon_col1:
        st.markdown(
            f"""
            <div style="text-align:center; padding-top:4px;">
                <div style="
                    font-size:64px;
                    color:#6f8fb3;
                    line-height:1;
                    margin-bottom:10px;
                ">♂</div>
                <div style="
                    display:inline-block;
                    min-width:70px;
                    padding:12px 18px;
                    border-radius:10px;
                    background:#f2f0ed;
                    font-size:24px;
                    font-weight:800;
                    color:#6f8fb3;
                ">
                    {male_count}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with icon_col2:
        st.markdown(
            f"""
            <div style="text-align:center; padding-top:4px;">
                <div style="
                    font-size:64px;
                    color:#c08a8a;
                    line-height:1;
                    margin-bottom:10px;
                ">♀</div>
                <div style="
                    display:inline-block;
                    min-width:70px;
                    padding:12px 18px;
                    border-radius:10px;
                    background:#f2f0ed;
                    font-size:24px;
                    font-weight:800;
                    color:#c08a8a;
                ">
                    {female_count}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height:18px;'></div>", unsafe_allow_html=True)

    # 下半部分：donut chart
    fig, ax = plt.subplots(figsize=(4.2, 3.2))

    values = [male_count, female_count]
    colors = ["#7c96b8", "#c79aa5"]   # 莫兰迪蓝 / 莫兰迪粉

    colors = ["#0e3a67", "#0096db"]
    wedges, texts, autotexts = ax.pie(
        values,
        labels=None,
        colors=colors,
        startangle=90,
        counterclock=False,
        wedgeprops=dict(width=0.36, edgecolor="white", linewidth=1.5),
        autopct=lambda p: f"{int(round(p))}%" if p > 0 else "",
        pctdistance=1.12,
        textprops={"fontsize": 10, "fontweight": "bold", "color": "#36536f"},
    )

    fig.patch.set_facecolor("#f7fcff")
    ax.set_facecolor("#f7fcff")
    ax.axis("equal")
    fig.tight_layout()
    st.pyplot(fig, clear_figure=True)


# =============================
# Row2 Right: New Hire + Service Anniversary
# =============================
if False:

    st.markdown(
        """
        <style>
        .block-card{
          border:1px solid #d6eaf6;
          border-radius:18px;
          padding:20px 22px;
          background:linear-gradient(180deg, #ffffff 0%, #f4fbff 100%);
          box-shadow: 0 14px 28px rgba(14,58,103,.08);
          margin-bottom:20px;
        }
        .tag-container{
          display:flex;
          flex-wrap:wrap;
          gap:10px;
        }
        .tag{
          border:1px solid #d8ebf7;
          background:#eaf7fd;
          border-radius:999px;
          padding:8px 12px;
          font-size:18px;
          font-weight:700;
          color: #1d4f77;
          line-height:1;
          white-space:nowrap;
        }
        .svc-row{
          display:flex;
          align-items:flex-start;
          gap:16px;
          margin-bottom:14px;
        }
        .svc-row:last-child{ margin-bottom:0; }
        .svc-year{
          min-width:110px;
          font-weight:900;
          font-size:18px;
          color: #0e3a67;
          padding-top:4px;
        }
        .section-title{
          font-size:21px;
          font-weight:900;
          color:#0e3a67;
          padding-left:14px;
          border-left:4px solid #0096db;
          margin-bottom:12px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(f'<div class="section-title">New Hire ({report_period_label})</div>', unsafe_allow_html=True)
    nh = new_hires_in_month(employees_core, as_of)

    if nh.empty:
        st.markdown(
            f"""
            <div class="block-card">
              <div style="color:rgba(0,0,0,.65); font-weight:700; font-size:18px;">
                No new hires through {report_as_of_label}.
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        nh_names = sorted(nh["name"].astype(str).str.strip().tolist())
        tags_html = "".join([f'<span class="tag">{n}</span>' for n in nh_names])
        st.markdown(
            f"""
            <div class="block-card">
              <div class="tag-container">
                {tags_html}
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown(f'<div class="section-title">Service Anniversary ({report_period_label})</div>', unsafe_allow_html=True)
    an = service_anniversaries_in_month(employees_core, as_of)

    if an.empty:
        st.markdown(
            f"""
            <div class="block-card">
              <div style="color:rgba(0,0,0,.65); font-weight:700; font-size:18px;">
                No service anniversaries through {report_as_of_label}.
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        an2 = an[["anniv_years", "name"]].copy()
        an2["name"] = an2["name"].astype(str).str.strip()

        grp = (
            an2.groupby("anniv_years")["name"]
            .apply(lambda s: sorted(s.tolist()))
            .sort_index(ascending=False)
        )

        rows_html = ""
        for years, names in grp.items():
            tags_html = "".join([f'<span class="tag">{n}</span>' for n in names])
            label = f"{int(years)} Year" + ("s" if int(years) > 1 else "")
            rows_html += f"""
              <div class="svc-row">
                <div class="svc-year">{label}</div>
                <div class="tag-container">
                  {tags_html}
                </div>
              </div>
            """

        st.markdown(
            f"""
            <div class="block-card">
              {rows_html}
            </div>
            """,
            unsafe_allow_html=True,
        )

def render_coop_overview():
    coop_semester_label = get_semester_label(as_of)
    coop_overview_title = (
        f"{coop_semester_label} Co-op Overview"
        if coop_semester_label
        else "Co-op Overview"
    )
    coop_count = int(len(employees_coop))
    coop_male_count = int(
        employees_coop["gender"]
        .fillna("Unknown")
        .astype(str)
        .str.strip()
        .replace({"Male": "M", "Female": "F"})
        .eq("M")
        .sum()
    )
    coop_female_count = int(
        employees_coop["gender"]
        .fillna("Unknown")
        .astype(str)
        .str.strip()
        .replace({"Male": "M", "Female": "F"})
        .eq("F")
        .sum()
    )
    coop_gender_total = coop_male_count + coop_female_count
    coop_male_pct = (coop_male_count / coop_gender_total * 100) if coop_gender_total else 0.0
    coop_female_pct = (coop_female_count / coop_gender_total * 100) if coop_gender_total else 0.0

    all_pd_labels = (
        employees_view["pd"]
        .fillna("Unknown")
        .astype(str)
        .str.strip()
        .value_counts()
        .index
        .tolist()
    )

    coop_pd_counts = (
        employees_coop["pd"]
        .fillna("Unknown")
        .astype(str)
        .str.strip()
        .value_counts()
        .reindex(all_pd_labels, fill_value=0)
        .sort_values(ascending=False)
    )
    coop_pd_counts = coop_pd_counts[coop_pd_counts > 0]

    coop_chart_rows = max(1, len(coop_pd_counts.index))
    coop_fig_height = 2.85 if coop_chart_rows <= 6 else min(4.1, 2.85 + (coop_chart_rows - 6) * 0.18)

    fig_coop_bar, ax_coop_bar = plt.subplots(figsize=(5.6, coop_fig_height))
    fig_coop_bar.patch.set_facecolor("#f7fcff")
    ax_coop_bar.set_facecolor("#f7fcff")

    x = np.arange(len(coop_pd_counts.index))
    bars = ax_coop_bar.bar(
        x,
        coop_pd_counts.values,
        width=0.62,
        color="#0e3a67",
        edgecolor="white",
        linewidth=1.0,
    )

    for bar, v in zip(bars, coop_pd_counts.values):
        if v > 0:
            ax_coop_bar.text(
                bar.get_x() + bar.get_width() / 2,
                bar.get_height() + 0.08,
                str(int(v)),
                ha="center",
                va="bottom",
                fontsize=10.5,
                fontweight="bold",
                color="#0e3a67",
            )

    ax_coop_bar.set_xticks(x)
    ax_coop_bar.set_xticklabels(coop_pd_counts.index, fontsize=10.5, color="#36536f")
    ax_coop_bar.tick_params(axis="x", labelsize=10.5, colors="#36536f")
    ax_coop_bar.tick_params(axis="y", length=0, labelleft=False, colors="#6b879f")
    ax_coop_bar.grid(axis="y", color="#d7ebf7", linewidth=0.8)
    ax_coop_bar.set_axisbelow(True)
    ax_coop_bar.spines["top"].set_visible(False)
    ax_coop_bar.spines["right"].set_visible(False)
    ax_coop_bar.spines["left"].set_color("#d7ebf7")
    ax_coop_bar.spines["bottom"].set_color("#d7ebf7")
    coop_max = int(coop_pd_counts.max()) if len(coop_pd_counts) else 1
    ax_coop_bar.set_ylim(0, coop_max + 0.7)
    ax_coop_bar.set_yticks(np.arange(0, coop_max + 1, 1))
    fig_coop_bar.tight_layout(rect=[0, 0, 1, 0.99])
    coop_bar_b64 = fig_to_base64(fig_coop_bar)
    plt.close(fig_coop_bar)

    coop_overview_card = f"""
<html>
<head>
<style>
  html, body {{
    width:100%;
    max-width:100%;
    overflow:hidden;
  }}
  body {{
    margin:0;
    background:transparent;
    font-family:"Segoe UI", Arial, sans-serif;
  }}
  .coop-card {{
    height: 374px;
    min-height: 374px;
    width:100%;
    max-width:100%;
    border:1px solid #d6eaf6;
    border-radius:22px;
    background:linear-gradient(180deg, #ffffff 0%, #f4fbff 100%);
    box-shadow:0 14px 28px rgba(14,58,103,.08);
    padding:10px 12px 22px 12px;
    box-sizing:border-box;
    overflow:hidden;
  }}
  .coop-head {{
    display:flex;
    justify-content:space-between;
    align-items:flex-start;
    gap:18px;
    margin-bottom:8px;
  }}
  .coop-title {{
    font-size:22px;
    font-weight:900;
    color:#0e3a67;
    padding-left:10px;
    border-left:4px solid #0096db;
    line-height:1.1;
  }}
  .top-grid {{
    display:grid;
    grid-template-columns:1.14fr 0.86fr;
    gap:14px;
    align-items:stretch;
    height: calc(100% - 46px);
  }}
  .mini-card {{
    border:1px solid #d8ebf7;
    border-radius:18px;
    background:#ffffff;
    padding:16px 16px 14px 16px;
    box-sizing:border-box;
    height:100%;
    min-height:0;
    display:flex;
    flex-direction:column;
  }}
  .mini-title {{
    font-size:20px;
    font-weight:900;
    color:#0e3a67;
    padding-left:10px;
    border-left:4px solid #25a7de;
    margin-bottom:0;
  }}
  .mini-head {{
    min-height:86px;
    display:flex;
    justify-content:space-between;
    align-items:flex-start;
    gap:14px;
    margin-bottom:12px;
  }}
  .donut-shell {{
    flex:1;
    display:flex;
    align-items:center;
    justify-content:center;
  }}
  .donut-wrap {{
    position:relative;
    display:flex;
    justify-content:center;
    align-items:center;
    min-height:128px;
    width:100%;
  }}
  .donut {{
    width:132px;
    height:132px;
    border-radius:50%;
    background:conic-gradient(#0e3a67 0 {coop_male_pct:.3f}%, #0096db {coop_male_pct:.3f}% 100%);
    position:relative;
  }}
  .donut::after {{
    content:"";
    position:absolute;
    inset:34px;
    background:#ffffff;
    border-radius:50%;
    box-shadow:inset 0 0 0 1px #e1eff8;
  }}
  .ring-count {{
    position:absolute;
    z-index:2;
    color:#ffffff;
    font-size:18px;
    font-weight:950;
    line-height:1;
    text-shadow:0 1px 3px rgba(14,58,103,.35);
  }}
  .ring-count.male {{
    top:50%;
    right:12px;
    transform:translateY(-50%);
  }}
  .ring-count.female {{
    top:50%;
    left:12px;
    transform:translateY(-50%);
  }}
  .pct {{
    position:absolute;
    font-size:14px;
    font-weight:900;
    white-space:nowrap;
    line-height:1;
  }}
  .pct.male {{
    color:#0e3a67;
    top:30px;
    left:calc(50% + 70px);
  }}
  .pct.female {{
    color:#0096db;
    bottom:30px;
    right:calc(50% + 70px);
    text-align:right;
  }}
  .donut-center {{
    position:absolute;
    inset:0;
    z-index:1;
    display:flex;
    align-items:center;
    justify-content:center;
    flex-direction:column;
    text-align:center;
  }}
  .center-counts {{
    display:none;
  }}
  .chart-wrap {{
    background:#f7fcff;
    border-radius:16px;
    padding:10px 12px 4px 12px;
  }}
  .chart-head {{
    display:flex;
    align-items:flex-start;
    justify-content:space-between;
    gap:14px;
    margin-bottom:0;
  }}
  .chart-head .mini-title {{
    margin-bottom:0;
  }}
  .pd-chart-wrap {{
    flex:1;
    display:flex;
    flex-direction:column;
    justify-content:flex-start;
  }}
  .pd-chart-wrap .chart-wrap {{
    flex:none;
  }}
  .headcount-chip {{
    min-width:110px;
    padding:10px 12px;
    border-radius:16px;
    background:#f7fcff;
    border:1px solid #d8ebf7;
    text-align:center;
  }}
  .headcount-chip .chip-label {{
    font-size:12px;
    font-weight:900;
    color:#6b879f;
    text-transform:uppercase;
    letter-spacing:.04em;
  }}
  .headcount-chip .chip-value {{
    margin-top:4px;
    font-size:34px;
    font-weight:950;
    color:#0e3a67;
    line-height:1;
  }}
  .headcount-chip.ghost {{
    visibility:hidden;
  }}
</style>
</head>
<body>
  <div class="coop-card">
    <div class="coop-head">
      <div class="coop-title">{coop_overview_title}</div>
    </div>
    <div class="top-grid">
      <div class="mini-card">
        <div class="mini-head chart-head">
          <div class="mini-title">PD Headcount</div>
          <div class="headcount-chip">
            <div class="chip-label">Co-ops</div>
            <div class="chip-value">{coop_count}</div>
          </div>
        </div>
        <div class="pd-chart-wrap">
          <div class="chart-wrap">
            <img src="data:image/png;base64,{coop_bar_b64}" alt="Co-op PD headcount chart" style="width:100%; display:block;" />
          </div>
        </div>
      </div>
      <div class="mini-card">
        <div class="mini-head">
          <div class="mini-title">Gender Distribution</div>
          <div class="headcount-chip ghost" aria-hidden="true">
            <div class="chip-label">Co-ops</div>
            <div class="chip-value">{coop_count}</div>
          </div>
        </div>
        <div class="donut-shell">
          <div class="donut-wrap">
            <div class="donut">
              <div class="ring-count male">{coop_male_count}</div>
              <div class="ring-count female">{coop_female_count}</div>
              <div class="donut-center">
                <div class="center-counts">
                  <div>{coop_male_count} M</div>
                  <div>{coop_female_count} F</div>
                </div>
              </div>
            </div>
            <div class="pct male">Male {coop_male_pct:.1f}%</div>
            <div class="pct female">Female {coop_female_pct:.1f}%</div>
          </div>
        </div>
      </div>
    </div>
  </div>
</body>
</html>
"""
    components.html(coop_overview_card, height=398, scrolling=False)


# =============================
# Dashboard Rows
# =============================
st.markdown("<div style='margin-top:-112px;'></div>", unsafe_allow_html=True)

row1_left, row1_right = st.columns([1, 1], gap="small")
with row1_left:
    render_workforce_snapshot()
with row1_right:
    render_gender_distribution()

st.markdown("<div style='margin-top:8px;'></div>", unsafe_allow_html=True)

row2_left, row2_right = st.columns([1, 1], gap="small")
with row2_left:
    render_pd_analysis()
with row2_right:
    render_coop_overview()

st.markdown("<div style='margin-top:8px;'></div>", unsafe_allow_html=True)

row3_left, row3_right = st.columns([1, 1], gap="small")
with row3_left:
    render_new_hires_card()
with row3_right:
    render_service_anniversary_card()

st.divider()

# =============================
# Org Chart (Executive -> PDs -> PD Reporting Tree) ✅ no d3
# - Header: Executive Director + Associate Executive Director + PD bus lines
# - Detail: Selected PD reporting tree (manager -> direct reports)
# - Interactions: click PD to show/hide, wheel-zoom, drag-pan, Fit, Reset
# =============================
import json
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

st.markdown('<div class="section-heading">Org Chart</div>', unsafe_allow_html=True)

# ---------- Build org data (uses name/manager/job/pd already in employees_view) ----------

def build_org_payload(df: pd.DataFrame):
    df = df.copy()
    df["name"] = df["name"].fillna("").astype(str).str.strip()
    df["manager"] = df["manager"].fillna("").astype(str).str.strip()
    df["job"] = df["job"].fillna("").astype(str).str.strip()
    df["pd"] = df["pd"].fillna("Unknown").astype(str).str.strip()

    def norm(s):
        return str(s).strip().lower()

    # Executive Director
    def is_exec_job(job: str) -> bool:
        t = norm(job)
        return (
            "executive director" in t
            or t == "exec director"
            or "exec. director" in t
        )

    # Associate Executive Director
    def is_assoc_job(job: str) -> bool:
        t = norm(job)
        return (
            "associate executive director" in t
            or "associate exec director" in t
            or "assoc executive director" in t
            or "associate executive" in t
        )

    def is_top_leader(job: str) -> bool:
        return is_exec_job(job) or is_assoc_job(job)

    # Top header names
    exec_names = sorted(
        df.loc[df["job"].apply(is_exec_job), "name"].dropna().unique().tolist()
    )
    assoc_names = sorted(
        df.loc[df["job"].apply(is_assoc_job), "name"].dropna().unique().tolist()
    )

    exec_primary = exec_names[0] if exec_names else "N/A"
    assoc_primary = assoc_names[0] if assoc_names else "N/A"

    # PD list still keeps all PDs
    pd_names = sorted(df["pd"].dropna().unique().tolist())

    pd_map = {}
    for pd_name in pd_names:
        g = df[df["pd"] == pd_name].copy()
        g = g[g["name"].ne("")].copy()

        # ✅ Remove Executive / Associate Executive from PD detail tree
        g = g[~g["job"].apply(is_top_leader)].copy()

        if g.empty:
            pd_map[pd_name] = {"nodes": [], "edges": [], "roots": [], "title": pd_name}
            continue

        name_to_id = {}
        nodes = []
        for r in g[["name", "job", "manager"]].itertuples(index=False):
            nm, jb, mgr = r
            _id = nm
            name_to_id[nm] = _id
            nodes.append({"id": _id, "name": nm, "job": jb, "manager": mgr})

        edges = []
        for nd in nodes:
            mgr = nd["manager"]
            if mgr and (mgr in name_to_id) and (mgr != nd["name"]):
                edges.append({"source": name_to_id[mgr], "target": nd["id"]})

        has_parent = set(e["target"] for e in edges)
        roots = [nd["id"] for nd in nodes if nd["id"] not in has_parent]
        if not roots and nodes:
            roots = [nodes[0]["id"]]

        pd_map[pd_name] = {
            "title": pd_name,
            "nodes": nodes,
            "edges": edges,
            "roots": roots
        }

    return {
        "exec_name": exec_primary,
        "assoc_name": assoc_primary,
        "pds": pd_names,
        "pd_map": pd_map
    }

org_payload = build_org_payload(employees_core)
org_json = json.dumps(org_payload, ensure_ascii=False)


browser_html = r"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<style>
  :root{
    --border:#d6eaf6;
    --muted:#5d7b94;
    --shadow:0 14px 28px rgba(14,58,103,.10);
    --accent:#0096db;
    --accent-dark:#0e3a67;
    --line:#a8d8ef;
  }
  body{
    margin:0;
    padding:0;
    font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,"Noto Sans","Helvetica Neue",sans-serif;
    background:linear-gradient(180deg, #fbfeff 0%, #f3faff 100%);
  }
  .wrap{
    border:1px solid var(--border);
    border-radius:18px;
    overflow:hidden;
    background:#fff;
    box-shadow:var(--shadow);
  }
  .topbar{
    display:flex;
    justify-content:space-between;
    align-items:center;
    gap:14px;
    padding:14px 16px;
    background:linear-gradient(90deg, #eaf7fd 0%, #f8fcff 100%);
    border-bottom:1px solid var(--border);
  }
  .topbar-title{
    font-size:18px;
    font-weight:950;
    color:var(--accent-dark);
  }
  .topbar-note{
    font-size:12px;
    font-weight:800;
    color:var(--muted);
  }
  .content{
    padding:16px;
  }
  .leadership{
    display:grid;
    grid-template-columns:repeat(auto-fit, minmax(280px, 1fr));
    gap:14px;
    margin-bottom:16px;
  }
  .leader-card{
    border:1px solid var(--border);
    border-radius:18px;
    background:linear-gradient(180deg, #ffffff 0%, #f5fbff 100%);
    padding:18px 20px;
    box-shadow:0 10px 22px rgba(14,58,103,.08);
  }
  .leader-role{
    font-size:12px;
    font-weight:900;
    text-transform:uppercase;
    letter-spacing:.08em;
    color:#6c8aa3;
    margin-bottom:8px;
  }
  .leader-name{
    font-size:34px;
    font-weight:950;
    line-height:1.04;
    color:var(--accent-dark);
  }
  .pdStrip{
    display:flex;
    gap:10px;
    flex-wrap:wrap;
    margin-bottom:16px;
  }
  .pdChip{
    border:1px solid #bfe2f7;
    background:linear-gradient(180deg, #ffffff 0%, #edf8fe 100%);
    border-radius:999px;
    padding:10px 16px;
    font-weight:900;
    font-size:16px;
    color:var(--accent-dark);
    cursor:pointer;
    user-select:none;
    transition:all .18s ease;
  }
  .pdChip:hover{
    transform:translateY(-1px);
    box-shadow:0 8px 18px rgba(14,58,103,.08);
  }
  .pdChip.active{
    background:linear-gradient(180deg, #0e3a67 0%, #1189cb 100%);
    border-color:#0e3a67;
    color:#ffffff;
    box-shadow:0 12px 22px rgba(14,58,103,.18);
  }
  .chartCard{
    border:1px solid var(--border);
    border-radius:18px;
    background:linear-gradient(180deg, #ffffff 0%, #f7fcff 100%);
    box-shadow:0 12px 24px rgba(14,58,103,.08);
    padding:18px;
  }
  .chartTools{
    display:flex;
    justify-content:flex-end;
    align-items:center;
    gap:10px;
    margin-bottom:12px;
  }
  .chartHint{
    font-size:12px;
    font-weight:800;
    color:var(--muted);
  }
  .toolBtn{
    border:1px solid #cfe5f3;
    background:#fff;
    border-radius:10px;
    padding:6px 10px;
    font-size:12px;
    font-weight:900;
    color:var(--accent-dark);
    cursor:pointer;
    user-select:none;
  }
  .toolBtn:hover{
    background:#eef8fd;
  }
  .chartHead{
    display:flex;
    justify-content:space-between;
    align-items:flex-start;
    gap:16px;
    margin-bottom:16px;
  }
  .chartTitle{
    font-size:34px;
    font-weight:950;
    color:var(--accent-dark);
    line-height:1.02;
  }
  .chartSub{
    margin-top:8px;
    color:var(--muted);
    font-size:14px;
    font-weight:800;
  }
  .metaRow{
    display:flex;
    flex-wrap:wrap;
    gap:8px;
    justify-content:flex-end;
  }
  .metaPill{
    border:1px solid #d2e8f6;
    background:#eef8fd;
    color:#1f5a86;
    border-radius:999px;
    padding:8px 12px;
    font-size:13px;
    font-weight:900;
  }
  .chartViewport{
    border:1px solid var(--border);
    border-radius:16px;
    background:linear-gradient(180deg, #fbfeff 0%, #f5fbff 100%);
    overflow:hidden;
    height:640px;
    position:relative;
  }
  #chartSvg{
    width:100%;
    height:100%;
    display:block;
    background:transparent;
  }
  .node{
    fill:#fbfeff;
    stroke:#2a78b0;
    stroke-width:1.6;
    rx:16;
    ry:16;
    filter: drop-shadow(0 8px 18px rgba(14,58,103,.08));
  }
  .rootNode{
    fill:#eef8fd;
    stroke:#0096db;
    stroke-width:2.2;
  }
  .nodeTitle{
    fill:#0e3a67;
    font-weight:950;
    text-anchor:middle;
  }
  .nodeSub{
    fill:#5d7b94;
    font-weight:800;
    text-anchor:middle;
  }
  .edge{
    stroke:var(--line);
    stroke-width:2.2;
    fill:none;
    stroke-linecap:round;
    stroke-linejoin:round;
  }
  .empty{
    min-height:240px;
    display:flex;
    align-items:center;
    justify-content:center;
    text-align:center;
    color:#6a879f;
    font-size:18px;
    font-weight:800;
    background:linear-gradient(180deg, #fbfeff 0%, #f5fbff 100%);
    border:1px dashed #cfe5f3;
    border-radius:16px;
  }
</style>
</head>
<body>
  <div class="wrap">
    <div class="topbar">
      <div class="topbar-title">Org Chart</div>
      <div class="topbar-note">Click a PD name to expand. Wheel=Zoom, Drag=Pan</div>
    </div>
    <div class="content">
      <div class="leadership" id="leadership"></div>
      <div class="pdStrip" id="pdStrip"></div>
      <div class="chartCard">
        <div class="chartHead">
          <div>
            <div class="chartTitle" id="chartTitle">PD</div>
            <div class="chartSub">Interactive org chart by PD</div>
          </div>
          <div class="metaRow" id="chartMeta"></div>
        </div>
        <div class="chartTools">
          <div class="chartHint">Wheel=Zoom, Drag=Pan</div>
          <div class="toolBtn" id="btnFit">Fit</div>
          <div class="toolBtn" id="btnReset">100%</div>
        </div>
        <div class="chartViewport" id="chartViewport">
          <div id="chartWrap">
            <div class="empty">Choose a PD to view the org chart.</div>
          </div>
        </div>
      </div>
    </div>
  </div>

<script>
  const ORG = __ORG_JSON__;
  const pds = ORG.pds || [];
  const leadershipEl = document.getElementById("leadership");
  const pdStripEl = document.getElementById("pdStrip");
  const chartWrapEl = document.getElementById("chartWrap");
  const chartViewportEl = document.getElementById("chartViewport");
  const chartTitleEl = document.getElementById("chartTitle");
  const chartMetaEl = document.getElementById("chartMeta");
  const btnFit = document.getElementById("btnFit");
  const btnReset = document.getElementById("btnReset");
  let activePD = "";
  let zoomState = { svg: null, zoomLayer: null, width: 0, height: 0, tx: 0, ty: 0, scale: 1 };

  function escapeHtml(text){
    return String(text || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function svgEl(tag){
    return document.createElementNS("http://www.w3.org/2000/svg", tag);
  }

  function renderLeadership(){
    leadershipEl.innerHTML = `
      <div class="leader-card">
        <div class="leader-role">Executive Director</div>
        <div class="leader-name">${escapeHtml(ORG.exec_name || "N/A")}</div>
      </div>
      <div class="leader-card">
        <div class="leader-role">Associate Executive Director</div>
        <div class="leader-name">${escapeHtml(ORG.assoc_name || "N/A")}</div>
      </div>
    `;
  }

  function buildAdj(pdData){
    const nodes = pdData.nodes || [];
    const edges = pdData.edges || [];
    const children = new Map();
    const idToNode = new Map();
    const parent = new Map();

    nodes.forEach(n => {
      children.set(n.id, []);
      idToNode.set(n.id, n);
    });

    edges.forEach(e => {
      if(children.has(e.source) && children.has(e.target)){
        children.get(e.source).push(e.target);
        parent.set(e.target, e.source);
      }
    });

    children.forEach(arr => {
      arr.sort((a, b) => (idToNode.get(a)?.name || "").localeCompare(idToNode.get(b)?.name || ""));
    });

    return { children, idToNode, parent };
  }

  function getLevels(pdData){
    const { children, idToNode, parent } = buildAdj(pdData);
    let roots = (pdData.roots || []).filter(id => idToNode.has(id));
    if(!roots.length){
      roots = (pdData.nodes || []).map(n => n.id).filter(id => !parent.has(id));
    }
    roots.sort((a, b) => (idToNode.get(a)?.name || "").localeCompare(idToNode.get(b)?.name || ""));

    const levels = [];
    let current = roots.slice();
    const seen = new Set();

    while(current.length){
      const unique = current.filter((id, idx, arr) => arr.indexOf(id) === idx && !seen.has(id));
      if(!unique.length) break;
      unique.forEach(id => seen.add(id));
      levels.push(unique);
      const next = [];
      unique.forEach(id => (children.get(id) || []).forEach(child => next.push(child)));
      current = next;
    }

    return { levels, children, idToNode, parent };
  }

  function trimLabel(text, maxChars){
    const clean = String(text || "").trim();
    if(clean.length <= maxChars) return clean;
    return clean.slice(0, Math.max(0, maxChars - 1)).trim() + "…";
  }

  function renderPDStrip(){
    pdStripEl.innerHTML = pds.map(pd => `
      <div class="pdChip ${pd === activePD ? "active" : ""}" data-pd="${escapeHtml(pd)}">${escapeHtml(pd)}</div>
    `).join("");

    pdStripEl.querySelectorAll(".pdChip").forEach(el => {
      el.addEventListener("click", () => setActive(el.getAttribute("data-pd") || ""));
    });
  }

  function renderChart(pdName){
    const pdData = ORG.pd_map[pdName];
    if(!pdData || !pdData.nodes || !pdData.nodes.length){
      chartWrapEl.innerHTML = '<div class="empty">No reporting data available for this PD.</div>';
      chartMetaEl.innerHTML = '';
      chartTitleEl.textContent = pdName || 'PD';
      return;
    }

    const { levels, children, idToNode, parent } = getLevels(pdData);
    const totalNodes = (pdData.nodes || []).length;
    const managers = (pdData.nodes || []).filter(n => (children.get(n.id) || []).length > 0).length;
    const maxRow = levels.reduce((m, row) => Math.max(m, row.length), 1);

    chartTitleEl.textContent = pdName;
    chartMetaEl.innerHTML = `
      <div class="metaPill">${totalNodes} employees</div>
      <div class="metaPill">${managers} managers</div>
      <div class="metaPill">${levels.length} levels</div>
    `;

    const wrapWidth = chartViewportEl.clientWidth || chartWrapEl.clientWidth || 1200;
    const sidePad = 22;
    const topPad = 18;
    const bottomPad = 18;
    const rowGap = maxRow <= 4 ? 132 : maxRow <= 8 ? 116 : 102;
    const usableWidth = wrapWidth - sidePad * 2;

    const profile = maxRow <= 4
      ? { maxW: 248, minW: 172, h: 72, title: 18, sub: 12.5, nameChars: 20, jobChars: 26 }
      : maxRow <= 7
        ? { maxW: 220, minW: 138, h: 66, title: 15, sub: 11, nameChars: 18, jobChars: 23 }
        : { maxW: 188, minW: 116, h: 60, title: 13.5, sub: 10, nameChars: 15, jobChars: 19 };

    const rowWidths = levels.map(row => {
      const count = row.length || 1;
      return Math.max(profile.minW, Math.min(profile.maxW, (usableWidth - (count - 1) * 14) / count));
    });
    const nodeW = Math.min(profile.maxW, Math.max(profile.minW, Math.min(...rowWidths)));
    const nodeH = profile.h;

    const positions = new Map();
    levels.forEach((row, depth) => {
      const count = row.length || 1;
      const rowTop = topPad + depth * rowGap;
      const step = count === 1 ? 0 : (usableWidth - nodeW) / (count - 1);
      row.forEach((id, idx) => {
        const x = sidePad + (count === 1 ? (usableWidth - nodeW) / 2 : idx * step);
        positions.set(id, { x, y: rowTop });
      });
    });

    const svg = svgEl("svg");
    const totalHeight = topPad + (levels.length - 1) * rowGap + nodeH + bottomPad;
    svg.setAttribute("id", "chartSvg");
    svg.setAttribute("viewBox", `0 0 ${wrapWidth} ${totalHeight}`);
    svg.setAttribute("preserveAspectRatio", "xMidYMid meet");
    const zoomLayer = svgEl("g");
    svg.appendChild(zoomLayer);

    (pdData.edges || []).forEach(edge => {
      const p = positions.get(edge.source);
      const c = positions.get(edge.target);
      if(!p || !c) return;
      const x1 = p.x + nodeW / 2;
      const y1 = p.y + nodeH;
      const x2 = c.x + nodeW / 2;
      const y2 = c.y;
      const midY = y1 + (y2 - y1) / 2;
      const path = svgEl("path");
      path.setAttribute("class", "edge");
      path.setAttribute("d", `M ${x1} ${y1} L ${x1} ${midY} L ${x2} ${midY} L ${x2} ${y2}`);
      zoomLayer.appendChild(path);
    });

    levels.forEach((row, depth) => {
      row.forEach(id => {
        const nd = idToNode.get(id);
        const pos = positions.get(id);
        if(!nd || !pos) return;

        const g = svgEl("g");
        const rect = svgEl("rect");
        rect.setAttribute("x", pos.x);
        rect.setAttribute("y", pos.y);
        rect.setAttribute("width", nodeW);
        rect.setAttribute("height", nodeH);
        rect.setAttribute("class", `node ${depth === 0 ? "rootNode" : ""}`);
        g.appendChild(rect);

        const title = svgEl("title");
        const manager = parent.get(id) ? ` | Reports to ${idToNode.get(parent.get(id))?.name || ""}` : "";
        title.textContent = `${nd.name || ""} | ${nd.job || ""}${manager}`;
        g.appendChild(title);

        const nameText = svgEl("text");
        nameText.setAttribute("x", pos.x + nodeW / 2);
        nameText.setAttribute("y", pos.y + nodeH * 0.40);
        nameText.setAttribute("class", "nodeTitle");
        nameText.setAttribute("font-size", String(profile.title));
        nameText.textContent = trimLabel(nd.name || "", profile.nameChars);
        g.appendChild(nameText);

        const jobText = svgEl("text");
        jobText.setAttribute("x", pos.x + nodeW / 2);
        jobText.setAttribute("y", pos.y + nodeH * 0.72);
        jobText.setAttribute("class", "nodeSub");
        jobText.setAttribute("font-size", String(profile.sub));
        jobText.textContent = trimLabel(nd.job || "", profile.jobChars);
        g.appendChild(jobText);

        zoomLayer.appendChild(g);
      });
    });

    chartWrapEl.innerHTML = "";
    chartWrapEl.appendChild(svg);
    attachZoom(svg, zoomLayer, wrapWidth, totalHeight);
    fitChart();
  }

  function setActive(pd){
    activePD = pd;
    renderPDStrip();
    renderChart(pd);
  }

  function applyTransform(){
    if(!zoomState.zoomLayer) return;
    zoomState.zoomLayer.setAttribute("transform", `translate(${zoomState.tx},${zoomState.ty}) scale(${zoomState.scale})`);
  }

  function fitChart(){
    if(!zoomState.svg) return;
    const vw = chartViewportEl.clientWidth || 1200;
    const vh = chartViewportEl.clientHeight || 640;
    const pad = 22;
    const scaleX = (vw - pad * 2) / Math.max(1, zoomState.width);
    const scaleY = (vh - pad * 2) / Math.max(1, zoomState.height);
    zoomState.scale = Math.max(0.35, Math.min(1.18, Math.min(scaleX, scaleY)));
    zoomState.tx = pad + (vw - pad * 2 - zoomState.width * zoomState.scale) / 2;
    zoomState.ty = pad + (vh - pad * 2 - zoomState.height * zoomState.scale) / 2;
    applyTransform();
  }

  function resetChartZoom(){
    if(!zoomState.svg) return;
    zoomState.scale = 1;
    zoomState.tx = 24;
    zoomState.ty = 24;
    applyTransform();
  }

  function attachZoom(svg, zoomLayer, chartWidth, chartHeight){
    zoomState = { svg, zoomLayer, width: chartWidth, height: chartHeight, tx: 0, ty: 0, scale: 1 };
    let dragging = false;
    let lastX = 0;
    let lastY = 0;

    svg.addEventListener("wheel", (e) => {
      e.preventDefault();
      const factor = e.deltaY < 0 ? 1.10 : 0.92;
      const newScale = Math.min(3.2, Math.max(0.35, zoomState.scale * factor));
      const pt = svg.createSVGPoint();
      pt.x = e.clientX;
      pt.y = e.clientY;
      const ctm = svg.getScreenCTM();
      if(!ctm) return;
      const p = pt.matrixTransform(ctm.inverse());
      const x0 = (p.x - zoomState.tx) / zoomState.scale;
      const y0 = (p.y - zoomState.ty) / zoomState.scale;
      zoomState.scale = newScale;
      zoomState.tx = p.x - x0 * zoomState.scale;
      zoomState.ty = p.y - y0 * zoomState.scale;
      applyTransform();
    }, { passive:false });

    svg.addEventListener("mousedown", (e) => {
      dragging = true;
      lastX = e.clientX;
      lastY = e.clientY;
    });

    window.addEventListener("mouseup", () => {
      dragging = false;
    });

    window.addEventListener("mousemove", (e) => {
      if(!dragging || !zoomState.svg) return;
      const dx = e.clientX - lastX;
      const dy = e.clientY - lastY;
      lastX = e.clientX;
      lastY = e.clientY;
      zoomState.tx += dx;
      zoomState.ty += dy;
      applyTransform();
    });
  }

  function init(){
    renderLeadership();
    renderPDStrip();
    if(pds.length){
      setActive(pds[0]);
    }
  }

  window.addEventListener("resize", () => {
    if(activePD){
      renderChart(activePD);
    }
  });

  btnFit.addEventListener("click", () => fitChart());
  btnReset.addEventListener("click", () => resetChartZoom());

  init();
</script>
</body>
</html>
""".replace("__ORG_JSON__", org_json)


# Temporary replacement disabled: use the original org chart below.
# components.html(browser_html, height=980, scrolling=True)
# st.stop()


# ---------- HTML ----------
html = r"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<style>
  :root{
    --border:#d6eaf6;
    --soft:#eef8fd;
    --muted:#5d7b94;
    --shadow:0 14px 28px rgba(14,58,103,.10);
    --line:#a6d7ec;

    --pd-bg:#e8f6fd;
    --pd-stroke:#0096db;

    --node-bg:#f7fcff;
    --node-stroke:#2a78b0;
  }

  body{
    margin:0; padding:0;
    font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,"Noto Sans","Helvetica Neue",sans-serif;
    background:linear-gradient(180deg, #fbfeff 0%, #f3faff 100%);
  }

  .wrap{
    border:1px solid var(--border);
    border-radius:16px;
    overflow:hidden;
    background:#fff;
    box-shadow: var(--shadow);
  }

  .topbar{
    display:flex;
    justify-content:flex-end;
    align-items:center;
    padding:12px 16px;
    background:linear-gradient(90deg, #eaf7fd 0%, #f8fcff 100%);
    border-bottom:1px solid var(--border);
    font-weight:900;
    font-size:14px;
    color:#0e3a67;
  }

  .righttools{
    display:flex; gap:10px; align-items:center;
    font-weight:850; font-size:13px; color:var(--muted);
  }

  .btn{
    border:1px solid #cfe5f3;
    background:#fff;
    border-radius:10px;
    padding:5px 10px;
    font-weight:900;
    color:#0e3a67;
    cursor:pointer;
    user-select:none;
  }
  .btn:hover{ background:#eef8fd; }

  .content{ padding:12px; }

  .headerArea{
    border:1px solid var(--border);
    border-radius:14px;
    background:#fafdff;
    overflow:hidden;
    margin-bottom:10px;
    position:relative;
    padding-top:18px;
    padding-bottom:14px;
  }

  .org-note{
    position:absolute;
    top:10px;
    left:16px;
    font-size:11px;
    font-weight:700;
    color:#111111;
    line-height:1;
    z-index:2;
  }

  #headerSvg{
    width:100%;
    height:214px;
    display:block;
    background:#fafdff;
  }

  .hdr-edge{
    stroke: var(--line);
    stroke-width: 2.0;
    fill: none;
    stroke-linecap: round;
    stroke-linejoin: round;
  }

  .hdr-box{
    fill: var(--node-bg);
    stroke: var(--node-stroke);
    stroke-width: 1.2;
    rx: 14;
    ry: 14;
    filter: drop-shadow(0 10px 20px rgba(14,58,103,.10));
  }

  .hdr-title{
    font-weight: 950;
    fill:#0e3a67;
    text-anchor: middle;
    dominant-baseline: middle;
  }

  .hdr-sub{
    fill: #5d7b94;
    font-weight: 850;
    text-anchor: middle;
    dominant-baseline: middle;
  }

  .pd-box{
    fill: var(--pd-bg);
    stroke: var(--pd-stroke);
    stroke-width: 1.8;
    rx: 14;
    ry: 14;
    filter: drop-shadow(0 10px 20px rgba(14,58,103,.10));
    cursor: pointer;
  }

  .pd-text{
    font-weight: 950;
    fill:#0e3a67;
    text-anchor: middle;
    dominant-baseline: middle;
    cursor: pointer;
    user-select:none;
  }

  .pd-active{
    stroke-width: 3.6;
  }

  .viewport{
    border:1px solid var(--border);
    border-radius:14px;
    overflow:hidden;
    background:#fafdff;
    position:relative;
    height: 460px;
  }

  .pdHeaderLine{
    position:absolute;
    left:14px; top:12px;
    font-weight:950; font-size:18px;
    line-height:1;
    color:#0e3a67;
  }

  .hint{
    position:absolute;
    right:14px; top:16px;
    font-weight:850; font-size:12px; color:var(--muted);
  }

  #detailSvg{
    width:100%;
    height:100%;
    display:block;
    background:#fafdff;
  }

  .node-rect{
    fill: var(--node-bg);
    stroke: var(--node-stroke);
    stroke-width: 1.2;
    rx: 12;
    ry: 12;
  }

  .node-title{
    font-weight: 950;
    fill:#0e3a67;
    text-anchor: middle;
    dominant-baseline: middle;
  }

  .node-sub{
    fill: #5d7b94;
    font-weight: 800;
    text-anchor: middle;
    dominant-baseline: middle;
  }

  .edge{
    stroke: var(--line);
    stroke-width: 1.6;
    fill: none;
    stroke-linecap: round;
    stroke-linejoin: round;
  }
</style>
</head>

<body>
  <div class="wrap">
    <div class="topbar">
      <div class="righttools"></div>
    </div>

    <div class="content">
      <div class="headerArea">
        <div class="org-note">Excludes coops and contractors</div>
        <svg id="headerSvg" xmlns="http://www.w3.org/2000/svg"></svg>
      </div>

      <div class="viewport" id="viewport">
        <div class="pdHeaderLine" id="pdTitle"></div>
        <div class="hint">Wheel=Zoom, Drag=Pan</div>

        <svg id="detailSvg" xmlns="http://www.w3.org/2000/svg">
          <g id="zoomGroup">
            <g id="edgeLayer"></g>
            <g id="nodeLayer"></g>
          </g>
        </svg>
      </div>
    </div>
  </div>

<script>
  const ORG = __ORG_JSON__;

  const headerSvg = document.getElementById("headerSvg");
  const pdTitleEl = document.getElementById("pdTitle");

  const viewport = document.getElementById("viewport");
  const svg = document.getElementById("detailSvg");
  const zoomGroup = document.getElementById("zoomGroup");
  const edgeLayer = document.getElementById("edgeLayer");
  const nodeLayer = document.getElementById("nodeLayer");

  const btnFit = document.getElementById("btnFit");
  const btnReset = document.getElementById("btnReset");

  const pds = ORG.pds || [];
  let activePD = "";
  let fitConfig = { maxScale: 1.0, padX: 24, padTop: 20, padBottom: 26, mode: "balanced", verticalBias: 0.52 };

  function svgEl(tag){
    return document.createElementNS("http://www.w3.org/2000/svg", tag);
  }

  function clearDetail(){
    edgeLayer.innerHTML = "";
    nodeLayer.innerHTML = "";
  }

  let tx = 0, ty = 0, scale = 1;

  function applyTransform(){
    zoomGroup.setAttribute("transform", "translate(" + tx + "," + ty + ") scale(" + scale + ")");
  }

  function resetToNeutral(){
    tx = 0; ty = 0; scale = 1;
    applyTransform();
  }

  svg.addEventListener("wheel", (e) => {
    e.preventDefault();
    const delta = -e.deltaY;
    const factor = delta > 0 ? 1.08 : 0.92;

    const newScale = Math.min(3.0, Math.max(0.25, scale * factor));

    const pt = svg.createSVGPoint();
    pt.x = e.clientX; pt.y = e.clientY;
    const ctm = svg.getScreenCTM();
    if(!ctm) return;

    const inv = ctm.inverse();
    const p = pt.matrixTransform(inv);

    const x0 = (p.x - tx) / scale;
    const y0 = (p.y - ty) / scale;

    scale = newScale;
    tx = p.x - x0 * scale;
    ty = p.y - y0 * scale;

    applyTransform();
  }, { passive:false });

  let dragging = false;
  let lastX = 0, lastY = 0;

  svg.addEventListener("mousedown", (e) => {
    dragging = true;
    lastX = e.clientX; lastY = e.clientY;
  });

  window.addEventListener("mouseup", () => dragging = false);

  window.addEventListener("mousemove", (e) => {
    if(!dragging) return;
    const dx = e.clientX - lastX;
    const dy = e.clientY - lastY;
    lastX = e.clientX; lastY = e.clientY;
    tx += dx;
    ty += dy;
    applyTransform();
  });

  function drawHeader(){
    headerSvg.innerHTML = "";

    const W = headerSvg.clientWidth || 1200;

    const execW = 330, execH = 64;
    const assocW = 316, assocH = 60;
    const n = pds.length;
    const pdProfile = n <= 8
      ? { w: 148, h: 46, font: 18 }
      : n <= 10
        ? { w: 138, h: 42, font: 17 }
        : { w: 126, h: 38, font: 15 };
    const pdW = pdProfile.w, pdH = pdProfile.h;

    const yExec = 10;
    const yAssoc = 20;
    const yPD = 146;
    const busY = 114;

    const execX = (W - execW) / 2;
    const assocX = execX + execW + 22;

    const execRect = svgEl("rect");
    execRect.setAttribute("class", "hdr-box");
    execRect.setAttribute("x", execX);
    execRect.setAttribute("y", yExec);
    execRect.setAttribute("width", execW);
    execRect.setAttribute("height", execH);
    headerSvg.appendChild(execRect);

    const execName = svgEl("text");
    execName.setAttribute("x", execX + execW/2);
    execName.setAttribute("y", yExec + 25);
    execName.setAttribute("font-size", "20");
    execName.setAttribute("class", "hdr-title");
    execName.textContent = ORG.exec_name || "N/A";
    headerSvg.appendChild(execName);

    const execSub = svgEl("text");
    execSub.setAttribute("x", execX + execW/2);
    execSub.setAttribute("y", yExec + 43);
    execSub.setAttribute("font-size", "14");
    execSub.setAttribute("class", "hdr-sub");
    execSub.textContent = "Executive Director";
    headerSvg.appendChild(execSub);

    const assocRect = svgEl("rect");
    assocRect.setAttribute("class", "hdr-box");
    assocRect.setAttribute("x", assocX);
    assocRect.setAttribute("y", yAssoc);
    assocRect.setAttribute("width", assocW);
    assocRect.setAttribute("height", assocH);
    headerSvg.appendChild(assocRect);

    const assocName = svgEl("text");
    assocName.setAttribute("x", assocX + assocW/2);
    assocName.setAttribute("y", yAssoc + 24);
    assocName.setAttribute("font-size", "18");
    assocName.setAttribute("class", "hdr-title");
    assocName.textContent = ORG.assoc_name || "N/A";
    headerSvg.appendChild(assocName);

    const assocSub = svgEl("text");
    assocSub.setAttribute("x", assocX + assocW/2);
    assocSub.setAttribute("y", yAssoc + 40);
    assocSub.setAttribute("font-size", "12");
    assocSub.setAttribute("class", "hdr-sub");
    assocSub.textContent = "Associate Executive Director";
    headerSvg.appendChild(assocSub);

    const exRightX = execX + execW;
    const exMidY = yExec + execH/2;
    const asLeftX = assocX;
    const asMidY = yAssoc + assocH/2;
    const elbowY = exMidY + 10;

    const execAssoc = svgEl("path");
    execAssoc.setAttribute("class", "hdr-edge");
    execAssoc.setAttribute("d",
      "M " + exRightX + " " + exMidY +
      " L " + (exRightX + 14) + " " + exMidY +
      " L " + (exRightX + 14) + " " + elbowY +
      " L " + (asLeftX - 14) + " " + elbowY +
      " L " + (asLeftX - 14) + " " + asMidY +
      " L " + asLeftX + " " + asMidY
    );
    headerSvg.appendChild(execAssoc);

    if(n === 0) return;

    const margin = 22;
    const step = n > 1
      ? (W - margin*2 - pdW) / (n - 1)
      : 0;
    const startX = margin;

    const execCenterX = execX + execW/2;
    const execBottomY = yExec + execH;

    const down = svgEl("path");
    down.setAttribute("class", "hdr-edge");
    down.setAttribute("d", "M " + execCenterX + " " + execBottomY + " L " + execCenterX + " " + busY);
    headerSvg.appendChild(down);

    const busX1 = startX + pdW/2;
    const busX2 = startX + (n-1)*step + pdW/2;

    const bus = svgEl("path");
    bus.setAttribute("class", "hdr-edge");
    bus.setAttribute("d", "M " + busX1 + " " + busY + " L " + busX2 + " " + busY);
    headerSvg.appendChild(bus);

    for(let i=0;i<n;i++){
      const pd = pds[i];
      const x = startX + i*step;
      const cx = x + pdW/2;

      const drop = svgEl("path");
      drop.setAttribute("class", "hdr-edge");
      drop.setAttribute("d", "M " + cx + " " + busY + " L " + cx + " " + (yPD - 6));
      headerSvg.appendChild(drop);

      const r = svgEl("rect");
      r.setAttribute("x", x);
      r.setAttribute("y", yPD);
      r.setAttribute("width", pdW);
      r.setAttribute("height", pdH);
      r.setAttribute("class", "pd-box" + (pd === activePD ? " pd-active" : ""));
      r.addEventListener("click", () => setActive(pd));
      headerSvg.appendChild(r);

      const t = svgEl("text");
      t.setAttribute("x", x + pdW/2);
      t.setAttribute("y", yPD + pdH/2);
      t.setAttribute("font-size", String(pdProfile.font));
      t.setAttribute("class", "pd-text");
      t.textContent = pd;
      t.addEventListener("click", () => setActive(pd));
      headerSvg.appendChild(t);
    }
  }

  function getTreeStats(pdData){
    const { children } = buildAdj(pdData);
    const roots = (pdData.roots || []).slice();
    const queue = roots.map(r => ({ id: r, depth: 0 }));
    const seen = new Set();
    const breadth = new Map();
    let maxDepth = 0;

    while(queue.length){
      const item = queue.shift();
      if(!item || seen.has(item.id)) continue;
      seen.add(item.id);
      breadth.set(item.depth, (breadth.get(item.depth) || 0) + 1);
      maxDepth = Math.max(maxDepth, item.depth);
      (children.get(item.id) || []).forEach(k => queue.push({ id: k, depth: item.depth + 1 }));
    }

    let maxBreadth = 1;
    breadth.forEach(v => { maxBreadth = Math.max(maxBreadth, v); });

    return {
      nodeCount: (pdData.nodes || []).length,
      maxDepth,
      maxBreadth
    };
  }

  function getLayoutProfile(pdData){
    const stats = getTreeStats(pdData);
    const nodeCount = stats.nodeCount;
    const maxBreadth = stats.maxBreadth;

    let profile = {
      nodeW: 238,
      nodeH: 70,
      gapX: 28,
      gapY: 24,
      padTop: 14,
      padLeft: 24,
      maxScale: 1.70,
      nameMax: 28,
      jobMax: 32,
      fitMode: "horizontal"
    };

    if(nodeCount > 10 || maxBreadth > 5){
      profile = {
        nodeW: 206,
        nodeH: 62,
        gapX: 18,
        gapY: 18,
        padTop: 12,
        padLeft: 20,
        maxScale: 1.14,
        nameMax: 22,
        jobMax: 25,
        fitMode: "balanced"
      };
    }
    if(nodeCount > 18 || maxBreadth > 8){
      profile = {
        nodeW: 170,
        nodeH: 54,
        gapX: 12,
        gapY: 18,
        padTop: 14,
        padLeft: 18,
        maxScale: 0.98,
        nameMax: 18,
        jobMax: 21,
        fitMode: "balanced"
      };
    }
    if(nodeCount > 28 || maxBreadth > 11){
      profile = {
        nodeW: 154,
        nodeH: 48,
        gapX: 8,
        gapY: 14,
        padTop: 12,
        padLeft: 14,
        maxScale: 0.88,
        nameMax: 16,
        jobMax: 18,
        fitMode: "balanced"
      };
    }

    return Object.assign({}, profile, stats);
  }

  function getFontSpec(nodeCount, maxBreadth){
    const pressure = Math.max(nodeCount / 14, maxBreadth / 5);
    if(pressure <= 0.7){
      return { title: 16.8, sub: 13.4 };
    }
    if(pressure <= 1.2){
      return { title: 15.2, sub: 12.2 };
    }
    if(pressure <= 1.8){
      return { title: 13.6, sub: 10.8 };
    }
    return { title: 12.2, sub: 9.8 };
  }

  function fitNameFontSize(text, baseSize){
    const len = (text || "").trim().length;
    if(len <= 16) return baseSize;
    if(len <= 20) return Math.max(baseSize - 0.8, 10);
    if(len <= 24) return Math.max(baseSize - 1.4, 9.5);
    return Math.max(baseSize - 2, 9);
  }

  function fitJobFontSize(text, baseSize){
    const len = (text || "").trim().length;
    if(len <= 22) return baseSize;
    if(len <= 28) return Math.max(baseSize - 0.8, 8.8);
    if(len <= 34) return Math.max(baseSize - 1.3, 8.3);
    return Math.max(baseSize - 1.8, 7.8);
  }

  function trimLabel(text, maxChars){
    const clean = (text || "").trim();
    if(clean.length <= maxChars) return clean;
    return clean.slice(0, Math.max(0, maxChars - 1)).trim() + "…";
  }

  function buildAdj(pdData){
    const nodes = pdData.nodes || [];
    const edges = pdData.edges || [];
    const children = new Map();

    nodes.forEach(n => children.set(n.id, []));
    edges.forEach(e => {
      if(children.has(e.source) && children.has(e.target)){
        children.get(e.source).push(e.target);
      }
    });

    const idToNode = new Map(nodes.map(n => [n.id, n]));
    children.forEach((arr) => {
      arr.sort((a,b) => (idToNode.get(a)?.name || "").localeCompare(idToNode.get(b)?.name || ""));
    });

    return { nodes, edges, children, idToNode };
  }

  function measureSubtree(rootId, children, memo){
    if(memo.has(rootId)) return memo.get(rootId);
    const kids = children.get(rootId) || [];
    if(kids.length === 0){
      memo.set(rootId, 1);
      return 1;
    }
    let sum = 0;
    kids.forEach(k => sum += measureSubtree(k, children, memo));
    memo.set(rootId, Math.max(1, sum));
    return memo.get(rootId);
  }

  function normalizePDGraph(pdData){
    const nodes = (pdData.nodes || []).slice();
    const edges = (pdData.edges || []).slice();
    const roots = (pdData.roots || []).slice();

    if(roots.length > 1){
      const VROOT = "__VROOT__";
      nodes.push({ id: VROOT, name: "", job: "", manager: "" });
      roots.forEach(rid => {
        edges.push({ source: VROOT, target: rid });
      });
      return { nodes, edges, roots: [VROOT], vroot: VROOT };
    }
    return { nodes, edges, roots: roots, vroot: null };
  }

  function layoutTree(pdData){
    const normed = normalizePDGraph(pdData);
    const nodes = normed.nodes;
    const edges = normed.edges;
    const roots = normed.roots;
    const vroot = normed.vroot;
    const profile = getLayoutProfile(pdData);

    const tmp = { nodes, edges, roots };
    const { children } = buildAdj(tmp);

    const memo = new Map();
    roots.forEach(r => measureSubtree(r, children, memo));

    const NODE_W = profile.nodeW, NODE_H = profile.nodeH;
    const GAP_X = profile.gapX, GAP_Y = profile.gapY;

    const PADDING_TOP = profile.padTop;
    const PADDING_LEFT = profile.padLeft;

    let curX = PADDING_LEFT;
    const pos = new Map();

    function placeNode(u, depth, xStart){
      const kids = children.get(u) || [];
      const unitW = memo.get(u) || 1;
      const subtreePixelW = unitW * (NODE_W + GAP_X) - GAP_X;

      const xCenter = xStart + subtreePixelW / 2 - NODE_W / 2;
      const y = PADDING_TOP + depth * (NODE_H + GAP_Y);

      pos.set(u, { x: xCenter, y: y });

      let childX = xStart;
      kids.forEach(k => {
        const ku = memo.get(k) || 1;
        const kPixelW = ku * (NODE_W + GAP_X) - GAP_X;
        placeNode(k, depth + 1, childX);
        childX += kPixelW + GAP_X;
      });
    }

    roots.forEach(r => {
      const ru = memo.get(r) || 1;
      const rPixelW = ru * (NODE_W + GAP_X) - GAP_X;
      placeNode(r, 0, curX);
      curX += rPixelW + (GAP_X * 2);
    });

    return { pos, NODE_W, NODE_H, nodes, edges, vroot, profile };
  }

  function renderPD(pdName){
    pdTitleEl.textContent = pdName || "";
    edgeLayer.innerHTML = "";
    nodeLayer.innerHTML = "";

    const pdDataRaw = ORG.pd_map[pdName];
    if(!pdDataRaw || !pdDataRaw.nodes || pdDataRaw.nodes.length === 0){
      resetToNeutral();
      return;
    }

    const { pos, NODE_W, NODE_H, nodes, edges, vroot, profile } = layoutTree(pdDataRaw);
    const idToNode = new Map(nodes.map(n => [n.id, n]));
    const realNodeCount = nodes.filter(n => !(vroot && n.id === vroot)).length;
    const fontSpec = getFontSpec(realNodeCount, profile.maxBreadth);
    if(profile.maxDepth <= 1 && realNodeCount <= 10){
      viewport.style.height = "300px";
    } else if(profile.maxDepth <= 2 && realNodeCount <= 18){
      viewport.style.height = "380px";
    } else {
      viewport.style.height = "470px";
    }
    fitConfig = {
      maxScale: profile.maxScale,
      padX: realNodeCount <= 10 ? 16 : 20,
      padTop: pdName === "RAC" ? 18 : 26,
      padBottom: pdName === "RAC" ? 34 : 24,
      mode: (profile.maxDepth <= 1 && realNodeCount <= 10) ? "horizontal" : profile.fitMode,
      verticalBias: pdName === "RAC" ? 0.24 : 0.52
    };

    edges.forEach(e => {
      const a = pos.get(e.source);
      const b = pos.get(e.target);
      if(!a || !b) return;

      const x1 = a.x + NODE_W/2;
      const y1 = a.y + NODE_H;
      const x2 = b.x + NODE_W/2;
      const y2 = b.y;
      const midY = (y1 + y2) / 2;

      const path = svgEl("path");
      path.setAttribute("class", "edge");
      path.setAttribute("d",
        "M " + x1 + " " + y1 +
        " L " + x1 + " " + midY +
        " L " + x2 + " " + midY +
        " L " + x2 + " " + y2
      );
      edgeLayer.appendChild(path);
    });

    pos.forEach((p, id) => {
      if(vroot && id === vroot) return;

      const nd = idToNode.get(id) || { name:id, job:"" };

      const g = svgEl("g");
      g.setAttribute("transform", "translate(" + p.x + "," + p.y + ")");

      const rect = svgEl("rect");
      rect.setAttribute("class", "node-rect");
      rect.setAttribute("width", NODE_W);
      rect.setAttribute("height", NODE_H);
      g.appendChild(rect);

      const tooltip = svgEl("title");
      tooltip.textContent = (nd.name || "") + (nd.job ? " | " + nd.job : "");
      g.appendChild(tooltip);

      const nameText = svgEl("text");
      nameText.setAttribute("x", NODE_W/2);
      nameText.setAttribute("y", String(Math.round(NODE_H * 0.40)));
      nameText.setAttribute("font-size", String(fitNameFontSize(nd.name || "", fontSpec.title)));
      nameText.setAttribute("class", "node-title");
      nameText.textContent = trimLabel(nd.name || "", profile.nameMax);
      g.appendChild(nameText);

      const jobText = svgEl("text");
      jobText.setAttribute("x", NODE_W/2);
      jobText.setAttribute("y", String(Math.round(NODE_H * 0.74)));
      jobText.setAttribute("font-size", String(fitJobFontSize(nd.job || "", fontSpec.sub)));
      jobText.setAttribute("class", "node-sub");
      jobText.textContent = trimLabel(nd.job || "", profile.jobMax);
      g.appendChild(jobText);

      nodeLayer.appendChild(g);
    });

    fitToContent();
  }

  function fitToContent(){
    const bbox = zoomGroup.getBBox();
    const w = svg.clientWidth || 1200;
    const h = svg.clientHeight || 700;

    const padX = fitConfig.padX;
    const padTop = fitConfig.padTop;
    const padBottom = fitConfig.padBottom;
    const verticalBias = fitConfig.verticalBias || 0.5;

    const scaleX = (w - padX*2) / Math.max(1, bbox.width);
    const scaleY = (h - padTop - padBottom) / Math.max(1, bbox.height);
    let newScale;
    if(fitConfig.mode === "horizontal"){
      newScale = Math.min(fitConfig.maxScale, Math.max(0.24, scaleX));
      if(bbox.height * newScale > (h - padTop - padBottom)){
        newScale = Math.min(newScale, scaleY);
      }
    } else {
      newScale = Math.min(fitConfig.maxScale, Math.max(0.24, Math.min(scaleX, scaleY)));
    }

    scale = newScale;
    tx = padX + (w - padX*2 - bbox.width * scale) / 2 - bbox.x * scale;
    const innerHeight = h - padTop - padBottom;
    const spareHeight = Math.max(0, innerHeight - bbox.height * scale);
    ty = padTop + spareHeight * verticalBias - bbox.y * scale;

    applyTransform();
  }

  function setActive(pd){
    if(activePD === pd){
      activePD = "";
      pdTitleEl.textContent = "";
      clearDetail();
      resetToNeutral();
      drawHeader();
      return;
    }
    activePD = pd;
    drawHeader();
    renderPD(pd);
  }

  if(btnFit){
    btnFit.addEventListener("click", () => {
      if(activePD) fitToContent();
    });
  }

  if(btnReset){
    btnReset.addEventListener("click", () => {
      tx = 0;
      ty = 0;
      scale = 1;
      applyTransform();
      if(activePD) fitToContent();
    });
  }

  window.addEventListener("resize", () => {
    drawHeader();
    if(activePD) setTimeout(() => fitToContent(), 60);
  });

  function init(){
    pdTitleEl.textContent = "";
    clearDetail();
    resetToNeutral();
    drawHeader();
  }

  init();
</script>
</body>
</html>
""".replace("__ORG_JSON__", org_json)


components.html(html, height=1180, scrolling=False)
