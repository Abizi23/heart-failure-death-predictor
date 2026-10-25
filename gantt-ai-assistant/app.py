import json
import time
import warnings
from datetime import date, datetime
from io import BytesIO

import pandas as pd
import plotly.express as px
import requests
import streamlit as st
import urllib3
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# --------------------------------------------------
# CONFIGURATION
# --------------------------------------------------

API_URL = "https://localhost:8000/events"

# Outlook add-in design tokens (matches taskpane.css exactly)
NAVY       = "#173b68"
ACCENT     = "#3282ce"
BG         = "#f4f7fb"
SIDEBAR_BG = "#edf3fa"
CARD_BG    = "#ffffff"
BORDER     = "#dce4ef"
TEXT       = "#172b4d"
MUTED      = "#52647b"
GREEN_OK   = "#15803d"
RED_ERR    = "#b91c1c"
AMBER      = "#b45309"

st.set_page_config(
    page_title="Gantt AI Assistant",
    page_icon="\U0001f4ca",
    layout="wide",
)

warnings.filterwarnings("ignore", message="Unverified HTTPS request")
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --------------------------------------------------
# DASHBOARD STYLING  (mirrors Outlook taskpane.css)
# --------------------------------------------------

st.markdown(
    f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Segoe+UI:wght@400;600;700;800&display=swap');

    html, body, [class*="css"] {{
        font-family: "Segoe UI", Arial, sans-serif !important;
        color: {TEXT};
        background-color: {BG};
    }}

    .stApp {{
        background-color: {BG};
    }}

    /* ── Sidebar ── */
    [data-testid="stSidebar"] {{
        background-color: {SIDEBAR_BG};
        border-right: 1px solid {BORDER};
    }}
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {{
        color: {NAVY} !important;
        font-weight: 700 !important;
    }}

    /* ── Page header banner ── */
    .gantt-header {{
        display: flex;
        align-items: center;
        gap: 14px;
        padding: 20px 24px;
        background: {NAVY};
        border-radius: 12px;
        margin-bottom: 24px;
        box-shadow: 0 4px 14px rgba(23,59,104,0.18);
    }}
    .gantt-logo {{
        display: grid;
        place-items: center;
        width: 48px;
        height: 48px;
        border-radius: 12px;
        background: #38bdf8;
        color: #082f49;
        font-size: 26px;
        font-weight: 800;
        flex-shrink: 0;
    }}
    .gantt-title {{
        color: white;
        font-size: 22px;
        font-weight: 800;
        margin: 0;
        line-height: 1.3;
    }}
    .gantt-subtitle {{
        color: #dceafa;
        font-size: 13px;
        margin: 2px 0 0;
    }}

    /* ── Metric cards ── */
    [data-testid="stMetric"] {{
        background: {CARD_BG};
        border: 1px solid {BORDER};
        border-radius: 12px;
        padding: 18px 20px;
        box-shadow: 0 2px 7px rgba(23,59,104,0.07);
    }}
    [data-testid="stMetricLabel"] {{
        color: {MUTED} !important;
        font-weight: 600 !important;
        font-size: 13px !important;
    }}
    [data-testid="stMetricValue"] {{
        color: {NAVY} !important;
        font-weight: 800 !important;
    }}

    /* ── Headings ── */
    h1, h2, h3, h4 {{
        color: {NAVY} !important;
        font-weight: 800 !important;
    }}

    /* ── Buttons ── */
    .stButton > button {{
        background-color: {NAVY};
        color: white !important;
        font-weight: 700;
        border-radius: 6px;
        border: 1px solid {NAVY};
        padding: 7px 16px;
        transition: filter 0.15s;
    }}
    .stButton > button:hover {{
        filter: brightness(1.12);
        border-color: {ACCENT};
    }}

    .stDownloadButton > button {{
        background-color: {GREEN_OK};
        color: white !important;
        font-weight: 700;
        border-radius: 6px;
        border: 1px solid {GREEN_OK};
    }}
    .stDownloadButton > button:hover {{
        filter: brightness(1.1);
    }}

    /* ── Dataframe / table ── */
    [data-testid="stDataFrame"] {{
        border: 1px solid {BORDER};
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 2px 7px rgba(23,59,104,0.05);
    }}

    /* ── Divider ── */
    hr {{
        border-color: {BORDER};
    }}

    /* ── Multiselect pills ── */
    [data-baseweb="tag"] {{
        background-color: {NAVY} !important;
    }}

    /* ── Status/caption text ── */
    .stCaption, small {{
        color: {MUTED} !important;
        font-size: 12px;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

# --------------------------------------------------
# FETCH EVENTS FROM FASTAPI
# --------------------------------------------------

def fetch_events():
    """Fetch events directly from the backend without caching so the
    dashboard always reflects the latest data."""
    try:
        response = requests.get(
            API_URL,
            timeout=10,
            verify=False,  # Local development certificate only
        )
        response.raise_for_status()
        data = response.json()

        if isinstance(data, dict):
            data = data.get("events", data.get("data", []))

        if not isinstance(data, list):
            return [], "The backend returned an unexpected data format."

        return data, None

    except requests.exceptions.RequestException as exc:
        return [], str(exc)


# --------------------------------------------------
# DATA PREPARATION
# --------------------------------------------------

def prepare_dataframe(events):
    df = pd.DataFrame(events)

    if df.empty:
        return df

    for column in [
        "event_type",
        "task",
        "project",
        "status",
        "start_date",
        "end_date",
        "start_time",
        "end_time",
        "resources",
        "location",
        "email_subject",
    ]:
        if column not in df.columns:
            df[column] = ""

    df["task"] = df["task"].fillna("").astype(str)
    df["project"] = df["project"].fillna("Unassigned").astype(str)
    df["status"] = df["status"].fillna("Proposed").astype(str)

    # Convert dates consistently to Python dates.
    for column in ["start_date", "end_date"]:
        df[column] = pd.to_datetime(
            df[column], errors="coerce"
        ).apply(lambda value: value.date() if pd.notna(value) else None)

    return df


# --------------------------------------------------
# EXCEL EXPORT
# --------------------------------------------------

def create_excel_file(df):
    output = BytesIO()
    workbook = Workbook()

    tasks_sheet = workbook.active
    tasks_sheet.title = "Tasks"

    export_columns = [
        column
        for column in [
            "id",
            "task",
            "project",
            "event_type",
            "start_date",
            "start_time",
            "end_date",
            "end_time",
            "hours",
            "resources",
            "location",
            "status",
            "email_subject",
        ]
        if column in df.columns
    ]

    tasks_sheet.append(export_columns)

    for _, row in df[export_columns].iterrows():
        tasks_sheet.append(
            [
                value.isoformat()
                if isinstance(value, (date, datetime))
                else str(value)
                if isinstance(value, (dict, list))
                else value
                for value in row
            ]
        )

    # Gantt chart worksheet.
    gantt_sheet = workbook.create_sheet("Gantt Chart")
    gantt_sheet.append(
        ["Task", "Project", "Start Date", "End Date", "Status"]
    )

    for _, row in df.iterrows():
        gantt_sheet.append(
            [
                row.get("task", ""),
                row.get("project", ""),
                str(row.get("start_date", "") or ""),
                str(row.get("end_date", "") or ""),
                row.get("status", ""),
            ]
        )

    # Summary worksheet.
    summary_sheet = workbook.create_sheet("Summary")
    summary_sheet.append(["Metric", "Value"])
    summary_sheet.append(["Total events", len(df)])
    summary_sheet.append(
        ["Accepted", int((df["status"] == "Accepted").sum())]
    )
    summary_sheet.append(
        ["Proposed", int((df["status"] == "Proposed").sum())]
    )
    summary_sheet.append(
        ["Rejected", int((df["status"] == "Rejected").sum())]
    )

    # Apply consistent formatting.
    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions

        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill(
                fill_type="solid",
                fgColor="217346",
            )
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

        for column_cells in sheet.columns:
            max_length = max(
                len(str(cell.value or "")) for cell in column_cells
            )
            column_letter = get_column_letter(column_cells[0].column)
            sheet.column_dimensions[column_letter].width = min(
                max(max_length + 3, 12), 40
            )

    workbook.save(output)
    output.seek(0)
    return output.getvalue()


# --------------------------------------------------
# HEADER
# --------------------------------------------------

# --------------------------------------------------
# HEADER BANNER  (matches Outlook add-in app-header)
# --------------------------------------------------

st.markdown(
    """
    <div class="gantt-header">
      <div class="gantt-logo">G</div>
      <div>
        <p class="gantt-title">GANTT AI ASSISTANT</p>
        <p class="gantt-subtitle">Project schedule dashboard &mdash; synced from your Outlook workflow</p>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# --------------------------------------------------
# AUTO-REFRESH (every 10 seconds)
# --------------------------------------------------

AUTO_REFRESH_INTERVAL = 10  # seconds

if "last_refresh" not in st.session_state:
    st.session_state["last_refresh"] = time.time()

time_since_refresh = time.time() - st.session_state["last_refresh"]
if time_since_refresh >= AUTO_REFRESH_INTERVAL:
    st.session_state["last_refresh"] = time.time()
    st.rerun()

col_refresh, col_status = st.columns([1, 3])

with col_refresh:
    if st.button("🔄 Refresh now", use_container_width=True):
        st.session_state["last_refresh"] = time.time()
        st.rerun()

with col_status:
    last_updated = datetime.fromtimestamp(st.session_state["last_refresh"]).strftime("%H:%M:%S")
    seconds_remaining = max(0, int(AUTO_REFRESH_INTERVAL - time_since_refresh))
    st.caption(
        f"🟢 Live — last updated at **{last_updated}** · "
        f"next auto-refresh in **{seconds_remaining}s**"
    )

events, error = fetch_events()

if error:
    st.error(
        "Could not connect to the backend. Make sure the FastAPI "
        "server is running on https://localhost:8000."
    )
    st.caption(f"Technical detail: {error}")
    st.stop()

df = prepare_dataframe(events)

if df.empty:
    st.info(
        "No events have been saved yet. Analyze an email in Outlook "
        "and approve or save its proposed events first."
    )
    st.stop()

# --------------------------------------------------
# FILTERS
# --------------------------------------------------

st.sidebar.title("Dashboard Filters")

projects = sorted(
    project
    for project in df["project"].unique()
    if project and project != "nan"
)

statuses = sorted(
    status
    for status in df["status"].unique()
    if status and status != "nan"
)

selected_projects = st.sidebar.multiselect(
    "Filter by project",
    options=projects,
    default=projects,
)

selected_statuses = st.sidebar.multiselect(
    "Filter by status",
    options=statuses,
    default=statuses,
)

filtered_df = df[
    df["project"].isin(selected_projects)
    & df["status"].isin(selected_statuses)
].copy()

# --------------------------------------------------
# SUMMARY METRICS
# --------------------------------------------------

st.subheader("Project Overview")

metric1, metric2, metric3, metric4 = st.columns(4)

metric1.metric("Total Events", len(filtered_df))
metric2.metric(
    "Accepted",
    int((filtered_df["status"] == "Accepted").sum()),
)
metric3.metric(
    "Proposed",
    int((filtered_df["status"] == "Proposed").sum()),
)
metric4.metric(
    "Rejected",
    int((filtered_df["status"] == "Rejected").sum()),
)

st.divider()

# --------------------------------------------------
# INTERACTIVE GANTT CHART
# --------------------------------------------------

st.subheader("Project Gantt Chart")

if filtered_df.empty:
    st.info("No events match your selected filters.")
else:
    gantt_df = filtered_df.dropna(
        subset=["start_date", "end_date"]
    ).copy()

    gantt_df = gantt_df[
        gantt_df["end_date"] >= gantt_df["start_date"]
    ].copy()

    if gantt_df.empty:
        st.warning(
            "There are no events with valid start and end dates "
            "to display on the Gantt chart. Check the dates in Outlook."
        )
    else:
        # Plotly works best with pandas datetime values.
        gantt_df["Start"] = pd.to_datetime(gantt_df["start_date"])
        gantt_df["Finish"] = pd.to_datetime(gantt_df["end_date"])
        gantt_df["Task Label"] = gantt_df["task"].replace(
            "", "Untitled task"
        )

        status_colors = {
            "Accepted": GREEN_OK,   # #15803d
            "Proposed": "#b45309",  # amber
            "Rejected": RED_ERR,    # #b91c1c
        }

        fig = px.timeline(
            gantt_df,
            x_start="Start",
            x_end="Finish",
            y="Task Label",
            color="status",
            color_discrete_map=status_colors,
            hover_data={
                "project": True,
                "status": True,
                "Start": True,
                "Finish": True,
                "Task Label": False,
            },
            labels={
                "Task Label": "Task",
                "project": "Project",
                "status": "Status",
            },
        )

        # Style chart to match the Outlook add-in design language
        fig.update_layout(
            title=dict(
                text="Project Schedule",
                font=dict(
                    color=NAVY,
                    size=20,
                    family="Segoe UI, Arial, sans-serif",
                ),
            ),
            font=dict(
                family="Segoe UI, Arial, sans-serif",
                size=13,
                color=TEXT,
            ),
            paper_bgcolor=BG,
            plot_bgcolor=CARD_BG,
            height=max(450, 45 * len(gantt_df)),
            margin=dict(l=30, r=30, t=70, b=50),
            legend=dict(
                title=dict(font=dict(color=TEXT)),
                font=dict(color=TEXT),
                bgcolor=CARD_BG,
                bordercolor=BORDER,
                borderwidth=1,
            ),
            xaxis=dict(
                title=dict(
                    text="Timeline",
                    font=dict(color=MUTED, size=13),
                ),
                tickfont=dict(color=MUTED, size=12),
                showgrid=True,
                gridcolor=BORDER,
                linecolor=BORDER,
                tickformat="%d %b %Y",
            ),
            yaxis=dict(
                title=dict(
                    text="Task",
                    font=dict(color=MUTED, size=13),
                ),
                tickfont=dict(
                    color=TEXT,
                    size=12,
                    family="Segoe UI, Arial, sans-serif",
                ),
                automargin=True,
                autorange="reversed",
                showgrid=False,
            ),
            hoverlabel=dict(
                bgcolor=CARD_BG,
                font=dict(
                    color=TEXT,
                    size=13,
                    family="Segoe UI, Arial, sans-serif",
                ),
                bordercolor=BORDER,
            ),
        )

        fig.update_traces(
            textfont=dict(color=TEXT, size=12),
            marker_line_color=CARD_BG,
            marker_line_width=1,
        )


        st.plotly_chart(
            fig,
            use_container_width=True,
            theme=None,
        )

st.divider()

# --------------------------------------------------
# EVENTS TABLE
# --------------------------------------------------

st.subheader("Project Events")

if filtered_df.empty:
    st.info("No events to display.")
else:
    display_columns = [
        column
        for column in [
            "id",
            "task",
            "project",
            "event_type",
            "start_date",
            "start_time",
            "end_date",
            "end_time",
            "hours",
            "resources",
            "location",
            "status",
            "email_subject",
        ]
        if column in filtered_df.columns
    ]

    display_df = filtered_df[display_columns].copy()

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
    )

# --------------------------------------------------
# EXCEL DOWNLOAD
# --------------------------------------------------

st.divider()
st.subheader("Export Project Data")

st.write(
    "Download the currently filtered events as an Excel workbook "
    "with Tasks, Gantt Chart, and Summary sheets."
)

excel_data = create_excel_file(filtered_df)

st.download_button(
    label="📥 Download Excel Workbook",
    data=excel_data,
    file_name=f"project_gantt_{date.today().isoformat()}.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
    use_container_width=True,
)

st.caption(
    "The dashboard reads saved events from your backend. "
    "Downloading Excel creates a snapshot; it does not automatically "
    "synchronize edits back to the database."
)
