import os
import json
from datetime import datetime

import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Spreadsheet IDs
# Starlink activation status source (new sheet)
SPREADSHEET_ID_STARLINK = "1XdByRZ3zYX5pfqEoufLXb3qnPTIh2rBmnfl4JzWoEbQ"
# Main schedule / outcome data
SPREADSHEET_ID_MAIN = "1zchK5za6LM5aj91s4KDn-CCgNJ5vQFDsCk_ov4XsSn4"

# Column names with line breaks in headers (as in the Sheets)
SCHEDULE_COL = "Schedule of Delivery/\\nInstallation".replace("\\n", "\n")
OUTCOME_COL = "Outcome Status \\n (to be Accomplished by Supplier)".replace("\\n", "\n")
BLOCKER_COL = "Blocker \\n (to be Accomplished by Supplier)".replace("\\n", "\n")


def _build_sheets_service():
    """Build Google Sheets service, using env var JSON if available."""
    json_env = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if json_env:
        # Be forgiving if the value was pasted with surrounding quotes
        json_env = json_env.strip()
        if json_env and json_env[0] in ("'", '"') and json_env[-1] == json_env[0]:
            json_env = json_env[1:-1]
        sa_info = json.loads(json_env)
        creds = Credentials.from_service_account_info(
            sa_info,
            scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
        )
    else:
        # Fallback for local use with JSON file
        service_account_file = "monitoring-dashboard-485505-73f943f6722d.json"
        creds = Credentials.from_service_account_file(
            service_account_file,
            scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
        )
    return build("sheets", "v4", credentials=creds)


_sheets_service = _build_sheets_service()


def _load_df(spreadsheet_id: str, sheet_name: str) -> pd.DataFrame:
    """Generic loader for a given sheet."""
    result = (
        _sheets_service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=f"'{sheet_name}'!A1:ZZ")
        .execute()
    )
    values = result.get("values", [])
    if not values:
        return pd.DataFrame()

    header, *rows = values
    cols_len = len(header)
    normalized_rows = []
    for r in rows:
        if len(r) < cols_len:
            r = r + [""] * (cols_len - len(r))
        elif len(r) > cols_len:
            r = r[:cols_len]
        normalized_rows.append(r)
    return pd.DataFrame(normalized_rows, columns=header)


def load_starlink_df() -> pd.DataFrame:
    """Load BEIS School ID + activation status from the activation sheet."""
    df = _load_df(SPREADSHEET_ID_STARLINK, "Master")
    required = ["BEIS School ID", "Status of Activation", "Approval (Accepted / Decline) "]
    if df.empty or any(col not in df.columns for col in required):
        return pd.DataFrame(columns=required)

    df = df[required].copy()
    # Drop duplicate "Status of Activation" columns, keep the first
    df = df.loc[:, ~df.columns.duplicated()]
    df["BEIS School ID"] = df["BEIS School ID"].astype(str).str.strip()
    df["Status of Activation"] = (
        df["Status of Activation"].fillna("").astype(str).str.strip()
    )
    df["Approval (Accepted / Decline) "] = (
        df["Approval (Accepted / Decline) "].fillna("").astype(str).str.strip()
    )

    # In case of duplicates, keep the last occurrence
    df = df.replace({"": pd.NA}).dropna(subset=["BEIS School ID"]).drop_duplicates(
        subset=["BEIS School ID"], keep="last"
    )
    return df


def load_main_df() -> pd.DataFrame:
    """Load main schedule/outcome data from the second sheet."""
    df = _load_df(SPREADSHEET_ID_MAIN, "Master")

    # In this sheet, the first column header is blank but contains Region values.
    # Normalize that header to "Region" so we can work with it.
    if not df.empty and df.columns[0].strip() == "":
        cols = list(df.columns)
        cols[0] = "Region"
        df.columns = cols

    required = [
        "Region",
        "Province",
        "BEIS School ID",
        SCHEDULE_COL,
        "Start Time",
        "End Time",
        OUTCOME_COL,
        BLOCKER_COL,
    ]
    if df.empty or any(col not in df.columns for col in required):
        return pd.DataFrame(columns=required)

    df = df[required].copy()

    # Clean up core text fields
    for col in ["Region", "Province", "BEIS School ID"]:
        df[col] = df[col].astype(str).str.strip()

    # Clean schedule, outcome, blocker
    df[SCHEDULE_COL] = df[SCHEDULE_COL].fillna("").astype(str).str.strip()
    df[OUTCOME_COL] = df[OUTCOME_COL].fillna("").astype(str).str.strip()
    df[BLOCKER_COL] = df[BLOCKER_COL].fillna("").astype(str).str.strip()

    # Only keep rows where schedule has a value
    df = df[df[SCHEDULE_COL] != ""].copy()

    return df


def get_table_data(selected_region: str | None = None, selected_schedule: str | None = None):
    """Return rows, filter options, and stats for the dashboard."""
    df_main = load_main_df()
    if df_main.empty:
        return [], [], [], {
            "active": False,
            "star_activated": 0,
            "star_not_activated": 0,
            "approval_accepted": 0,
            "approval_pending": 0,
            "approval_decline": 0,
        }

    df_star = load_starlink_df()

    # Join Starlink activation status by BEIS School ID
    if not df_star.empty:
        df_merged = df_main.merge(
            df_star, on="BEIS School ID", how="left", suffixes=("", "_starlink")
        )
        df_merged["Starlink Status"] = df_merged["Status of Activation"]
    else:
        df_merged = df_main.copy()
        df_merged["Starlink Status"] = ""

    # Installation Status derived from Outcome Status (for now, just mirror it)
    df_merged["Installation Status"] = df_merged[OUTCOME_COL]

    # Expose cleaned schedule as a simple field (string)
    df_merged["Schedule"] = df_merged[SCHEDULE_COL].fillna("").astype(str).str.strip()

    # For sorting, parse schedule as a date where possible with several formats
    def _parse_schedule(val: str):
        val = (val or "").strip()
        if not val:
            return pd.NaT
        # Explicit formats we expect to see
        fmts = [
            "%b %d, %Y",      # Feb 05, 2026
            "%b. %d, %Y",     # Feb. 05, 2026
            "%B %d, %Y",      # February 17, 2026
            "%d-%b-%y",       # 05-Feb-26
            "%d-%b-%Y",       # 05-Feb-2026
            "%m/%d/%y",       # 02/04/26 (MM/DD/YY)
            "%m/%d/%Y",       # 02/04/2026
            "%d/%m/%y",       # 04/02/26 (DD/MM/YY)
            "%d/%m/%Y",       # 04/02/2026
        ]
        for fmt in fmts:
            try:
                return datetime.strptime(val, fmt)
            except Exception:
                continue
        # Fallback to pandas parser
        try:
            return pd.to_datetime(val, errors="raise")
        except Exception:
            return pd.NaT

    df_merged["Schedule_sort"] = df_merged["Schedule"].apply(_parse_schedule)
    # Parse times for better ordering within a day
    df_merged["Start_sort"] = pd.to_datetime(
        df_merged["Start Time"], errors="coerce", format="%I:%M %p"
    )
    df_merged["End_sort"] = pd.to_datetime(
        df_merged["End Time"], errors="coerce", format="%I:%M %p"
    )

    # Build a consistent display string for schedule dates, e.g. "Feb. 02, 2026"
    def _format_schedule(row):
        ts = row.get("Schedule_sort")
        s = row["Schedule"]
        if pd.notna(ts):
            try:
                return ts.strftime("%b. %d, %Y")
            except Exception:
                return s
        return s

    df_merged["Schedule_display"] = df_merged.apply(_format_schedule, axis=1)

    # Sort by earliest schedule date, then start/end time, then by region/province/school
    df_sorted = df_merged.sort_values(
        by=["Schedule_sort", "Start_sort", "End_sort", "Region", "Province", "BEIS School ID"],
        kind="stable",
    )

    # All distinct regions for filter options
    region_options = sorted(
        r for r in df_sorted["Region"].astype(str).unique() if r.strip()
    )
    # All distinct schedule display values for filter options, ordered by date
    sched_unique = (
        df_sorted[["Schedule_display", "Schedule_sort"]]
        .drop_duplicates()
        .sort_values(["Schedule_sort", "Schedule_display"])
    )
    schedule_options = [s for s in sched_unique["Schedule_display"].tolist() if s]

    # Optional filters
    if selected_region:
        df_sorted = df_sorted[df_sorted["Region"] == selected_region]
    if selected_schedule:
        df_sorted = df_sorted[df_sorted["Schedule_display"] == selected_schedule]

    # Build stats based on the filtered set (for selected schedule/region)
    if df_sorted.empty:
        stats = {
            "active": False,
            "star_activated": 0,
            "star_not_activated": 0,
            "approval_accepted": 0,
            "approval_pending": 0,
            "approval_decline": 0,
        }
    else:
        star_series = (
            df_sorted["Starlink Status"].fillna("").astype(str).str.strip().str.lower()
        )
        appr_series = (
            df_sorted["Approval (Accepted / Decline) "]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.lower()
        )
        # Starlink: only "activated" vs everything else
        star_activated = (star_series == "activated").sum()
        star_not_activated = len(df_sorted) - star_activated

        # Approval:
        # treat any value containing "accept" as accepted,
        # any value containing "declin" (decline/declined) as decline,
        # the rest (including blank/pending/others) as pending/blank.
        accepted_mask = appr_series.str.contains("accept", na=False)
        decline_mask = appr_series.str.contains("declin", na=False)
        approval_accepted = accepted_mask.sum()
        approval_decline = decline_mask.sum()
        approval_pending = len(df_sorted) - approval_accepted - approval_decline
        stats = {
            "active": True,  # show stats for overall or filtered
            "star_activated": int(star_activated),
            "star_not_activated": int(star_not_activated),
            "approval_accepted": int(approval_accepted),
            "approval_pending": int(approval_pending),
            "approval_decline": int(approval_decline),
        }

    rows = []
    for _, row in df_sorted.iterrows():
        # Normalize Starlink and approval text
        star = row["Starlink Status"]
        if pd.isna(star):
            star = ""
        approval_raw = row.get("Approval (Accepted / Decline) ", "")
        if pd.isna(approval_raw):
            approval_raw = ""

        # Normalize schedule display to a consistent text format, e.g. "Feb. 02, 2026"
        schedule_display = row["Schedule"]
        ts = row.get("Schedule_sort")
        if pd.notna(ts):
            try:
                schedule_display = ts.strftime("%b. %d, %Y")  # e.g. "Feb. 02, 2026"
            except Exception:
                # Fallback to original string if formatting fails
                schedule_display = row["Schedule"]
        rows.append(
            {
                "Region": row["Region"],
                "Province": row["Province"],
                "BEIS School ID": row["BEIS School ID"],
                "Schedule": row["Schedule_display"],
                "Start Time": row["Start Time"],
                "End Time": row["End Time"],
                "Installation Status": row["Installation Status"],
                "Starlink Status": star,
                "Approval": approval_raw,
                "Blocker": row[BLOCKER_COL],
            }
        )

    return rows, region_options, schedule_options, stats


TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>LEOxSOLAR Schedule Monitoring</title>
    <meta http-equiv="refresh" content="60">
    <style>
        html, body {
            height: 100%;
            margin: 0;
        }
        body {
            font-family: Arial, sans-serif;
            font-size: 12px;
            background: #f1f5f9;
            overflow: hidden; /* prevent whole-page scrolling */
        }
        .page {
            max-width: 1200px;
            height: 100%;
            margin: 0 auto;
            padding: 10px;
            box-sizing: border-box;
            display: flex;
            flex-direction: column;
        }
        h1 {
            margin: 4px 0 6px 0;
            font-size: 20px;
            color: #1f2933;
        }
        .meta-line {
            font-size: 11px;
            color: #6b7280;
            margin-bottom: 6px;
        }
        .filter-bar {
            display: flex;
            align-items: center;
            gap: 8px;
            margin-bottom: 6px;
            font-size: 12px;
        }
        .filter-bar label {
            color: #4b5563;
            font-weight: 500;
        }
        .filter-bar select {
            font-size: 12px;
            padding: 2px 8px;
            border-radius: 9999px;
            border: 1px solid #d0d7e2;
            background-color: #ffffff;
        }
        .card {
            background: #ffffff;
            border-radius: 6px;
            box-shadow: 0 1px 3px rgba(15, 23, 42, 0.15);
            padding: 8px;
            box-sizing: border-box;
            flex: 1;                 /* take remaining space below header/meta */
            display: flex;
            flex-direction: row;
            gap: 8px;
            overflow: hidden;
        }
        .table-wrapper {
            flex: 3;
            overflow-y: auto;        /* scroll only table area */
        }
        table {
            border-collapse: collapse;
            width: 100%;
            table-layout: fixed;
        }
        th, td {
            border: 1px solid #d0d7e2;
            padding: 2px 4px;
            text-align: center;
            font-size: 11px;
        }
        th {
            background-color: #e5edf7;
            font-weight: 600;
            color: #25313d;
            white-space: normal;           /* allow header text to wrap */
            word-wrap: break-word;         /* break long tokens like 'Delivery/Installation' */
            word-break: break-word;
        }
        thead th {
            position: sticky;
            top: 0;
            z-index: 2;
        }
        tbody tr:nth-child(even) td {
            background-color: #f8fafc;
        }
        tbody tr:hover td {
            background-color: #e5f0ff;
        }
        .row-warning td {
            background-color: #fef9c3;
        }
        .row-critical td {
            background-color: #fee2e2;
        }
        .region-cell {
            text-align: left;
            font-weight: 600;
            padding-left: 6px;
        }
        .school-cell {
            text-align: left;
            padding-left: 6px;
            white-space: normal;
            word-wrap: break-word;
        }
        .status-cell {
            text-align: left;
            padding-left: 6px;
            white-space: normal;
            word-wrap: break-word;
        }
        .stats-card {
            flex: 1;
            max-width: 320px;
            padding: 6px;
            background: #ffffff;
            border-radius: 6px;
            box-shadow: 0 1px 3px rgba(15, 23, 42, 0.15);
            font-size: 12px;
            overflow: hidden;
            display: flex;
            flex-direction: column;
        }
        .stats-title {
            font-weight: 600;
            margin-bottom: 4px;
            color: #25313d;
        }
        .stats-main {
            flex: 1;
            display: flex;
            flex-direction: column;
            gap: 6px;
            overflow: hidden;
        }
        .stats-grid {
            flex: 0 0 auto;
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
        }
        .stats-item {
            min-width: 110px;
            padding: 4px 6px;
            border-radius: 4px;
            background: #f8fafc;
        }
        .stats-label {
            color: #6b7280;
            font-size: 11px;
        }
        .stats-value {
            font-weight: 700;
            font-size: 14px;
            color: #111827;
        }
        .stats-charts {
            flex: 1;
            display: flex;
            flex-direction: column;  /* stack charts vertically */
            gap: 6px;
            align-items: center;
            justify-content: center;
        }
        .stats-chart {
            flex: 1;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
        }
        .status-pill {
            display: inline-block;
            padding: 2px 6px;
            border-radius: 9999px;
            font-weight: 600;
            font-size: 11px;
        }
        .status-ok {
            background-color: #16a34a;
            color: #ffffff;
        }
        .status-bad {
            background-color: #dc2626;
            color: #ffffff;
        }
        .status-warn {
            background-color: #facc15;
            color: #1f2933;
        }
    </style>
</head>
<body>
    <div class="page">
        <h1>LEOxSOLAR Schedule Monitoring</h1>
        <div class="meta-line">
            Auto-refresh: 60s | Last update: {{ last_updated }}
        </div>
        <div class="meta-line">
            Showing {{ rows|length }} records
            • Region: {{ selected_region or 'All' }}
            • Schedule: {{ selected_schedule or 'All' }}
        </div>

        <form method="get" class="filter-bar">
            <label for="region-select">Region:</label>
            <select id="region-select" name="region" onchange="this.form.submit()">
                <option value="">All Regions</option>
                {% for r in region_options %}
                <option value="{{ r }}" {% if selected_region == r %}selected{% endif %}>{{ r }}</option>
                {% endfor %}
            </select>
            <label for="schedule-select">Schedule:</label>
            <select id="schedule-select" name="schedule" onchange="this.form.submit()">
                <option value="">All Dates</option>
                {% for d in schedule_options %}
                <option value="{{ d }}" {% if selected_schedule == d %}selected{% endif %}>{{ d }}</option>
                {% endfor %}
            </select>
        </form>

        <div class="card">
            <div class="table-wrapper">
                <table>
                    <thead>
                        <tr>
                            <th>Region</th>
                            <th>Province</th>
                            <th>BEIS ID</th>
                            <th title="Schedule of Delivery/Installation">Schedule</th>
                            <th title="Start Time">Start</th>
                            <th title="End Time">End</th>
                            <th title="Outcome Status (to be Accomplished by Supplier)">Installation</th>
                            <th>Starlink</th>
                            <th title="Approval (Accepted / Decline)">Approval</th>
                            <th title="Blocker (to be Accomplished by Supplier)">Blocker</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for row in rows %}
                        {% set star = (row["Starlink Status"] or "") | lower %}
                        {% set appr = (row["Approval"] or "") | lower %}
                        {% set row_class = '' %}
                        {% if 'declin' in appr %}
                            {% set row_class = 'row-critical' %}
                        {% elif star != 'activated' %}
                            {% set row_class = 'row-warning' %}
                        {% endif %}
                        <tr class="{{ row_class }}">
                            <td class="region-cell">{{ row["Region"] }}</td>
                            <td>{{ row["Province"] }}</td>
                            <td class="school-cell">{{ row["BEIS School ID"] }}</td>
                            <td>{{ row["Schedule"] }}</td>
                            <td>{{ row["Start Time"] }}</td>
                            <td>{{ row["End Time"] }}</td>
                            <td>{{ row["Installation Status"] }}</td>
                            <td>
                                <span class="status-pill {% if star == 'activated' %}status-ok{% else %}status-bad{% endif %}">
                                    {{ row["Starlink Status"] or 'Not Activated' }}
                                </span>
                            </td>
                            <td>
                                <span class="status-pill
                                    {% if appr == 'accepted' %}
                                        status-ok
                                    {% elif appr == 'pending' or appr == '' %}
                                        status-warn
                                    {% else %}
                                        status-bad
                                    {% endif %}">
                                    {{ row["Approval"] or 'Pending' }}
                                </span>
                            </td>
                            <td class="status-cell">{{ row["Blocker"] }}</td>
                        </tr>
                        {% endfor %}
                        {% if rows|length == 0 %}
                        <tr>
                            <td colspan="9">No data available (check sheet names/columns or schedule values).</td>
                        </tr>
                        {% endif %}
                    </tbody>
                </table>
            </div>
            {% if stats.active %}
            <div class="stats-card">
                <div class="stats-title">
                    Summary for {{ selected_schedule or 'All Schedules' }}
                </div>
                <div class="stats-main">
                    <div class="stats-grid">
                        <div class="stats-item">
                            <div class="stats-label">✔ Starlink Activated</div>
                            <div class="stats-value">{{ stats.star_activated }}</div>
                        </div>
                        <div class="stats-item">
                            <div class="stats-label">⚠ Starlink Not Activated</div>
                            <div class="stats-value">{{ stats.star_not_activated }}</div>
                        </div>
                        <div class="stats-item">
                            <div class="stats-label">✔ Approval Accepted</div>
                            <div class="stats-value">{{ stats.approval_accepted }}</div>
                        </div>
                        <div class="stats-item">
                            <div class="stats-label">⚠ Approval Pending / Blank</div>
                            <div class="stats-value">{{ stats.approval_pending }}</div>
                        </div>
                        <div class="stats-item">
                            <div class="stats-label">✖ Approval Decline / Other</div>
                            <div class="stats-value">{{ stats.approval_decline }}</div>
                        </div>
                    </div>
                    <div class="stats-charts">
                        <div class="stats-chart">
                            <div class="stats-label">Starlink Status</div>
                            <canvas id="starChart" style="width: 100%; max-width: 210px; height: 100px;"></canvas>
                        </div>
                        <div class="stats-chart">
                            <div class="stats-label">Approval Status</div>
                            <canvas id="approvalChart" style="width: 100%; max-width: 210px; height: 100px;"></canvas>
                        </div>
                    </div>
                </div>
            </div>
            {% endif %}
        </div>
    </div>
</body>
</html>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<script>
(function () {
    const active = {{ 'true' if stats.active else 'false' }};
    if (!active) return;

    // Starlink pie (Activated vs Not Activated)
    const starEl = document.getElementById('starChart');
    if (starEl) {
        const starCtx = starEl.getContext('2d');
        const starData = [{{ stats.star_activated }}, {{ stats.star_not_activated }}];
        new Chart(starCtx, {
            type: 'pie',
            data: {
                labels: ['Activated', 'Not Activated'],
                datasets: [{
                    data: starData,
                    backgroundColor: [
                        'rgba(22, 163, 74, 0.85)',  // green
                        'rgba(220, 38, 38, 0.85)',  // red
                    ],
                    borderWidth: 0,
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: { boxWidth: 10, font: { size: 9 } }
                    }
                }
            }
        });
    }

    // Approval pie (Accepted / Pending / Decline)
    const apprEl = document.getElementById('approvalChart');
    if (apprEl) {
        const apprCtx = apprEl.getContext('2d');
        const apprData = [
            {{ stats.approval_accepted }},
            {{ stats.approval_pending }},
            {{ stats.approval_decline }},
        ];
        new Chart(apprCtx, {
            type: 'pie',
            data: {
                labels: ['Accepted', 'Pending/Blank', 'Decline/Other'],
                datasets: [{
                    data: apprData,
                    backgroundColor: [
                        'rgba(34, 197, 94, 0.85)',   // green
                        'rgba(250, 204, 21, 0.9)',   // yellow
                        'rgba(220, 38, 38, 0.85)',   // red
                    ],
                    borderWidth: 0,
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: { boxWidth: 10, font: { size: 9 } }
                    }
                }
            }
        });
    }
})();
</script>
"""
