from flask import Flask, render_template_string
import pandas as pd
from datetime import datetime

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

app = Flask(__name__)

# Service account JSON in this folder
SERVICE_ACCOUNT_FILE = "monitoring-dashboard-485505-73f943f6722d.json"

# Live Google Sheet ID (from the sheet URL)
SPREADSHEET_ID = "1u6CjGchWZ7ZWzJefGS0HDOX0GU44ODCwi_HajiH6q2E"


def _build_sheets_service():
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
    )
    return build("sheets", "v4", credentials=creds)


_sheets_service = _build_sheets_service()


def load_leo_solar_df():
    """Load the LEO SOLAR sheet from Google Sheets into a DataFrame."""
    result = (
        _sheets_service.spreadsheets()
        .values()
        # Quote sheet name because it has a space
        .get(spreadsheetId=SPREADSHEET_ID, range="'LEO SOLAR'!A1:ZZ")
        .execute()
    )
    values = result.get("values", [])
    if not values:
        return pd.DataFrame()

    header, *rows = values
    # Pad rows so all have same length as header
    rows = [r + [""] * (len(header) - len(r)) for r in rows]
    return pd.DataFrame(rows, columns=header)


def get_pivots():
    # Read only the "LEO SOLAR" sheet and focus on the needed columns
    sheet = load_leo_solar_df()
    # Note: "Final Status " in the file has a trailing space
    required_cols = ["Region", "Final Status ", "Starlink Status", "Starlink Installation Date"]
    missing = [c for c in required_cols if c not in sheet.columns]
    if missing or sheet.empty:
        # If any required column is missing, return empties so the UI doesn't break
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    df = sheet[required_cols].copy()

    # Normalize
    df["Region"] = df["Region"].astype(str)
    # Normalize column names we will use later
    df = df.rename(columns={"Final Status ": "Final Status"})

    # Clean up status values and map blanks to 'Blank'
    df["Final Status"] = (
        df["Final Status"]
        .fillna("")
        .astype(str)
        .str.strip()
        .replace("", "Blank")
    )
    df["Starlink Status"] = (
        df["Starlink Status"]
        .fillna("")
        .astype(str)
        .str.strip()
        .replace("", "Blank")
    )

    # Count per Region + Final Status
    final_status_cols = ["Ready to Deploy", "For Removal", "For Replacement", "Not Ready", "Blank"]
    final_pivot = df.pivot_table(
        index="Region",
        columns="Final Status",
        aggfunc="size",
        fill_value=0,
    )
    for col in final_status_cols:
        if col not in final_pivot.columns:
            final_pivot[col] = 0
    final_pivot = final_pivot[final_status_cols]

    # Count per Region + Starlink Status
    star_status_cols = ["For Delivery", "For Installation", "Installed", "Blank"]
    star_pivot = df.pivot_table(
        index="Region",
        columns="Starlink Status",
        aggfunc="size",
        fill_value=0,
    )
    for col in star_status_cols:
        if col not in star_pivot.columns:
            star_pivot[col] = 0
    star_pivot = star_pivot[star_status_cols]

    # Ensure both pivots share the same Region index
    all_regions = sorted(set(final_pivot.index) | set(star_pivot.index))
    final_pivot = final_pivot.reindex(all_regions, fill_value=0)
    star_pivot = star_pivot.reindex(all_regions, fill_value=0)

    # Installation summary: rows with Starlink Status == "For Installation"
    install_df = df[df["Starlink Status"] == "For Installation"].copy()
    if not install_df.empty:
        # Keep whatever date format is entered, just group by the cleaned string
        install_df["Starlink Installation Date"] = (
            install_df["Starlink Installation Date"]
            .fillna("")
            .astype(str)
            .str.strip()
        )
        install_df = install_df[install_df["Starlink Installation Date"] != ""]
        install_summary = (
            install_df.groupby(["Region", "Starlink Installation Date"])
            .size()
            .reset_index(name="Count")
            .sort_values(["Region", "Starlink Installation Date"])
        )
    else:
        install_summary = pd.DataFrame(columns=["Region", "Starlink Installation Date", "Count"])

    return final_pivot, star_pivot, install_summary


TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>LEO / Starlink Monitoring</title>
    <!-- Auto-refresh every 60 seconds -->
    <meta http-equiv="refresh" content="60">
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 10px;
            font-size: 12px;
            background: #f1f5f9;
        }
        .page {
            max-width: 1200px;
            margin: 0 auto;
        }
        h1 {
            margin: 4px 0 6px 0;
            font-size: 20px;
            color: #1f2933;
        }
        h2 {
            margin: 10px 0 4px 0;
            font-size: 14px;
            color: #364152;
        }
        h3 {
            margin: 6px 0 4px 0;
            font-size: 13px;
            color: #52606d;
        }
        .layout {
            display: flex;
            align-items: flex-start;
            gap: 16px;
            margin-top: 8px;
        }
        .summary-bar {
            display: flex;
            gap: 8px;
            margin-top: 4px;
            margin-bottom: 4px;
        }
        .summary-card {
            flex: 0 0 auto;
            min-width: 130px;
            padding: 6px 8px;
            border-radius: 6px;
            background: #ffffff;
            box-shadow: 0 1px 2px rgba(15, 23, 42, 0.12);
            font-size: 11px;
            border-left: 3px solid transparent;
        }
        .summary-card.total { border-left-color: #0ea5e9; }
        .summary-card.install { border-left-color: #6366f1; }
        .summary-card.installed { border-left-color: #16a34a; }
        .summary-card.notready { border-left-color: #f97316; }
        }
        .summary-label {
            color: #616e7c;
            margin-bottom: 2px;
        }
        .summary-value {
            font-size: 16px;
            font-weight: 700;
            color: #102a43;
        }
        .card {
            background: #ffffff;
            border-radius: 6px;
            box-shadow: 0 1px 3px rgba(15, 23, 42, 0.15);
            padding: 8px;
        }
        .chart-container {
            flex: 0 0 260px;
        }
        .table-container {
            flex: 1;
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
        }
        th {
            background-color: #e5edf7;
            font-weight: 600;
            color: #25313d;
        }
        tbody tr:nth-child(even) td {
            background-color: #f8fafc;
        }
        tbody tr:hover td {
            background-color: #e5f0ff;
        }
        .region-cell {
            text-align: left;
            font-weight: 600;
            padding-left: 6px;
        }
        .number-cell {
            font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        }
        .install-table {
            margin-top: 8px;
        }
        .starlink-sep {
            border-left: 3px solid #000;
        }
        .group-header {
            background-color: #dbeafe;
        }
        tfoot th {
            background-color: #dbeafe;
        }
        .meta-line {
            font-size: 11px;
            color: #6b7280;
            margin-bottom: 4px;
        }
    </style>
</head>
<body>
    <div class="page">
        <h1>LEO / Starlink Monitoring</h1>
        <div class="summary-bar">
            <div class="summary-card total">
                <div class="summary-label">Total Sites</div>
                <div class="summary-value">{{ total_sites }}</div>
            </div>
            <div class="summary-card install">
                <div class="summary-label">For Installation</div>
                <div class="summary-value">{{ total_installation }}</div>
            </div>
            <div class="summary-card installed">
                <div class="summary-label">Installed</div>
                <div class="summary-value">{{ total_installed }}</div>
            </div>
            <div class="summary-card notready">
                <div class="summary-label">Not Ready</div>
                <div class="summary-value">{{ total_final_notready }}</div>
            </div>
        </div>
        <div class="meta-line">
            Auto-refresh: 60s &nbsp;•&nbsp; Last update: {{ last_updated }}
        </div>

        <div class="layout">
        <div class="chart-container card">
            <div>
                <h3>Final Status (Overall)</h3>
                <canvas id="finalChart" width="220" height="220"></canvas>
            </div>
            <div style="margin-top: 20px;">
                <h3>Starlink Status (Overall)</h3>
                <canvas id="starlinkChart" width="220" height="220"></canvas>
            </div>
        </div>
        <div class="table-container">
            <div class="card">
                <h2>Deployment & Starlink Status by Region</h2>
                <table>
                    <thead>
                        <tr class="group-header">
                            <th rowspan="2">Region</th>
                            <th colspan="5">Final Status</th>
                            <th class="starlink-sep" colspan="4">Starlink Status</th>
                        </tr>
                        <tr>
                            <th>Ready</th>
                            <th>For Removal</th>
                            <th>For Replacement</th>
                            <th>Not Ready</th>
                            <th>Blank</th>
                            <th class="starlink-sep">For Delivery</th>
                            <th>For Installation</th>
                            <th>Installed</th>
                            <th>Blank</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for i in range(regions|length) %}
                        <tr>
                            <td class="region-cell">{{ regions[i] }}</td>
                            <td class="number-cell">{{ final_ready_vals[i] }}</td>
                            <td class="number-cell">{{ final_removal_vals[i] }}</td>
                            <td class="number-cell">{{ final_replacement_vals[i] }}</td>
                            <td class="number-cell">{{ final_notready_vals[i] }}</td>
                            <td class="number-cell">{{ final_blank_vals[i] }}</td>
                            <td class="starlink-sep number-cell">{{ delivery_vals[i] }}</td>
                            <td class="number-cell">{{ installation_vals[i] }}</td>
                            <td class="number-cell">{{ installed_vals[i] }}</td>
                            <td class="number-cell">{{ star_blank_vals[i] }}</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                    <tfoot>
                        <tr>
                            <th>Total</th>
                            <th>{{ total_final_ready }}</th>
                            <th>{{ total_final_removal }}</th>
                            <th>{{ total_final_replacement }}</th>
                            <th>{{ total_final_notready }}</th>
                            <th>{{ total_final_blank }}</th>
                            <th class="starlink-sep">{{ total_delivery }}</th>
                            <th>{{ total_installation }}</th>
                            <th>{{ total_installed }}</th>
                            <th>{{ total_star_blank }}</th>
                        </tr>
                    </tfoot>
                </table>
            </div>

            <div class="install-table card">
                <h2>Installation Schedule (For Installation)</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Region</th>
                            <th>Installation Date</th>
                            <th>Count (For Installation)</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for i in range(install_regions|length) %}
                        <tr>
                            <td class="region-cell">{{ install_regions[i] }}</td>
                            <td>{{ install_dates[i] }}</td>
                            <td class="number-cell">{{ install_counts[i] }}</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script>
        const finalLabels = ["Ready to Deploy", "For Removal", "For Replacement", "Not Ready", "Blank"];
        const finalData = [
            {{ total_final_ready|default(0) }},
            {{ total_final_removal|default(0) }},
            {{ total_final_replacement|default(0) }},
            {{ total_final_notready|default(0) }},
            {{ total_final_blank|default(0) }},
        ];

        const starLabels = ["For Delivery", "For Installation", "Installed", "Blank"];
        const starData = [
            {{ total_delivery|default(0) }},
            {{ total_installation|default(0) }},
            {{ total_installed|default(0) }},
            {{ total_star_blank|default(0) }},
        ];

        const finalCtx = document.getElementById('finalChart').getContext('2d');
        const starCtx = document.getElementById('starlinkChart').getContext('2d');

        new Chart(finalCtx, {
            type: 'pie',
            data: {
                labels: finalLabels,
                datasets: [{
                    data: finalData,
                    backgroundColor: [
                        'rgba(34, 197, 94, 0.8)',   // Ready
                        'rgba(239, 68, 68, 0.8)',   // Removal
                        'rgba(251, 146, 60, 0.8)',  // Replacement
                        'rgba(234, 179, 8, 0.8)',   // Not Ready
                        'rgba(148, 163, 184, 0.8)', // Blank
                    ],
                    borderWidth: 0
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: { boxWidth: 10, font: { size: 10 } }
                    }
                }
            }
        });

        new Chart(starCtx, {
            type: 'pie',
            data: {
                labels: starLabels,
                datasets: [{
                    data: starData,
                    backgroundColor: [
                        'rgba(59, 130, 246, 0.8)',  // For Delivery
                        'rgba(96, 165, 250, 0.8)', // For Installation
                        'rgba(37, 99, 235, 0.8)',  // Installed
                        'rgba(148, 163, 184, 0.8)',// Blank
                    ],
                    borderWidth: 0
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: { boxWidth: 10, font: { size: 10 } }
                    }
                }
            }
        });
    </script>
</body>
</html>
"""


@app.route("/")
def index():
    final_pivot, star_pivot, install_summary = get_pivots()

    if final_pivot.empty and star_pivot.empty:
        regions = []
        final_ready_vals = final_removal_vals = final_replacement_vals = final_notready_vals = final_blank_vals = []
        delivery_vals = installation_vals = installed_vals = star_blank_vals = []
        total_final_ready = total_final_removal = total_final_replacement = total_final_notready = total_final_blank = 0
        total_delivery = total_installation = total_installed = total_star_blank = 0
        total_sites = 0
    else:
        regions = final_pivot.index.tolist()

        final_ready_vals = final_pivot["Ready to Deploy"].tolist()
        final_removal_vals = final_pivot["For Removal"].tolist()
        final_replacement_vals = final_pivot["For Replacement"].tolist()
        final_notready_vals = final_pivot["Not Ready"].tolist()
        final_blank_vals = final_pivot["Blank"].tolist()

        delivery_vals = star_pivot["For Delivery"].tolist()
        installation_vals = star_pivot["For Installation"].tolist()
        installed_vals = star_pivot["Installed"].tolist()
        star_blank_vals = star_pivot["Blank"].tolist()

        total_final_ready = final_pivot["Ready to Deploy"].sum()
        total_final_removal = final_pivot["For Removal"].sum()
        total_final_replacement = final_pivot["For Replacement"].sum()
        total_final_notready = final_pivot["Not Ready"].sum()
        total_final_blank = final_pivot["Blank"].sum()
        total_sites = (
            total_final_ready
            + total_final_removal
            + total_final_replacement
            + total_final_notready
            + total_final_blank
        )

        total_delivery = star_pivot["For Delivery"].sum()
        total_installation = star_pivot["For Installation"].sum()
        total_installed = star_pivot["Installed"].sum()
        total_star_blank = star_pivot["Blank"].sum()

    if install_summary is None or install_summary.empty:
        install_regions = []
        install_dates = []
        install_counts = []
    else:
        install_regions = install_summary["Region"].astype(str).tolist()
        install_dates = install_summary["Starlink Installation Date"].astype(str).tolist()
        install_counts = install_summary["Count"].astype(int).tolist()

    return render_template_string(
        TEMPLATE,
        regions=regions,
        final_ready_vals=final_ready_vals,
        final_removal_vals=final_removal_vals,
        final_replacement_vals=final_replacement_vals,
        final_notready_vals=final_notready_vals,
        final_blank_vals=final_blank_vals,
        delivery_vals=delivery_vals,
        installation_vals=installation_vals,
        installed_vals=installed_vals,
        star_blank_vals=star_blank_vals,
        total_final_ready=total_final_ready,
        total_final_removal=total_final_removal,
        total_final_replacement=total_final_replacement,
        total_final_notready=total_final_notready,
        total_final_blank=total_final_blank,
        total_delivery=total_delivery,
        total_installation=total_installation,
        total_installed=total_installed,
        total_star_blank=total_star_blank,
        total_sites=total_sites,
        install_regions=install_regions,
        install_dates=install_dates,
        install_counts=install_counts,
        last_updated=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )


if __name__ == "__main__":
    # Access in browser at http://127.0.0.1:5000
    app.run(debug=True, host="0.0.0.0", port=5000)
