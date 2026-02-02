from flask import Flask, render_template_string
import pandas as pd

app = Flask(__name__)

# If app.py is in the same folder as the Excel file, just use the filename:
FILE = r"C:\Users\iOne3\Desktop\DepEd\Monitoring\Managed Service LEO x SOLAR (DepEd-Funded) - masterlist as of 23Jan2026.xlsx"


def get_pivots():
    # Read only the "LEO SOLAR" sheet and focus on the needed columns
    sheet = pd.read_excel(FILE, sheet_name="LEO SOLAR")
    # Note: "Final Status " in the file has a trailing space
    required_cols = ["Region", "Final Status ", "Starlink Status", "Starlink Installation Date"]
    missing = [c for c in required_cols if c not in sheet.columns]
    if missing:
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
        install_df["Starlink Installation Date"] = pd.to_datetime(
            install_df["Starlink Installation Date"], errors="coerce"
        ).dt.date
        install_df = install_df.dropna(subset=["Starlink Installation Date"])
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
        body { font-family: Arial, sans-serif; margin: 10px; font-size: 12px; }
        h1 { margin: 4px 0 6px 0; font-size: 18px; }
        h2 { margin: 10px 0 4px 0; font-size: 14px; }
        h3 { margin: 6px 0 4px 0; font-size: 13px; }
        .layout { display: flex; align-items: flex-start; gap: 16px; margin-top: 8px; }
        .chart-container { flex: 0 0 260px; }
        .table-container { flex: 1; }
        table { border-collapse: collapse; width: 100%; table-layout: fixed; }
        th, td { border: 1px solid #ccc; padding: 2px 4px; text-align: center; }
        th { background-color: #f5f5f5; }
        .starlink-sep { border-left: 3px solid #000; }
        .install-table { margin-top: 8px; }
    </style>
</head>
<body>
    <h1>LEO / Starlink Monitoring</h1>

    <div class="layout">
        <div class="chart-container">
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
            <table>
                <tr>
                    <th>Region</th>
                    <th>Ready to Deploy</th>
                    <th>For Removal</th>
                    <th>For Replacement</th>
                    <th>Not Ready</th>
                    <th>Blank (Final)</th>
                    <th class="starlink-sep">For Delivery</th>
                    <th>For Installation</th>
                    <th>Installed</th>
                    <th>Blank (Starlink)</th>
                </tr>
                {% for i in range(regions|length) %}
                <tr>
                    <td>{{ regions[i] }}</td>
                    <td>{{ final_ready_vals[i] }}</td>
                    <td>{{ final_removal_vals[i] }}</td>
                    <td>{{ final_replacement_vals[i] }}</td>
                    <td>{{ final_notready_vals[i] }}</td>
                    <td>{{ final_blank_vals[i] }}</td>
                    <td class="starlink-sep">{{ delivery_vals[i] }}</td>
                    <td>{{ installation_vals[i] }}</td>
                    <td>{{ installed_vals[i] }}</td>
                    <td>{{ star_blank_vals[i] }}</td>
                </tr>
                {% endfor %}
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
            </table>
        </div>
    </div>

    <div class="install-table">
        <h2>Installation Schedule (For Installation)</h2>
        <table>
            <tr>
                <th>Region</th>
                <th>Installation Date</th>
                <th>Count (For Installation)</th>
            </tr>
            {% for i in range(install_regions|length) %}
            <tr>
                <td>{{ install_regions[i] }}</td>
                <td>{{ install_dates[i] }}</td>
                <td>{{ install_counts[i] }}</td>
            </tr>
            {% endfor %}
        </table>
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
                        'rgba(54, 162, 235, 0.7)',
                        'rgba(255, 99, 132, 0.7)',
                        'rgba(255, 206, 86, 0.7)',
                        'rgba(153, 102, 255, 0.7)',
                        'rgba(201, 203, 207, 0.7)',
                    ],
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: { position: 'right' }
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
                        'rgba(75, 192, 192, 0.7)',
                        'rgba(255, 159, 64, 0.7)',
                        'rgba(54, 162, 235, 0.7)',
                        'rgba(201, 203, 207, 0.7)',
                    ],
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: { position: 'right' }
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
        install_regions=install_regions,
        install_dates=install_dates,
        install_counts=install_counts,
    )


if __name__ == "__main__":
    # Access in browser at http://127.0.0.1:5000
    app.run(debug=True, host="0.0.0.0", port=5000)
