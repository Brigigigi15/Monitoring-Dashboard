from datetime import datetime
from io import BytesIO

from flask import Flask, render_template_string, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.chart import PieChart, Reference

from auto_table_core import TEMPLATE, get_table_data

app = Flask(__name__)


def _build_workbook(rows, stats, selected_columns, include_stats, filters):
    """Build an XLSX workbook matching current table + optional stats/charts."""
    wb = Workbook()
    ws_table = wb.active
    ws_table.title = "Table"

    # Default columns if none selected
    default_columns = [
        "Region",
        "Province",
        "BEIS School ID",
        "Schedule",
        "Calendar Status",
        "Start Time",
        "End Time",
        "Installation Status",
        "Starlink Status",
        "Approval",
        "Blocker",
    ]
    columns = selected_columns or default_columns

    # Header style
    header_font = Font(bold=True, color="020617")
    header_fill = PatternFill("solid", fgColor="CBD5F5")
    thin = Side(style="thin", color="CBD5E1")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # Write header row
    for col_idx, col_name in enumerate(columns, start=1):
        cell = ws_table.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center

    # Data rows
    warning_fill = PatternFill("solid", fgColor="FEF9C3")
    critical_fill = PatternFill("solid", fgColor="FEE2E2")

    for row_idx, row in enumerate(rows, start=2):
        star = (row.get("Starlink Status") or "").lower()
        appr = (row.get("Approval") or "").lower()
        row_fill = None
        if "declin" in appr:
            row_fill = critical_fill
        elif star != "activated":
            row_fill = warning_fill

        for col_idx, col_name in enumerate(columns, start=1):
            value = row.get(col_name, "")
            cell = ws_table.cell(row=row_idx, column=col_idx, value=value)
            cell.border = border
            # Align Region / School / Blocker to left, others center
            if col_name in ("Region", "Province", "BEIS School ID", "Blocker", "Installation Status"):
                cell.alignment = left
            else:
                cell.alignment = center
            if row_fill is not None:
                cell.fill = row_fill

    # Auto-fit-ish column widths based on header length
    for col_idx, col_name in enumerate(columns, start=1):
        width = max(10, min(30, len(str(col_name)) + 4))
        ws_table.column_dimensions[ws_table.cell(row=1, column=col_idx).column_letter].width = width

    if include_stats:
        ws_stats = wb.create_sheet(title="Summary")
        ws_stats["A1"] = "Filters"
        ws_stats["A1"].font = Font(bold=True)
        filters_map = [
            ("Region", filters.get("region") or "All"),
            ("Schedule", filters.get("schedule") or "All"),
            ("Installation", filters.get("installation") or "All"),
            ("Tile", filters.get("tile") or "None"),
        ]
        for idx, (label, value) in enumerate(filters_map, start=2):
            ws_stats[f"A{idx}"] = label
            ws_stats[f"B{idx}"] = value

        # Report generation timestamp
        ts_row = 2 + len(filters_map)
        ws_stats[f"A{ts_row}"] = "Generated at"
        ws_stats[f"B{ts_row}"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Basic stats table
        ws_stats["A7"] = "Metric"
        ws_stats["B7"] = "Value"
        ws_stats["A7"].font = ws_stats["B7"].font = Font(bold=True)
        metrics = [
            ("Starlink Activated", stats.get("star_activated", 0)),
            ("Starlink Not Activated", stats.get("star_not_activated", 0)),
            ("Approval Accepted", stats.get("approval_accepted", 0)),
            ("Approval Pending/Blank", stats.get("approval_pending", 0)),
            ("Approval Decline/Other", stats.get("approval_decline", 0)),
            ("Calendar Sent", stats.get("calendar_sent", 0)),
            ("Calendar Invite Not Sent", stats.get("calendar_not_sent", 0)),
            ("S1 - Installed (Success)", stats.get("s1_success", 0)),
        ]
        for idx, (label, value) in enumerate(metrics, start=8):
            ws_stats[f"A{idx}"] = label
            ws_stats[f"B{idx}"] = int(value)

        # Starlink pie chart
        ws_stats["D2"] = "Starlink Status"
        ws_stats["D3"] = "Activated"
        ws_stats["E3"] = int(stats.get("star_activated", 0))
        ws_stats["D4"] = "Not Activated"
        ws_stats["E4"] = int(stats.get("star_not_activated", 0))

        star_pie = PieChart()
        star_pie.title = "Starlink Status"
        data = Reference(ws_stats, min_col=5, min_row=3, max_row=4)
        labels = Reference(ws_stats, min_col=4, min_row=3, max_row=4)
        star_pie.add_data(data, titles_from_data=False)
        star_pie.set_categories(labels)
        star_pie.width = 10
        star_pie.height = 6
        ws_stats.add_chart(star_pie, "H2")

        # Approval pie chart
        ws_stats["D8"] = "Approval Status"
        ws_stats["D9"] = "Accepted"
        ws_stats["E9"] = int(stats.get("approval_accepted", 0))
        ws_stats["D10"] = "Pending/Blank"
        ws_stats["E10"] = int(stats.get("approval_pending", 0))
        ws_stats["D11"] = "Decline/Other"
        ws_stats["E11"] = int(stats.get("approval_decline", 0))

        appr_pie = PieChart()
        appr_pie.title = "Approval Status"
        data2 = Reference(ws_stats, min_col=5, min_row=9, max_row=11)
        labels2 = Reference(ws_stats, min_col=4, min_row=9, max_row=11)
        appr_pie.add_data(data2, titles_from_data=False)
        appr_pie.set_categories(labels2)
        appr_pie.width = 10
        appr_pie.height = 6
        ws_stats.add_chart(appr_pie, "H18")

    return wb


@app.route("/", defaults={"path": ""}, methods=["GET"])
@app.route("/<path:path>", methods=["GET"])
def index(path: str):
    # Main dashboard handler (also handles report download when ?download=xlsx)
    selected_region = request.args.get("region", "").strip() or None
    selected_schedule = request.args.get("schedule", "").strip() or None
    selected_installation = request.args.get("installation", "").strip() or None
    selected_tile = request.args.get("tile", "").strip() or None

    (
        rows,
        region_options,
        schedule_options,
        installation_options,
        stats,
    ) = get_table_data(
        selected_region,
        selected_schedule,
        selected_installation,
        selected_tile,
    )

    # If download flag is present, stream XLSX instead of HTML
    if request.args.get("download") == "xlsx":
        selected_columns = request.args.getlist("col")
        include_stats = request.args.get("include_stats", "1") == "1"
        filters = {
            "region": selected_region,
            "schedule": selected_schedule,
            "installation": selected_installation,
            "tile": selected_tile,
        }
        wb = _build_workbook(rows, stats, selected_columns, include_stats, filters)
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(
            buf,
            as_attachment=True,
            download_name="monitoring-report.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    show_report = request.args.get("report", "") == "1"

    return render_template_string(
        TEMPLATE,
        rows=rows,
        region_options=region_options,
        schedule_options=schedule_options,
        installation_options=installation_options,
        selected_region=selected_region or "",
        selected_schedule=selected_schedule or "",
        selected_installation=selected_installation or "",
        selected_tile=selected_tile or "",
        stats=stats,
        show_report=show_report,
        last_updated=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )


# On Vercel, the `app` object is used as the WSGI entrypoint.
