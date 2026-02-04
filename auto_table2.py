from io import BytesIO
from datetime import datetime

from flask import Flask, render_template_string, request, send_file

from auto_table_core import TEMPLATE, get_table_data
from api.index import _build_workbook

app = Flask(__name__)


@app.route("/")
def index():
    selected_region = request.args.get("region", "").strip() or None
    selected_schedule = request.args.get("schedule", "").strip() or None
    selected_installation = request.args.get("installation", "").strip() or None
    selected_tile = request.args.get("tile", "").strip() or None
    selected_lot = request.args.get("lot", "").strip() or None
    selected_search = request.args.get("search", "").strip() or None
    include_unscheduled = request.args.get("full", "") == "1"

    rows, region_options, schedule_options, installation_options, stats = get_table_data(
        selected_region,
        selected_schedule,
        selected_installation,
        selected_tile,
        selected_lot,
        include_unscheduled=include_unscheduled,
        selected_search=selected_search,
    )

    # Handle XLSX download when the report form is submitted
    if request.args.get("download") == "xlsx":
        selected_columns = request.args.getlist("col")
        include_stats = request.args.get("include_stats", "1") == "1"
        filters = {
            "region": selected_region,
            "schedule": selected_schedule,
            "installation": selected_installation,
            "tile": selected_tile,
            "lot": selected_lot,
        }
        wb = _build_workbook(rows, stats, selected_columns, include_stats, filters)
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        stamp = datetime.now().strftime("%Y%m%d-%H%M")
        lot_tag = ""
        if selected_lot:
            lot_tag = "-" + selected_lot.lower().replace(" ", "").replace("#", "")
        filename = f"monitoring-report{lot_tag}-{stamp}.xlsx"
        return send_file(
            buf,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    show_report = request.args.get("report", "") == "1"

    return render_template_string(
        TEMPLATE,
        rows=rows,
        region_options=region_options,
        schedule_options=schedule_options,
        selected_region=selected_region or "",
        selected_schedule=selected_schedule or "",
        installation_options=installation_options,
        selected_installation=selected_installation or "",
        selected_tile=selected_tile or "",
        show_report=show_report,
        selected_lot=selected_lot or "",
        selected_search=selected_search or "",
        stats=stats,
        include_unscheduled=include_unscheduled,
        last_updated=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )


if __name__ == "__main__":
    # Access in browser at http://127.0.0.1:5000
    app.run(debug=True, host="0.0.0.0", port=5000)
