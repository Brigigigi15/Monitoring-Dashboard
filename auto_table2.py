from flask import Flask, render_template_string, request
from datetime import datetime

from auto_table_core import TEMPLATE, get_table_data

app = Flask(__name__)


@app.route("/")
def index():
    selected_region = request.args.get("region", "").strip() or None
    selected_schedule = request.args.get("schedule", "").strip() or None
    rows, region_options, schedule_options, stats = get_table_data(selected_region, selected_schedule)
    return render_template_string(
        TEMPLATE,
        rows=rows,
        region_options=region_options,
        schedule_options=schedule_options,
        selected_region=selected_region or "",
        selected_schedule=selected_schedule or "",
        stats=stats,
        last_updated=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )


if __name__ == "__main__":
    # Access in browser at http://127.0.0.1:5000
    app.run(debug=True, host="0.0.0.0", port=5000)
