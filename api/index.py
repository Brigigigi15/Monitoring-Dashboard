from datetime import datetime

from flask import Flask, render_template_string, request

from auto_table_core import TEMPLATE, get_table_data

app = Flask(__name__)


@app.route("/", defaults={"path": ""}, methods=["GET"])
@app.route("/<path:path>", methods=["GET"])
def index(path: str):
    # Single handler for any path within this function (/, /api, /api/index, etc.)
    selected_region = request.args.get("region", "").strip() or None
    rows, region_options = get_table_data(selected_region)
    return render_template_string(
        TEMPLATE,
        rows=rows,
        region_options=region_options,
        selected_region=selected_region or "",
        last_updated=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )


# On Vercel, the `app` object is used as the WSGI entrypoint.
