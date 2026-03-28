import io
import os
import tempfile
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path

import requests
from flask import Flask, render_template, request

from nemo_user_importer import EXPECTED_HEADERS, run_import


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024


def allowed_filename(filename: str) -> bool:
    suffix = Path(filename).suffix.lower()
    return suffix in {".xlsx", ".csv"}


@app.route("/", methods=["GET", "POST"])
def index():
    result = None
    error = None
    status = None

    if request.method == "POST":
        token = request.form.get("token", "").strip()
        dry_run = request.form.get("dry_run") == "on"
        spreadsheet = request.files.get("spreadsheet")

        if not token:
            error = "Enter your NEMO API token."
        elif not spreadsheet or not spreadsheet.filename:
            error = "Choose an Excel or CSV spreadsheet to upload."
        elif not allowed_filename(spreadsheet.filename):
            error = "Upload a .xlsx or .csv file."
        else:
            suffix = Path(spreadsheet.filename).suffix.lower()
            temp_path = None
            output = io.StringIO()

            try:
                with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as handle:
                    spreadsheet.save(handle)
                    temp_path = Path(handle.name)

                with redirect_stdout(output), redirect_stderr(output):
                    print(
                        "Run started via web app.\n"
                        f"Uploaded file: {spreadsheet.filename}\n"
                        f"Mode: {'Dry Run' if dry_run else 'Live Import'}\n"
                    )
                    run_import(str(temp_path), token, dry_run=dry_run)

                status = "success"
                result = output.getvalue().strip()
            except requests.HTTPError as error_obj:
                response = error_obj.response
                details = response.text if response is not None else str(error_obj)
                error = f"API request failed.\n\n{details}"
                status = "error"
                result = output.getvalue().strip()
            except Exception as error_obj:
                error = str(error_obj)
                status = "error"
                result = output.getvalue().strip()
            finally:
                if temp_path and temp_path.exists():
                    temp_path.unlink()

    return render_template(
        "index.html",
        error=error,
        expected_headers=sorted(EXPECTED_HEADERS),
        result=result,
        status=status,
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    print(f"NEMO web app starting on http://127.0.0.1:{port}")
    app.run(debug=True, use_reloader=False, host="0.0.0.0", port=port)
