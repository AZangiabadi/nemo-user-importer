from nemo_user_importer import EXPECTED_HEADERS, run_import
from web_import_app import create_import_app, run_dev_server


app = create_import_app(
    import_fn=run_import,
    template_name="user_importer.html",
    allowed_suffixes={".csv", ".xlsx"},
)


@app.context_processor
def inject_template_context() -> dict[str, object]:
    return {"expected_headers": sorted(EXPECTED_HEADERS)}


if __name__ == "__main__":
    run_dev_server(app, default_port=5002, label="NEMO user importer")
