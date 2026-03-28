from qualification_importer import run_import
from web_import_app import create_import_app, run_dev_server


app = create_import_app(
    import_fn=run_import,
    template_name="qualification_importer.html",
    allowed_suffixes={".xlsx"},
)


if __name__ == "__main__":
    run_dev_server(app, default_port=5001, label="NEMO qualification importer")
