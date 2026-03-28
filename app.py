from app_qualification_importer import app


if __name__ == "__main__":
    from web_import_app import run_dev_server

    run_dev_server(app, default_port=5001, label="NEMO qualification importer")
