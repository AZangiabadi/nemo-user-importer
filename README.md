# NEMO Import Tools

This project now has separate web apps for the two main workflows:

- `app_user_importer.py` for accounts, PI users, projects, and other users
- `app_qualification_importer.py` for qualification uploads

There are also desktop scripts:

- `nemo_user_importer.py`
- `Qualifications entry.py`

## Run Locally With uv

Create the environment and install dependencies:

```bash
uv venv
source .venv/bin/activate
uv sync
```

Run the user importer web app:

```bash
uv run python app_user_importer.py
```

Run the qualification importer web app:

```bash
uv run python app_qualification_importer.py
```

`app.py` is kept as a compatibility entrypoint for the qualification importer.

## Deploy Online

Pick the app you want to deploy:

- User importer: `gunicorn app_user_importer:app`
- Qualification importer: `gunicorn app_qualification_importer:app`

Python `3.11+` is supported.
