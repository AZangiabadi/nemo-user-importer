# NEMO User Importer

This project imports accounts, PI users, projects, and additional users into NEMO from an Excel or CSV spreadsheet.

It now supports two interfaces:

- A desktop Tk app in `nemo_user_importer.py`
- A deployable web app in `app.py`

## Spreadsheet Rules

Required headers:

- `name`
- `uni`
- `email`
- `pi`
- `account type`
- `project number`

Important rules:

1. Put the PI row before student rows when introducing a new project.
2. Enter `PI` in the `pi` column for PI rows.
3. If `uni` is blank, the app generates one like `xxjd42`.
4. Use `External Academic`, not `External Academia`.

## Run Locally With uv

Create the environment and install dependencies:

```bash
uv venv
source .venv/bin/activate
uv sync
```

Run the web app:

```bash
uv run python app.py
```

Then open `http://127.0.0.1:8000`.

## Deploy Online

This repo is ready for simple deployment on platforms like Render or Railway.

- Start command: `gunicorn app:app`
- Python version: `3.11` or newer

Because the app asks for a NEMO API token at runtime, you do not need to store the token in the repository.

## Desktop Version

If you still want the local popup workflow:

```bash
uv run python nemo_user_importer.py
```
