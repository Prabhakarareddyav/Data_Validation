# mRNA Data Validator

Compares an Excel source file with a CSV target (for example a Snowflake export): matching keys, keys missing/extra in the target, and column-by-column value mismatches. Exports a colour-coded Excel report.

Files are uploaded through the browser, so it works on any machine (no server file paths needed).

## Run locally (Windows / macOS / Linux)

Requires Python 3.9+.

    git clone https://github.com/<your-user>/mrna-validator.git
    cd mrna-validator
    # Windows: double-click start.bat      macOS/Linux: ./start.sh

Then open http://localhost:5000

Manual: `pip install -r requirements.txt && python app.py`

## Run with Docker

    docker build -t mrna-validator .
    docker run -p 5000:5000 mrna-validator

## Host for a team
Deploy the Dockerfile on Render, Railway, Azure App Service, etc. (start command: `gunicorn app:app`).
**Data note:** uploaded files pass through the server. For confidential R&D data, host inside your company network, not on a public service.

## Settings
| Setting | Meaning |
|---|---|
| Header row | 0-based row holding the source column names (default 3) |
| Key column | Unique ID used to match rows (default Antigen_ID) |
| Ignore columns | Extra columns to skip; `RECORD_COUNT` and `UNNAMED*` always skipped |
| Skip first column | Drops the first source column (as in the original script) |

Values are compared after trimming whitespace, treating blanks/NaN as equal and `1` = `1.0`.
