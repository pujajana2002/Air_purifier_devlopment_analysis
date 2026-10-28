# Import project data into MySQL

This project imports every `.csv` and `.xlsx` file directly inside `data/` into a
MySQL database. CSV files become tables named after the file; worksheets become
tables named after the workbook (and, when there are multiple worksheets, the
worksheet name).

## Setup

1. Make sure a MySQL server is running and you have a MySQL user with permission
   to create the database and tables.
2. Copy `.env.example` to `.env` and set your MySQL user and password:

   ```powershell
   Copy-Item .env.example .env
   notepad .env
   ```

3. Install the Python packages:

   ```powershell
   uv pip install --python .venv\Scripts\python.exe -r requirements.txt
   ```

4. Run the importer from the project folder:

   ```powershell
   .venv\Scripts\python.exe app.py
   ```

The default input folder is `data/`. To import from another folder, run:

```powershell
.venv\Scripts\python.exe app.py --data-dir "C:\path\to\data"
```

The importer creates the database if it does not exist, infers MySQL column
types from the source values, and inserts rows in batches. Each run replaces
the contents of the corresponding tables so rerunning the import does not
duplicate rows. If an import fails while inserting a table, that table's
previous contents are preserved.

Keep `.env` private; it is excluded from Git by `.gitignore`.
