# Music database

This project creates a SQLite database of artists and albums, prints the albums,
and produces CSV and Word reports. The Word report summarizes the artists,
their genres and album counts, and the range of known release years.

## Requirements

- Python 3.11 or later
- `python-docx` (already listed in the parent project's `requirements.txt`)
- Windows and associated applications for automatically opening the reports

If needed, install the document dependency with `python -m pip install python-docx`.
SQLite is included with Python; a separate SQLite installation is not required.

## Run

From the project directory:

```powershell
python music_db.py
```

The script creates `music.db`, `music_results.csv`, and `music_results.docx`
beside the script. It opens both reports using the default Windows applications.
Use `python music_db.py --no-open` to create the reports without opening them,
including on non-Windows systems.

Existing database records are retained, and the sample artists and albums are
not duplicated on subsequent runs. The CSV and Word reports are **overwritten**
each time to reflect the current database; close them in other applications
first if they are locked. The script reports errors if a file cannot be saved or
opened. A viewer can open `music.db` directly, or use `sqlite3 music.db` if the
SQLite command-line tool is installed.

## Tests

```powershell
python -m unittest discover -s . -p test_music_db.py
```

Tests use a temporary database and do not modify the project's own data.
