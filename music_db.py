import argparse
import csv
import os
import sqlite3
import sys
from contextlib import closing
from pathlib import Path

from docx import Document

Album = tuple[str, int | None, str | None]
Artist = tuple[str, str | None, int]

SAMPLE_ARTISTS = (
    ("The Beatles", "Rock"),
    ("Adele", "Pop"),
    ("Miles Davis", "Jazz"),
)
SAMPLE_ALBUMS = (
    ("Abbey Road", 1969, "The Beatles"),
    ("Revolver", 1966, "The Beatles"),
    ("21", 2011, "Adele"),
    ("25", 2015, "Adele"),
    ("Kind of Blue", 1959, "Miles Davis"),
)


def initialize_database(connection: sqlite3.Connection) -> None:
    """Create the tables and add sample records without duplicating them."""
    connection.execute("""
        CREATE TABLE IF NOT EXISTS artists (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            genre TEXT
        )
    """)
    connection.execute("""
        CREATE TABLE IF NOT EXISTS albums (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            year INTEGER,
            artist_id INTEGER,
            FOREIGN KEY (artist_id) REFERENCES artists(id)
        )
    """)

    artist_ids: dict[str, int] = {}
    for name, genre in SAMPLE_ARTISTS:
        row = connection.execute(
            "SELECT id FROM artists WHERE name = ? ORDER BY id LIMIT 1", (name,)
        ).fetchone()
        if row is None:
            cursor = connection.execute(
                "INSERT INTO artists (name, genre) VALUES (?, ?)", (name, genre)
            )
            if cursor.lastrowid is None:
                raise sqlite3.DatabaseError(f"Could not retrieve ID for artist {name!r}")
            artist_ids[name] = cursor.lastrowid
        else:
            artist_ids[name] = row[0]

    for title, year, artist_name in SAMPLE_ALBUMS:
        artist_id = artist_ids[artist_name]
        connection.execute("""
            INSERT INTO albums (title, year, artist_id)
            SELECT ?, ?, ?
            WHERE NOT EXISTS (
                SELECT 1 FROM albums
                WHERE title = ? AND year = ? AND artist_id = ?
            )
        """, (title, year, artist_id, title, year, artist_id))


def read_results(connection: sqlite3.Connection) -> tuple[list[Album], list[Artist]]:
    albums = connection.execute("""
        SELECT albums.title, albums.year, artists.name
        FROM albums
        LEFT JOIN artists ON albums.artist_id = artists.id
        ORDER BY albums.id
    """).fetchall()
    artists = connection.execute("""
        SELECT artists.name, artists.genre, COUNT(albums.id)
        FROM artists
        LEFT JOIN albums ON albums.artist_id = artists.id
        GROUP BY artists.id, artists.name, artists.genre
        ORDER BY artists.id
    """).fetchall()
    return albums, artists


def export_results(directory: Path, albums: list[Album], artists: list[Artist]) -> tuple[Path, Path]:
    csv_path = directory / "music_results.csv"
    docx_path = directory / "music_results.docx"

    with csv_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(["Album", "Year", "Artist"])
        writer.writerows(albums)

    document = Document()
    document.add_heading("Music database report", level=0)
    document.add_heading("Summary", level=1)
    document.add_paragraph(
        f"The database contains {len(albums)} albums and {len(artists)} artists."
    )
    years = [year for _, year, _ in albums if year is not None]
    if years:
        document.add_paragraph(f"Album release years range from {min(years)} to {max(years)}.")

    document.add_heading("Artists", level=1)
    table = document.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    for cell, label in zip(table.rows[0].cells, ("Artist", "Genre", "Albums")):
        cell.text = label
    for name, genre, album_count in artists:
        cells = table.add_row().cells
        cells[0].text = name
        cells[1].text = genre or "Not specified"
        cells[2].text = str(album_count)

    document.add_heading("Albums", level=1)
    table = document.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    for cell, label in zip(table.rows[0].cells, ("Album", "Year", "Artist")):
        cell.text = label
    for title, year, artist_name in albums:
        cells = table.add_row().cells
        cells[0].text = title
        cells[1].text = str(year) if year is not None else "Not specified"
        cells[2].text = artist_name or "Unknown artist"
    document.save(str(docx_path))
    return csv_path, docx_path


def open_reports(paths: tuple[Path, Path]) -> bool:
    if not hasattr(os, "startfile"):
        print("Automatic opening is only supported on Windows; use --no-open.", file=sys.stderr)
        return False

    success = True
    for path in paths:
        try:
            os.startfile(path)
        except OSError as error:
            print(f"Could not open {path}: {error}", file=sys.stderr)
            success = False
    return success


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--no-open", action="store_true", help="create reports without opening them")
    args = parser.parse_args(argv)
    directory = Path(__file__).resolve().parent

    try:
        with closing(sqlite3.connect(directory / "music.db")) as connection:
            connection.execute("PRAGMA foreign_keys = ON")
            with connection:
                initialize_database(connection)
                albums, artists = read_results(connection)

        for title, year, artist_name in albums:
            print(f"{title} ({year}) belongs to {artist_name or 'Unknown artist'}.")
        paths = export_results(directory, albums, artists)
        for path in paths:
            print(f"Saved {path}")
        return 0 if args.no_open or open_reports(paths) else 1
    except (sqlite3.Error, OSError) as error:
        print(f"Could not create music reports: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
