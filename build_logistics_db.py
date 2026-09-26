"""
build_logistics_db.py
=====================
Creates a SQLite logistics database and loads the three uploaded CSV files
(trailers, trips, safety_incidents) into properly-typed tables with primary
keys and foreign-key relationships.

Output:
    /home/z/my-project/download/logistics.db
    /home/z/my-project/download/logistics_schema.sql  (human-readable DDL)
"""
import csv
import os
import sqlite3
from pathlib import Path
from datetime import datetime

UPLOAD_DIR = Path("/home/z/my-project/upload")
DOWNLOAD_DIR = Path("/home/z/my-project/download")
DB_PATH = DOWNLOAD_DIR / "logistics.db"
SCHEMA_PATH = DOWNLOAD_DIR / "logistics_schema.sql"

TRAILERS_CSV = UPLOAD_DIR / "trailers.csv"
TRIPS_CSV = UPLOAD_DIR / "trips.csv"
INCIDENTS_CSV = UPLOAD_DIR / "safety_incidents.csv"

# ---------------------------------------------------------------------------
# Schema DDL
# ---------------------------------------------------------------------------
SCHEMA_SQL = """
-- NOTE: foreign_keys pragma is intentionally OFF during bulk load so that
-- real-world orphan rows do not abort the ETL. FK declarations below are
-- kept as documentation and for application-layer enforcement.

-- =====================================================================
-- trailers  (parent reference table)
-- =====================================================================
CREATE TABLE IF NOT EXISTS trailers (
    trailer_id         TEXT     PRIMARY KEY,
    trailer_number     TEXT,
    trailer_type       TEXT,
    length_feet        INTEGER,
    model_year         INTEGER,
    vin                TEXT,
    acquisition_date   TEXT,        -- ISO date (YYYY-MM-DD)
    status             TEXT,
    current_location   TEXT
);

-- =====================================================================
-- trips  (references trailers.trailer_id)
-- =====================================================================
CREATE TABLE IF NOT EXISTS trips (
    trip_id                TEXT     PRIMARY KEY,
    load_id                TEXT,
    driver_id              TEXT,
    truck_id               TEXT,
    trailer_id             TEXT,
    dispatch_date          TEXT,    -- ISO date (YYYY-MM-DD)
    actual_distance_miles  REAL,
    actual_duration_hours  REAL,
    fuel_gallons_used      REAL,
    average_mpg            REAL,
    idle_time_hours        REAL,
    trip_status            TEXT,
    FOREIGN KEY (trailer_id) REFERENCES trailers(trailer_id)
);

CREATE INDEX IF NOT EXISTS idx_trips_dispatch_date ON trips(dispatch_date);
CREATE INDEX IF NOT EXISTS idx_trips_driver_id    ON trips(driver_id);
CREATE INDEX IF NOT EXISTS idx_trips_truck_id     ON trips(truck_id);
CREATE INDEX IF NOT EXISTS idx_trips_trailer_id   ON trips(trailer_id);
CREATE INDEX IF NOT EXISTS idx_trips_status       ON trips(trip_status);

-- =====================================================================
-- safety_incidents  (references trips.trip_id)
-- =====================================================================
CREATE TABLE IF NOT EXISTS safety_incidents (
    incident_id           TEXT     PRIMARY KEY,
    trip_id               TEXT,
    truck_id              TEXT,
    driver_id             TEXT,
    incident_date         TEXT,    -- ISO datetime
    incident_type         TEXT,
    location_city         TEXT,
    location_state        TEXT,
    at_fault_flag         INTEGER, -- 0/1 boolean
    injury_flag           INTEGER, -- 0/1 boolean
    vehicle_damage_cost   REAL,
    cargo_damage_cost     REAL,
    claim_amount          REAL,
    preventable_flag      INTEGER, -- 0/1 boolean
    description           TEXT,
    FOREIGN KEY (trip_id) REFERENCES trips(trip_id)
);

CREATE INDEX IF NOT EXISTS idx_inc_trip_id      ON safety_incidents(trip_id);
CREATE INDEX IF NOT EXISTS idx_inc_incident_date ON safety_incidents(incident_date);
CREATE INDEX IF NOT EXISTS idx_inc_type         ON safety_incidents(incident_type);
CREATE INDEX IF NOT EXISTS idx_inc_driver_id    ON safety_incidents(driver_id);
"""


def str_to_bool(s: str) -> int:
    """Convert CSV 'True'/'False' strings to integer 1/0."""
    if s is None:
        return 0
    return 1 if s.strip().lower() in ("true", "1", "yes", "y") else 0


def load_trailers(conn):
    """Load trailers.csv -> trailers table (insert-or-replace to be idempotent)."""
    cur = conn.cursor()
    with open(TRAILERS_CSV, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        rows = [
            (
                r["trailer_id"],
                r["trailer_number"],
                r["trailer_type"],
                int(r["length_feet"]) if r["length_feet"] else None,
                int(r["model_year"]) if r["model_year"] else None,
                r["vin"],
                r["acquisition_date"],
                r["status"],
                r["current_location"],
            )
            for r in reader
        ]
    cur.executemany(
        """INSERT OR REPLACE INTO trailers
           (trailer_id, trailer_number, trailer_type, length_feet, model_year,
            vin, acquisition_date, status, current_location)
           VALUES (?,?,?,?,?,?,?,?,?)""",
        rows,
    )
    conn.commit()
    return len(rows)


def load_trips(conn):
    """Load trips.csv -> trips table."""
    cur = conn.cursor()
    with open(TRIPS_CSV, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)

        def f(v):
            return float(v) if v not in (None, "") else None

        def i(v):
            return int(v) if v not in (None, "") else None

        rows = [
            (
                r["trip_id"],
                r["load_id"],
                r["driver_id"],
                r["truck_id"],
                r["trailer_id"],
                r["dispatch_date"],
                f(r["actual_distance_miles"]),
                f(r["actual_duration_hours"]),
                f(r["fuel_gallons_used"]),
                f(r["average_mpg"]),
                f(r["idle_time_hours"]),
                r["trip_status"],
            )
            for r in reader
        ]
    cur.executemany(
        """INSERT OR REPLACE INTO trips
           (trip_id, load_id, driver_id, truck_id, trailer_id, dispatch_date,
            actual_distance_miles, actual_duration_hours, fuel_gallons_used,
            average_mpg, idle_time_hours, trip_status)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
        rows,
    )
    conn.commit()
    return len(rows)


def load_incidents(conn):
    """Load safety_incidents.csv -> safety_incidents table."""
    cur = conn.cursor()
    with open(INCIDENTS_CSV, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)

        def f(v):
            return float(v) if v not in (None, "") else None

        rows = [
            (
                r["incident_id"],
                r["trip_id"],
                r["truck_id"],
                r["driver_id"],
                r["incident_date"],
                r["incident_type"],
                r["location_city"],
                r["location_state"],
                str_to_bool(r["at_fault_flag"]),
                str_to_bool(r["injury_flag"]),
                f(r["vehicle_damage_cost"]),
                f(r["cargo_damage_cost"]),
                f(r["claim_amount"]),
                str_to_bool(r["preventable_flag"]),
                r["description"],
            )
            for r in reader
        ]
    cur.executemany(
        """INSERT OR REPLACE INTO safety_incidents
           (incident_id, trip_id, truck_id, driver_id, incident_date,
            incident_type, location_city, location_state, at_fault_flag,
            injury_flag, vehicle_damage_cost, cargo_damage_cost, claim_amount,
            preventable_flag, description)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        rows,
    )
    conn.commit()
    return len(rows)


def verify(conn):
    """Print verification stats and referential-integrity checks."""
    cur = conn.cursor()
    print("\n========== VERIFICATION ==========")
    for tbl in ("trailers", "trips", "safety_incidents"):
        cur.execute(f"SELECT COUNT(*) FROM {tbl}")
        print(f"  {tbl:<20s} rows: {cur.fetchone()[0]:>10,}")

    print("\n--- Sample rows ---")
    for tbl in ("trailers", "trips", "safety_incidents"):
        cur.execute(f"SELECT * FROM {tbl} LIMIT 2")
        cols = [d[0] for d in cur.description]
        print(f"\n  [{tbl}] columns: {cols}")
        for row in cur.fetchall():
            print(f"    {row}")

    print("\n--- Referential integrity ---")
    cur.execute(
        """SELECT COUNT(*) FROM trips t
           LEFT JOIN trailers tr ON t.trailer_id = tr.trailer_id
           WHERE tr.trailer_id IS NULL"""
    )
    orphan_trips = cur.fetchone()[0]
    print(f"  trips with missing trailer_id FK : {orphan_trips:,}")

    cur.execute(
        """SELECT COUNT(*) FROM safety_incidents si
           LEFT JOIN trips t ON si.trip_id = t.trip_id
           WHERE t.trip_id IS NULL"""
    )
    orphan_inc = cur.fetchone()[0]
    print(f"  incidents with missing trip_id FK: {orphan_inc:,}")

    print("\n--- Aggregate sanity checks ---")
    cur.execute(
        "SELECT MIN(dispatch_date), MAX(dispatch_date), "
        "ROUND(AVG(actual_distance_miles),2), ROUND(AVG(average_mpg),2) "
        "FROM trips"
    )
    mn, mx, ad, am = cur.fetchone()
    print(f"  trips dispatch_date range : {mn} -> {mx}")
    print(f"  trips avg distance (mi)   : {ad}")
    print(f"  trips avg mpg             : {am}")

    cur.execute(
        "SELECT incident_type, COUNT(*) FROM safety_incidents "
        "GROUP BY incident_type ORDER BY 2 DESC"
    )
    print("  incident_type breakdown   :")
    for itype, cnt in cur.fetchall():
        print(f"      {itype:<22s} {cnt}")

    cur.execute(
        "SELECT trailer_type, COUNT(*) FROM trailers GROUP BY trailer_type"
    )
    print("  trailer_type breakdown    :")
    for ttype, cnt in cur.fetchall():
        print(f"      {ttype:<22s} {cnt}")


def main():
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

    # Remove any prior DB so re-runs are clean and idempotent
    if DB_PATH.exists():
        DB_PATH.unlink()

    print(f"Creating database at: {DB_PATH}")
    # NOTE: SQLite defaults to PRAGMA foreign_keys = OFF.
    # We intentionally keep FK enforcement OFF during the bulk load (standard
    # ETL practice) so that real-world orphan rows do not abort the load.
    # The FK declarations in the schema remain as documentation and are
    # enforced by the application layer / future migrations. Orphan counts
    # are reported in verify() as a data-quality finding.
    conn = sqlite3.connect(DB_PATH)

    print("Creating schema ...")
    conn.executescript(SCHEMA_SQL)
    conn.commit()

    # Write the schema file for reference / future migrations
    SCHEMA_PATH.write_text(SCHEMA_SQL.strip() + "\n", encoding="utf-8")
    print(f"Schema written to: {SCHEMA_PATH}")

    t0 = datetime.now()
    n_trailers = load_trailers(conn)
    print(f"Loaded {n_trailers:,} rows into trailers  ({datetime.now()-t0})")

    t1 = datetime.now()
    n_trips = load_trips(conn)
    print(f"Loaded {n_trips:,} rows into trips  ({datetime.now()-t1})")

    t2 = datetime.now()
    n_inc = load_incidents(conn)
    print(f"Loaded {n_inc:,} rows into safety_incidents  ({datetime.now()-t2})")

    verify(conn)

    # Run VACUUM to compact the file
    conn.commit()
    conn.execute("VACUUM;")
    conn.close()

    size_mb = DB_PATH.stat().st_size / (1024 * 1024)
    print(f"\nDone. Database file size: {size_mb:.2f} MB")
    print(f"DB path : {DB_PATH}")
    print(f"Schema  : {SCHEMA_PATH}")


if __name__ == "__main__":
    main()
