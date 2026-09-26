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
