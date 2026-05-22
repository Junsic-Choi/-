DROP TABLE IF EXISTS plans;
CREATE TABLE plans (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    equipment TEXT NOT NULL,
    weekId TEXT NOT NULL DEFAULT '2026-W08',
    manager TEXT,
    model TEXT,
    partName TEXT,
    partNo TEXT,
    mon TEXT,
    tue TEXT,
    wed TEXT,
    thu TEXT,
    fri TEXT,
    sat TEXT,
    sun TEXT,
    mon_act TEXT DEFAULT '',
    tue_act TEXT DEFAULT '',
    wed_act TEXT DEFAULT '',
    thu_act TEXT DEFAULT '',
    fri_act TEXT DEFAULT '',
    sat_act TEXT DEFAULT '',
    sun_act TEXT DEFAULT '',
    updatedAt DATETIME DEFAULT CURRENT_TIMESTAMP
);

DROP TABLE IF EXISTS equipment_holidays;
CREATE TABLE equipment_holidays (
    equipment TEXT NOT NULL,
    weekId TEXT NOT NULL,
    mon INTEGER DEFAULT 0,
    tue INTEGER DEFAULT 0,
    wed INTEGER DEFAULT 0,
    thu INTEGER DEFAULT 0,
    fri INTEGER DEFAULT 0,
    sat INTEGER DEFAULT 0,
    sun INTEGER DEFAULT 0,
    PRIMARY KEY (equipment, weekId)
);

CREATE INDEX IF NOT EXISTS idx_plans_eq_week ON plans(equipment, weekId);
CREATE INDEX IF NOT EXISTS idx_plans_weekId ON plans(weekId);
CREATE INDEX IF NOT EXISTS idx_plans_manager ON plans(manager);
CREATE INDEX IF NOT EXISTS idx_plans_eq_part ON plans(equipment, partNo);
CREATE INDEX IF NOT EXISTS idx_plans_weekId_equipment ON plans(weekId, equipment);
