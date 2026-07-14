-- Silver Line Partner Referral Portal Database Schema
-- Streamlit app uses SQLite locally. For Streamlit Cloud, the database file is created automatically.

CREATE TABLE IF NOT EXISTS partners (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    category TEXT NOT NULL,
    code TEXT UNIQUE NOT NULL,
    phone TEXT,
    area TEXT,
    contact_person TEXT,
    joined_date TEXT NOT NULL,
    is_active INTEGER DEFAULT 1
);

CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    password_salt TEXT NOT NULL,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL CHECK(role IN ('admin','partner','branch_manager')),
    partner_id INTEGER,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(partner_id) REFERENCES partners(id)
);

CREATE TABLE IF NOT EXISTS referrals (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    partner_id INTEGER NOT NULL,
    customer_name TEXT NOT NULL,
    customer_phone TEXT,
    product TEXT NOT NULL,
    product_amount REAL DEFAULT 0,
    commission_amount REAL DEFAULT 1000,
    status TEXT DEFAULT 'Pending' CHECK(status IN ('Pending','Closed','Lost')),
    referral_date TEXT NOT NULL,
    notes TEXT,
    added_by TEXT NOT NULL,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(partner_id) REFERENCES partners(id)
);
