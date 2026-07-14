import sqlite3
import hashlib
import secrets
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

DB_PATH = Path(__file__).with_name("silverline.db")
ADMIN_USERNAME = "admin"
DEFAULT_ADMIN_PASSWORD = "admin123"
COMMISSION_DEFAULT = 1000

CATEGORIES = {
    "ac": "AC technician",
    "property": "Property dealer",
    "grocery": "Grocery store",
    "gym": "Gym",
    "barber": "Barber shop",
    "carwash": "Car wash/detailing",
    "society": "Housing society stall",
    "employer": "Employer partnership",
}
PRODUCTS = [
    "Split AC 1.5 Ton",
    "Refrigerator 18cft",
    "Washing Machine",
    "Microwave Oven",
    "LED TV 43 inch",
    "Water Dispenser",
    "Air Fryer",
    "Blender",
    "Other",
]
STATUSES = ["Pending", "Closed", "Lost"]

st.set_page_config(page_title="Silver Line Portal", page_icon="⚡", layout="wide")

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .stApp {
        background: radial-gradient(circle at top left, #283447 0, #121821 34%, #0b0f15 100%);
        color:#eef3f8;
    }
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #101720 0%, #171f2b 100%);
        border-right: 1px solid rgba(255,255,255,0.08);
    }
    [data-testid="stSidebar"] .stButton>button {
        border-radius: 12px;
        border: 1px solid rgba(255,255,255,0.12);
        background: rgba(255,255,255,0.04);
    }
    .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
    .hero-card {
        padding: 28px 30px;
        border-radius: 24px;
        background: linear-gradient(135deg, rgba(255,255,255,0.10), rgba(255,255,255,0.035));
        border: 1px solid rgba(255,255,255,0.12);
        box-shadow: 0 18px 55px rgba(0,0,0,0.28);
        margin-bottom: 22px;
    }
    .hero-title { font-size: 2.15rem; font-weight: 800; margin: 0; letter-spacing: -0.04em; }
    .hero-subtitle { color:#aab6c4; margin-top: 8px; font-size: 1rem; }
    .role-pill {
        display:inline-block; padding: 7px 12px; border-radius: 999px;
        background: rgba(245, 192, 80, 0.16); color:#ffd27a;
        border: 1px solid rgba(245, 192, 80, 0.28); font-weight: 700; font-size: .82rem;
    }
    div[data-testid="stMetric"] {
        background: linear-gradient(135deg, rgba(255,255,255,.09), rgba(255,255,255,.035));
        border:1px solid rgba(255,255,255,.10);
        border-radius:18px;
        padding:18px;
        box-shadow: 0 14px 36px rgba(0,0,0,0.22);
    }
    div[data-testid="stMetric"] label { color:#aab6c4 !important; }
    div[data-testid="stMetric"] [data-testid="stMetricValue"] { font-weight:800; color:#ffffff; }
    .silver-card {
        background: rgba(18,24,33,0.72);
        border:1px solid rgba(255,255,255,0.10);
        border-radius:18px;
        padding:20px;
        margin-bottom:16px;
        box-shadow: 0 14px 32px rgba(0,0,0,0.18);
    }
    .section-title { font-size:1.15rem; font-weight:800; margin-bottom: 6px; }
    .small-muted { color:#aab6c4; font-size:0.92rem; }
    .status-closed { color:#68d391; font-weight:800; }
    .status-pending { color:#f6c453; font-weight:800; }
    .status-lost { color:#ff7a6b; font-weight:800; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] {
        background: rgba(255,255,255,0.055);
        border-radius: 999px;
        padding: 10px 18px;
        border: 1px solid rgba(255,255,255,0.08);
    }
    .stTabs [aria-selected="true"] { background: rgba(245, 192, 80, 0.18); color:#ffd27a; }
    .stTextInput input, .stNumberInput input, .stTextArea textarea { border-radius: 12px; }
    .stSelectbox div[data-baseweb="select"] > div { border-radius: 12px; }
    .stButton>button, .stDownloadButton>button, .stFormSubmitButton>button {
        border-radius: 12px; font-weight: 700;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def hero(title: str, subtitle: str, role: str | None = None):
    role_html = f'<span class="role-pill">{role}</span>' if role else ""
    st.markdown(
        f"""
        <div class="hero-card">
            {role_html}
            <h1 class="hero-title">{title}</h1>
            <div class="hero-subtitle">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section_card(title: str, subtitle: str = ""):
    st.markdown(
        f'<div class="silver-card"><div class="section-title">{title}</div><div class="small-muted">{subtitle}</div></div>',
        unsafe_allow_html=True,
    )


def format_currency(value: float) -> str:
    return f"Rs {value:,.0f}"


def prepare_referrals_table(data: pd.DataFrame) -> pd.DataFrame:
    if data.empty:
        return data
    shown = data.copy()
    for col in ["product_amount", "commission_amount", "commission"]:
        if col in shown.columns:
            shown[col] = shown[col].apply(lambda x: format_currency(float(x or 0)))
    return shown


def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def hash_password(password: str, salt: str | None = None) -> tuple[str, str]:
    salt = salt or secrets.token_hex(16)
    hashed = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 100_000).hex()
    return salt, hashed


def verify_password(password: str, salt: str, stored_hash: str) -> bool:
    _, check_hash = hash_password(password, salt)
    return secrets.compare_digest(check_hash, stored_hash)


def run_query(query: str, params: tuple = (), fetch: bool = False):
    with get_conn() as conn:
        cur = conn.execute(query, params)
        conn.commit()
        if fetch:
            return cur.fetchall()
        return None


def generate_partner_code(category: str, name: str) -> str:
    prefix = {
        "ac": "AC",
        "property": "PR",
        "grocery": "GR",
        "gym": "GY",
        "barber": "BB",
        "carwash": "CW",
        "society": "HS",
        "employer": "EP",
    }.get(category, "PT")
    clean = "".join(ch for ch in name.upper() if ch.isalnum())[:6] or "PARTNR"
    count = run_query("SELECT COUNT(*) AS total FROM partners", fetch=True)[0]["total"] + 1
    return f"SL-{prefix}-{clean}{count:02d}"


def create_tables():
    run_query(
        """
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
        """
    )
    run_query(
        """
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
        """
    )
    run_query(
        """
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
        """
    )


def migrate_database():
    """Allow older local databases to support the new branch_manager role."""
    with get_conn() as conn:
        user_table = conn.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='users'").fetchone()
        if user_table and "branch_manager" not in user_table["sql"]:
            conn.execute("ALTER TABLE users RENAME TO users_old")
            conn.execute(
                """
                CREATE TABLE users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT UNIQUE NOT NULL,
                    password_salt TEXT NOT NULL,
                    password_hash TEXT NOT NULL,
                    role TEXT NOT NULL CHECK(role IN ('admin','partner','branch_manager')),
                    partner_id INTEGER,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY(partner_id) REFERENCES partners(id)
                );
                """
            )
            conn.execute(
                """
                INSERT INTO users(id,username,password_salt,password_hash,role,partner_id,created_at)
                SELECT id,username,password_salt,password_hash,role,partner_id,created_at FROM users_old
                """
            )
            conn.execute("DROP TABLE users_old")
            conn.commit()


def seed_data():
    admin_exists = run_query("SELECT id FROM users WHERE username=?", (ADMIN_USERNAME,), fetch=True)
    if not admin_exists:
        salt, hashed = hash_password(DEFAULT_ADMIN_PASSWORD)
        run_query(
            "INSERT INTO users(username,password_salt,password_hash,role) VALUES(?,?,?,?)",
            (ADMIN_USERNAME, salt, hashed, "admin"),
        )

    manager_exists = run_query("SELECT id FROM users WHERE username=?", ("manager",), fetch=True)
    if not manager_exists:
        salt, hashed = hash_password("manager123")
        run_query(
            "INSERT INTO users(username,password_salt,password_hash,role) VALUES(?,?,?,?)",
            ("manager", salt, hashed, "branch_manager"),
        )

    partners_exist = run_query("SELECT COUNT(*) AS total FROM partners", fetch=True)[0]["total"]
    if partners_exist == 0:
        demo_partners = [
            ("Imran Cool Care", "ac", "0301-2223344", "Gulberg", "Imran", "partner001"),
            ("Al-Habib Estates", "property", "0321-5556677", "DHA", "Habib", "partner002"),
            ("Al-Fateh Grocers", "grocery", "0345-1231234", "Model Town", "Fateh", "partner003"),
            ("Iron Gym Gulberg", "gym", "0322-7897890", "Gulberg", "Manager", "partner004"),
            ("Style Cutz Barber", "barber", "0334-6546543", "Johar Town", "Ali", "partner005"),
        ]
        for name, category, phone, area, person, password in demo_partners:
            code = generate_partner_code(category, name)
            run_query(
                "INSERT INTO partners(name,category,code,phone,area,contact_person,joined_date) VALUES(?,?,?,?,?,?,?)",
                (name, category, code, phone, area, person, str(date.today())),
            )
            partner_id = run_query("SELECT id FROM partners WHERE code=?", (code,), fetch=True)[0]["id"]
            salt, hashed = hash_password(password)
            run_query(
                "INSERT INTO users(username,password_salt,password_hash,role,partner_id) VALUES(?,?,?,?,?)",
                (code, salt, hashed, "partner", partner_id),
            )

        demo_refs = [
            (1, "Ahmed Raza", "0300-1111111", "Split AC 1.5 Ton", 125000, "Closed"),
            (1, "Sana Malik", "0300-2222222", "Refrigerator 18cft", 155000, "Pending"),
            (2, "Bilal Sheikh", "0300-3333333", "Washing Machine", 95000, "Closed"),
            (3, "Fatima Noor", "0300-4444444", "LED TV 43 inch", 105000, "Lost"),
            (4, "Usman Tariq", "0300-5555555", "Microwave Oven", 45000, "Closed"),
            (5, "Ayesha Khan", "0300-6666666", "Water Dispenser", 38000, "Pending"),
        ]
        for partner_id, customer, phone, product, amount, status in demo_refs:
            commission = 0 if status == "Lost" else COMMISSION_DEFAULT
            run_query(
                """
                INSERT INTO referrals(partner_id,customer_name,customer_phone,product,product_amount,commission_amount,status,referral_date,notes,added_by)
                VALUES(?,?,?,?,?,?,?,?,?,?)
                """,
                (partner_id, customer, phone, product, amount, commission, status, str(date.today()), "Demo referral", "seed"),
            )


def init_db():
    create_tables()
    migrate_database()
    seed_data()


def df(query: str, params: tuple = ()) -> pd.DataFrame:
    with get_conn() as conn:
        return pd.read_sql_query(query, conn, params=params)


def login_screen():
    left, right = st.columns([1.2, 0.8], gap="large")
    with left:
        st.markdown("""
        <div class="hero-card" style="min-height:360px; display:flex; flex-direction:column; justify-content:center;">
            <span class="role-pill">RD Electronics Referral System</span>
            <h1 class="hero-title">Silver Line Partner Portal</h1>
            <div class="hero-subtitle">Track referrals, manage partners, and monitor branch performance from one clean dashboard.</div>
            <br>
            <div class="small-muted">Admin creates partners and managers. Branch managers add referrals and view the directory. Partners submit their own customer leads.</div>
        </div>
        """, unsafe_allow_html=True)
    with right:
        st.markdown('<div class="silver-card">', unsafe_allow_html=True)
        st.subheader("Sign in")
        st.caption("Use your admin, branch manager, or partner credentials.")
        username = st.text_input("Username / Partner Code", placeholder="admin, manager, or SL-...")
        password = st.text_input("Password", type="password", placeholder="Enter password")
        if st.button("Login", type="primary", use_container_width=True):
            rows = run_query("SELECT * FROM users WHERE username=?", (username.strip(),), fetch=True)
            if rows and verify_password(password, rows[0]["password_salt"], rows[0]["password_hash"]):
                st.session_state.user = dict(rows[0])
                st.rerun()
            else:
                st.error("Wrong username or password")
        st.markdown('</div>', unsafe_allow_html=True)

        with st.expander("Demo logins"):
            st.write("Admin: `admin` / `admin123`")
            st.write("Branch manager: `manager` / `manager123`")
            st.write("Partner examples: partner code shown in Admin panel / passwords `partner001`, `partner002`, etc.")

def logout_button():
    user = st.session_state.get("user", {})
    role_label = str(user.get("role", "")).replace("_", " ").title()
    st.sidebar.markdown("### ⚡ Silver Line")
    st.sidebar.caption("Partner Referral Portal")
    st.sidebar.divider()
    if user:
        st.sidebar.markdown(f"**Logged in as**  \n`{user.get('username', '')}`")
        st.sidebar.markdown(f"**Role**  \n{role_label}")
    st.sidebar.divider()
    if st.sidebar.button("Logout", use_container_width=True):
        st.session_state.clear()
        st.rerun()


def add_referral_form(partner_id: int, added_by: str, form_key: str):
    with st.form(form_key, clear_on_submit=True):
        c1, c2 = st.columns(2)
        customer = c1.text_input("Customer name *")
        phone = c2.text_input("Customer phone")
        product = c1.selectbox("Product buying *", PRODUCTS)
        amount = c2.number_input("Product amount (Rs)", min_value=0.0, step=1000.0)
        status = c1.selectbox("Status", STATUSES)
        commission = c2.number_input("Commission (Rs)", min_value=0.0, value=float(COMMISSION_DEFAULT), step=500.0)
        referral_date = c1.date_input("Referral date", value=date.today())
        notes = st.text_area("Notes")
        submit = st.form_submit_button("Save referral", type="primary")
    if submit:
        if not customer.strip():
            st.error("Customer name is required")
        else:
            final_commission = 0 if status == "Lost" else commission
            run_query(
                """
                INSERT INTO referrals(partner_id,customer_name,customer_phone,product,product_amount,commission_amount,status,referral_date,notes,added_by)
                VALUES(?,?,?,?,?,?,?,?,?,?)
                """,
                (partner_id, customer.strip(), phone.strip(), product, amount, final_commission, status, str(referral_date), notes, added_by),
            )
            st.success("Referral saved")
            st.rerun()


def admin_dashboard():
    hero("Admin Portal", "Create partners, branch managers, referrals, and monitor complete performance.", "Admin")
    logout_button()

    partners = df("SELECT * FROM partners ORDER BY id DESC")
    referrals = df(
        """
        SELECT r.*, p.name AS partner_name, p.category, p.code
        FROM referrals r JOIN partners p ON r.partner_id=p.id
        ORDER BY r.id DESC
        """
    )

    total_partners = len(partners[partners["is_active"] == 1]) if not partners.empty else 0
    total_referrals = len(referrals)
    closed = referrals[referrals["status"] == "Closed"] if not referrals.empty else pd.DataFrame()
    pending = referrals[referrals["status"] == "Pending"] if not referrals.empty else pd.DataFrame()
    conversion = round((len(closed) / total_referrals) * 100, 1) if total_referrals else 0

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Active partners", total_partners)
    m2.metric("Total referrals", total_referrals)
    m3.metric("Conversion", f"{conversion}%")
    m4.metric("Commission paid", f"Rs {closed['commission_amount'].sum():,.0f}" if not closed.empty else "Rs 0")
    m5.metric("Pending payout", f"Rs {pending['commission_amount'].sum():,.0f}" if not pending.empty else "Rs 0")

    tab1, tab2, tab3, tab4 = st.tabs(["Dashboard", "Add Referral", "Partners", "All Referrals"])

    with tab1:
        c1, c2 = st.columns(2)
        if not referrals.empty:
            category_chart = referrals.groupby("category").size().reset_index(name="referrals")
            category_chart["category"] = category_chart["category"].map(CATEGORIES).fillna(category_chart["category"])
            c1.subheader("Referrals by category")
            c1.bar_chart(category_chart, x="category", y="referrals")

            leaderboard = referrals.groupby(["partner_id", "partner_name"]).agg(
                referrals=("id", "count"),
                closed=("status", lambda s: (s == "Closed").sum()),
                commission=("commission_amount", "sum"),
            ).reset_index().sort_values("closed", ascending=False)
            c2.subheader("Partner leaderboard")
            c2.dataframe(leaderboard[["partner_name", "referrals", "closed", "commission"]], use_container_width=True, hide_index=True)
        else:
            st.info("No referrals yet.")

    with tab2:
        section_card("Admin add referral", "Choose a partner and enter customer lead details.")
        if partners.empty:
            st.warning("Add a partner first.")
        else:
            partner_labels = {f"{row['name']} — {row['code']}": int(row["id"]) for _, row in partners.iterrows()}
            selected = st.selectbox("Select partner", list(partner_labels.keys()))
            add_referral_form(partner_labels[selected], "admin", "admin_referral_form")

    with tab3:
        section_card("Add new partner", "Only admin can create new partner accounts.")
        with st.form("add_partner", clear_on_submit=True):
            c1, c2 = st.columns(2)
            name = c1.text_input("Partner business/name *")
            category = c2.selectbox("Category", list(CATEGORIES.keys()), format_func=lambda x: CATEGORIES[x])
            phone = c1.text_input("Phone")
            area = c2.text_input("Area")
            contact_person = c1.text_input("Contact person")
            password = c2.text_input("Partner password *", type="password")
            submitted = st.form_submit_button("Create partner login", type="primary")
        if submitted:
            if not name.strip() or not password.strip():
                st.error("Partner name and password are required")
            else:
                code = generate_partner_code(category, name)
                run_query(
                    "INSERT INTO partners(name,category,code,phone,area,contact_person,joined_date) VALUES(?,?,?,?,?,?,?)",
                    (name.strip(), category, code, phone.strip(), area.strip(), contact_person.strip(), str(date.today())),
                )
                partner_id = run_query("SELECT id FROM partners WHERE code=?", (code,), fetch=True)[0]["id"]
                salt, hashed = hash_password(password)
                run_query(
                    "INSERT INTO users(username,password_salt,password_hash,role,partner_id) VALUES(?,?,?,?,?)",
                    (code, salt, hashed, "partner", partner_id),
                )
                st.success(f"Partner created. Username/code: {code}")
                st.rerun()

        st.divider()
        section_card("Create branch manager login", "Branch managers can add referrals and view directories, but cannot add partners.")
        with st.form("add_branch_manager", clear_on_submit=True):
            c1, c2 = st.columns(2)
            manager_username = c1.text_input("Branch manager username *")
            manager_password = c2.text_input("Branch manager password *", type="password")
            manager_submitted = st.form_submit_button("Create branch manager", type="secondary")
        if manager_submitted:
            if not manager_username.strip() or not manager_password.strip():
                st.error("Username and password are required")
            else:
                exists = run_query("SELECT id FROM users WHERE username=?", (manager_username.strip(),), fetch=True)
                if exists:
                    st.error("This username already exists")
                else:
                    salt, hashed = hash_password(manager_password)
                    run_query(
                        "INSERT INTO users(username,password_salt,password_hash,role) VALUES(?,?,?,?)",
                        (manager_username.strip(), salt, hashed, "branch_manager"),
                    )
                    st.success(f"Branch manager created. Username: {manager_username.strip()}")
                    st.rerun()

        section_card("Partner directory", "View all active and inactive partner details.")
        show = partners.copy()
        if not show.empty:
            show["category"] = show["category"].map(CATEGORIES).fillna(show["category"])
            st.dataframe(show[["id", "name", "category", "code", "phone", "area", "contact_person", "joined_date", "is_active"]], use_container_width=True, hide_index=True)

    with tab4:
        section_card("All referrals", "Filter and review referral records.")
        if not referrals.empty:
            status_filter = st.multiselect("Filter status", STATUSES, default=STATUSES)
            shown = referrals[referrals["status"].isin(status_filter)]
            st.dataframe(prepare_referrals_table(shown[["id", "partner_name", "code", "customer_name", "customer_phone", "product", "product_amount", "commission_amount", "status", "referral_date", "notes"]]), use_container_width=True, hide_index=True)
            csv = shown.to_csv(index=False).encode("utf-8")
            st.download_button("Download CSV", csv, "silverline_referrals.csv", "text/csv")
        else:
            st.info("No referrals yet.")


def branch_manager_dashboard():
    hero("Branch Manager Portal", "Add referrals, view partner directory, and monitor sales performance. Partner creation remains admin-only.", "Branch Manager")
    logout_button()

    partners = df("SELECT * FROM partners ORDER BY id DESC")
    referrals = df(
        """
        SELECT r.*, p.name AS partner_name, p.category, p.code
        FROM referrals r JOIN partners p ON r.partner_id=p.id
        ORDER BY r.id DESC
        """
    )

    total_partners = len(partners[partners["is_active"] == 1]) if not partners.empty else 0
    total_referrals = len(referrals)
    closed = referrals[referrals["status"] == "Closed"] if not referrals.empty else pd.DataFrame()
    pending = referrals[referrals["status"] == "Pending"] if not referrals.empty else pd.DataFrame()
    conversion = round((len(closed) / total_referrals) * 100, 1) if total_referrals else 0

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Active partners", total_partners)
    m2.metric("Total referrals", total_referrals)
    m3.metric("Conversion", f"{conversion}%")
    m4.metric("Commission paid", f"Rs {closed['commission_amount'].sum():,.0f}" if not closed.empty else "Rs 0")
    m5.metric("Pending payout", f"Rs {pending['commission_amount'].sum():,.0f}" if not pending.empty else "Rs 0")

    tab1, tab2, tab3, tab4 = st.tabs(["Dashboard", "Add Referral", "Partner Directory", "All Referrals"])

    with tab1:
        c1, c2 = st.columns(2)
        if not referrals.empty:
            category_chart = referrals.groupby("category").size().reset_index(name="referrals")
            category_chart["category"] = category_chart["category"].map(CATEGORIES).fillna(category_chart["category"])
            c1.subheader("Referrals by category")
            c1.bar_chart(category_chart, x="category", y="referrals")

            leaderboard = referrals.groupby(["partner_id", "partner_name"]).agg(
                referrals=("id", "count"),
                closed=("status", lambda s: (s == "Closed").sum()),
                commission=("commission_amount", "sum"),
            ).reset_index().sort_values("closed", ascending=False)
            c2.subheader("Partner leaderboard")
            c2.dataframe(leaderboard[["partner_name", "referrals", "closed", "commission"]], use_container_width=True, hide_index=True)
        else:
            st.info("No referrals yet.")

    with tab2:
        section_card("Add referral", "Select the partner and submit customer lead details.")
        if partners.empty:
            st.warning("No partners available. Ask admin to create partners first.")
        else:
            active_partners = partners[partners["is_active"] == 1] if "is_active" in partners.columns else partners
            partner_labels = {f"{row['name']} — {row['code']}": int(row["id"]) for _, row in active_partners.iterrows()}
            selected = st.selectbox("Select partner", list(partner_labels.keys()))
            add_referral_form(partner_labels[selected], st.session_state.user["username"], "manager_referral_form")

    with tab3:
        section_card("Partner directory", "View all active and inactive partner details.")
        if partners.empty:
            st.info("No partners found.")
        else:
            show = partners.copy()
            show["category"] = show["category"].map(CATEGORIES).fillna(show["category"])
            st.dataframe(show[["id", "name", "category", "code", "phone", "area", "contact_person", "joined_date", "is_active"]], use_container_width=True, hide_index=True)

    with tab4:
        section_card("All referrals", "Filter and review referral records.")
        if not referrals.empty:
            status_filter = st.multiselect("Filter status", STATUSES, default=STATUSES, key="manager_status_filter")
            shown = referrals[referrals["status"].isin(status_filter)]
            st.dataframe(prepare_referrals_table(shown[["id", "partner_name", "code", "customer_name", "customer_phone", "product", "product_amount", "commission_amount", "status", "referral_date", "notes", "added_by"]]), use_container_width=True, hide_index=True)
        else:
            st.info("No referrals yet.")


def partner_dashboard():
    user = st.session_state.user
    partner_id = int(user["partner_id"])
    partner = run_query("SELECT * FROM partners WHERE id=?", (partner_id,), fetch=True)[0]
    refs = df("SELECT * FROM referrals WHERE partner_id=? ORDER BY id DESC", (partner_id,))

    hero(f"Partner Portal — {partner['name']}", f"Referral code: {partner['code']} • Submit and track your customer referrals.", "Partner")
    logout_button()

    closed = refs[refs["status"] == "Closed"] if not refs.empty else pd.DataFrame()
    pending = refs[refs["status"] == "Pending"] if not refs.empty else pd.DataFrame()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("My referrals", len(refs))
    c2.metric("Closed sales", len(closed))
    c3.metric("Commission earned", f"Rs {closed['commission_amount'].sum():,.0f}" if not closed.empty else "Rs 0")
    c4.metric("Pending payout", f"Rs {pending['commission_amount'].sum():,.0f}" if not pending.empty else "Rs 0")

    tab1, tab2 = st.tabs(["Add Referral", "My Referrals"])
    with tab1:
        section_card("Add customer referral", "Enter the customer and product details.")
        add_referral_form(partner_id, partner["code"], "partner_referral_form")
    with tab2:
        section_card("Referral history", "Track your submitted leads and commission status.")
        if refs.empty:
            st.info("You have not added any referrals yet.")
        else:
            st.dataframe(prepare_referrals_table(refs[["customer_name", "customer_phone", "product", "product_amount", "commission_amount", "status", "referral_date", "notes"]]), use_container_width=True, hide_index=True)


def main():
    init_db()
    if "user" not in st.session_state:
        login_screen()
        return
    role = st.session_state.user["role"]
    if role == "admin":
        admin_dashboard()
    elif role == "branch_manager":
        branch_manager_dashboard()
    else:
        partner_dashboard()


if __name__ == "__main__":
    main()
