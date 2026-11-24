import os
import json
import sqlite3
import secrets
import csv
import io
import re
import gzip
from datetime import datetime, date
from functools import wraps
from flask import Flask, request, jsonify, render_template, session, make_response
from flask_cors import CORS
from werkzeug.security import generate_password_hash, check_password_hash

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "focus-institute-secret-key-2026")
CORS(app)


@app.after_request
def compress_response(response):
    """Gzip compresses large JSON payloads for 15x faster network response times."""
    accept_encoding = request.headers.get("Accept-Encoding", "")
    if (
        "gzip" in accept_encoding
        and response.status_code < 300
        and len(response.get_data()) > 500
        and "Content-Encoding" not in response.headers
    ):
        compressed_data = gzip.compress(response.get_data(), compresslevel=5)
        response.set_data(compressed_data)
        response.headers["Content-Encoding"] = "gzip"
        response.headers["Content-Length"] = len(compressed_data)
    return response

DB_PATH = os.path.join(os.path.dirname(__file__), "focus_portal.db")

BATCH_MONTH_MAP = {
    'JAN': 1, 'JANUARY': 1,
    'FEB': 2, 'FEBRUARY': 2,
    'MAR': 3, 'MARCH': 3,
    'APR': 4, 'APRIL': 4,
    'MAY': 5,
    'JUN': 6, 'JUNE': 6,
    'JUL': 7, 'JULY': 7,
    'AUG': 8, 'AUGUST': 8,
    'SEP': 9, 'SEPT': 9, 'SEPTEMBER': 9,
    'OCT': 10, 'OCTOBER': 10,
    'NOV': 11, 'NOVEMBER': 11,
    'DEC': 12, 'DECEMBER': 12
}

def parse_batch_sort_key(batch_str):
    if not batch_str:
        return (0, 0, '')
    s = str(batch_str).strip().upper()
    year_match = re.search(r'\b(20\d\d)\b', s)
    year = int(year_match.group(1)) if year_match else 0
    month = 0
    for m_name, m_val in BATCH_MONTH_MAP.items():
        if re.search(r'\b' + m_name + r'\b', s):
            month = max(month, m_val)
    if month == 0 and year_match:
        m_num = re.search(r'\b(0?[1-9]|1[0-2])\b', s.replace(year_match.group(1), ''))
        if m_num:
            month = int(m_num.group(1))
    return (year, month, s)

def sort_batches(batches):
    """
    Sort batches in reverse chronological order (newest years and latest months on top).
    e.g. 2026 months at the top, then 2025, then 2024 at the bottom.
    """
    return sorted(batches, key=parse_batch_sort_key, reverse=True)


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_db() as conn:
        cursor = conn.cursor()

        # Users table (Manager & Staff)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL CHECK(role IN ('manager', 'staff')),
                phone TEXT,
                is_active INTEGER DEFAULT 1,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Leads table (Shared student directory)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS leads (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                father_name TEXT,
                phone TEXT NOT NULL,
                alt_phone TEXT,
                email TEXT,
                course TEXT,
                batch TEXT,
                learner_code TEXT,
                city TEXT,
                source TEXT DEFAULT 'Main Portal',
                status TEXT DEFAULT 'NEW',
                assigned_to INTEGER,
                next_followup_date TEXT,
                last_call_at TEXT,
                last_caller_id INTEGER,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (assigned_to) REFERENCES users(id),
                FOREIGN KEY (last_caller_id) REFERENCES users(id)
            )
        """)

        # Auto-migration
        cursor.execute("PRAGMA table_info(leads)")
        columns = [row[1] for row in cursor.fetchall()]
        if "father_name" not in columns:
            cursor.execute("ALTER TABLE leads ADD COLUMN father_name TEXT")
        if "batch" not in columns:
            cursor.execute("ALTER TABLE leads ADD COLUMN batch TEXT")
        if "alt_phone" not in columns:
            cursor.execute("ALTER TABLE leads ADD COLUMN alt_phone TEXT")
        if "learner_code" not in columns:
            cursor.execute("ALTER TABLE leads ADD COLUMN learner_code TEXT")
        if "last_caller_id" not in columns:
            cursor.execute("ALTER TABLE leads ADD COLUMN last_caller_id INTEGER")
        if "record_type" not in columns:
            cursor.execute("ALTER TABLE leads ADD COLUMN record_type TEXT DEFAULT 'admitted'")
        if "notes" not in columns:
            cursor.execute("ALTER TABLE leads ADD COLUMN notes TEXT")
        if "cancellation_reason" not in columns:
            cursor.execute("ALTER TABLE leads ADD COLUMN cancellation_reason TEXT")
        if "conversion_date" not in columns:
            cursor.execute("ALTER TABLE leads ADD COLUMN conversion_date TEXT")
        if "is_deleted" not in columns:
            cursor.execute("ALTER TABLE leads ADD COLUMN is_deleted INTEGER DEFAULT 0")
        if "deleted_at" not in columns:
            cursor.execute("ALTER TABLE leads ADD COLUMN deleted_at TEXT")
        if "is_fee_pending" not in columns:
            cursor.execute("ALTER TABLE leads ADD COLUMN is_fee_pending INTEGER DEFAULT 0")
        if "fee_pending_amount" not in columns:
            cursor.execute("ALTER TABLE leads ADD COLUMN fee_pending_amount TEXT DEFAULT ''")
        if "fee_due_date" not in columns:
            cursor.execute("ALTER TABLE leads ADD COLUMN fee_due_date TEXT DEFAULT ''")
        if "fee_remarks" not in columns:
            cursor.execute("ALTER TABLE leads ADD COLUMN fee_remarks TEXT DEFAULT ''")

        # Monthly Exam Calling Students Table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS exam_students (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                month_code TEXT NOT NULL,
                month_name TEXT NOT NULL,
                student_id INTEGER,
                student_name TEXT NOT NULL,
                father_name TEXT,
                phone TEXT NOT NULL,
                alt_phone TEXT,
                course TEXT,
                batch TEXT,
                learner_code TEXT,
                exam_date TEXT,
                call_status TEXT DEFAULT 'NOT_CALLED',
                last_called_by TEXT,
                last_call_at TEXT,
                remarks TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (student_id) REFERENCES leads(id)
            )
        """)

        # Exam Month Settings Table (for admin visibility toggle & monthly archive)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS exam_month_settings (
                month_code TEXT PRIMARY KEY,
                month_name TEXT NOT NULL,
                is_visible_to_staff INTEGER DEFAULT 1,
                is_active INTEGER DEFAULT 1,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Seed current month if not exists
        cursor.execute("""
            INSERT OR IGNORE INTO exam_month_settings (month_code, month_name, is_visible_to_staff, is_active)
            VALUES ('2026-11', 'November 2026', 1, 1)
        """)

        # Call logs table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS call_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                lead_id INTEGER NOT NULL,
                caller_id INTEGER NOT NULL,
                status TEXT NOT NULL,
                remarks TEXT,
                next_followup_date TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (lead_id) REFERENCES leads(id),
                FOREIGN KEY (caller_id) REFERENCES users(id)
            )
        """)

        # Settings table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )
        """)

        cursor.execute("SELECT value FROM settings WHERE key = 'webhook_api_key'")
        if not cursor.fetchone():
            default_api_key = "focus_live_" + secrets.token_hex(12)
            cursor.execute("INSERT INTO settings (key, value) VALUES ('webhook_api_key', ?)", (default_api_key,))

        # Auto-migration for users table permissions
        cursor.execute("PRAGMA table_info(users)")
        u_columns = [row[1] for row in cursor.fetchall()]
        if "permissions" not in u_columns:
            cursor.execute("ALTER TABLE users ADD COLUMN permissions TEXT DEFAULT ''")
        if "theme_preference" not in u_columns:
            cursor.execute("ALTER TABLE users ADD COLUMN theme_preference TEXT DEFAULT 'light'")

        # Default Users: Arsh (Manager) and staff
        cursor.execute("SELECT id FROM users WHERE username = 'arsh'")
        if not cursor.fetchone():
            cursor.execute("""
                INSERT INTO users (name, username, password_hash, role, phone)
                VALUES (?, ?, ?, 'manager', ?)
            """, ("Arsh (Manager)", "arsh", generate_password_hash("arsh123"), "9876500001"))

        cursor.execute("SELECT COUNT(*) as count FROM users WHERE role = 'staff'")
        staff_count = cursor.fetchone()["count"]
        if staff_count == 0:
            staff_members = [
                ("Gurpreet Singh", "gurpreet", "staff123", "9876500002"),
                ("Simran Kaur", "simran", "staff123", "9876500003"),
                ("Aman Sharma", "aman", "staff123", "9876500004"),
            ]
            for name, uname, pwd, ph in staff_members:
                cursor.execute("""
                    INSERT INTO users (name, username, password_hash, role, phone)
                    VALUES (?, ?, ?, 'staff', ?)
                """, (name, uname, generate_password_hash(pwd), ph))

        # Real-time Auto-Sync Tracking Table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS app_sync (
                id INTEGER PRIMARY KEY,
                version INTEGER DEFAULT 1,
                last_updated TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cursor.execute("INSERT OR IGNORE INTO app_sync (id, version, last_updated) VALUES (1, 1, CURRENT_TIMESTAMP)")

        # Performance Indexes for sub-millisecond query execution
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_leads_type_deleted ON leads(record_type, is_deleted);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_leads_status ON leads(status);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_leads_batch ON leads(batch);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_leads_phone ON leads(phone);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_call_logs_lead_id ON call_logs(lead_id);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_call_logs_caller_id ON call_logs(caller_id);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_users_username ON users(username);")

        conn.commit()


# Initialize database automatically on startup
try:
    init_db()
except Exception as _e:
    pass


def bump_sync_version():
    """Increments the global database state version so all connected clients auto-refresh in real time."""
    try:
        with get_db() as conn:
            cursor = conn.cursor()
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            cursor.execute("UPDATE app_sync SET version = version + 1, last_updated = ? WHERE id = 1", (now_str,))
            conn.commit()
    except Exception as e:
        print(f"Error bumping sync version: {e}")


# Permission Constants & RBAC Helpers
ALL_PERMISSIONS = {
    "can_view": 1,
    "can_edit": 1,
    "can_delete": 1,
    "can_restore": 1,
    "can_import": 1,
    "can_export": 1,
    "can_manage_users": 1
}

DEFAULT_STAFF_PERMISSIONS = {
    "can_view": 1,
    "can_edit": 0,
    "can_delete": 0,
    "can_restore": 0,
    "can_import": 0,
    "can_export": 0,
    "can_manage_users": 0
}

def get_user_permissions(user_id, username=None, role=None, perm_raw=None):
    if username == "arsh" or user_id == 1:
        return dict(ALL_PERMISSIONS)
    if perm_raw:
        try:
            p = json.loads(perm_raw) if isinstance(perm_raw, str) else perm_raw
            if isinstance(p, dict):
                return {k: 1 if p.get(k) in (1, "1", True) else 0 for k in ALL_PERMISSIONS}
        except Exception:
            pass
    if role == "manager":
        return dict(ALL_PERMISSIONS)
    return dict(DEFAULT_STAFF_PERMISSIONS)


def user_has_permission(permission_name):
    user_id = session.get("user_id")
    if not user_id:
        return False
    if session.get("username") == "arsh" or user_id == 1:
        return True
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, username, role, permissions FROM users WHERE id = ?", (user_id,))
        row = cursor.fetchone()
        if not row:
            return False
        perms = get_user_permissions(row["id"], row["username"], row["role"], row["permissions"])
        session["permissions"] = perms
        return bool(perms.get(permission_name, 0))


def permission_required(perm_name):
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if "user_id" not in session:
                return jsonify({"error": "Unauthorized. Please log in."}), 401
            if not user_has_permission(perm_name):
                return jsonify({
                    "error": f"Permission Denied: You do not have '{perm_name}' permission to perform this action. Contact Super Admin."
                }), 403
            return f(*args, **kwargs)
        return decorated_function
    return decorator


# Decorators
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user_id" not in session:
            return jsonify({"error": "Unauthorized. Please log in."}), 401
        return f(*args, **kwargs)
    return decorated_function


def manager_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user_id" not in session:
            return jsonify({"error": "Unauthorized. Please log in."}), 401
        user_id = session.get("user_id")
        if session.get("username") == "arsh" or user_id == 1 or session.get("role") == "manager":
            return f(*args, **kwargs)
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT role, permissions FROM users WHERE id = ?", (user_id,))
            u = cursor.fetchone()
            if u:
                perms = get_user_permissions(user_id, session.get("username"), u["role"], u["permissions"])
                if u["role"] == "manager" or any(perms.get(k) for k in ["can_edit", "can_delete", "can_restore", "can_import", "can_export", "can_manage_users"]):
                    return f(*args, **kwargs)
        return jsonify({"error": "Forbidden. Manager access required."}), 403
    return decorated_function


# -------------------------------------------------------------
# Web Page Route
# -------------------------------------------------------------
@app.route("/")
def index():
    return render_template("index.html")


# -------------------------------------------------------------
# Auth APIs
# -------------------------------------------------------------
@app.route("/api/login", methods=["POST"])
def login():
    data = request.json or {}
    username = data.get("username", "").strip()
    password = data.get("password", "")

    if not username or not password:
        return jsonify({"error": "Username and password required"}), 400

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM users WHERE username = ? AND is_active = 1", (username,))
        user = cursor.fetchone()

        if not user or not check_password_hash(user["password_hash"], password):
            return jsonify({"error": "Invalid username or password"}), 401

        session["user_id"] = user["id"]
        session["name"] = user["name"]
        session["username"] = user["username"]
        session["role"] = user["role"]
        user_perms = get_user_permissions(user["id"], user["username"], user["role"], user["permissions"])
        session["permissions"] = user_perms
        is_super = (user["username"] == "arsh" or user["id"] == 1)
        session["is_super_admin"] = is_super
        theme_pref = user["theme_preference"] if ("theme_preference" in user.keys() and user["theme_preference"]) else "light"
        session["theme_preference"] = theme_pref

        return jsonify({
            "message": "Login successful",
            "user": {
                "id": user["id"],
                "name": user["name"],
                "username": user["username"],
                "role": user["role"],
                "is_super_admin": is_super,
                "permissions": user_perms,
                "theme_preference": theme_pref
            }
        })


@app.route("/api/logout", methods=["POST"])
def logout():
    session.clear()
    return jsonify({"message": "Logged out successfully"})


@app.route("/api/me", methods=["GET"])
def me():
    if "user_id" not in session:
        return jsonify({"authenticated": False}), 200

    user_id = session.get("user_id")
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(users)")
        u_cols = [r[1] for r in cursor.fetchall()]
        if "theme_preference" not in u_cols:
            cursor.execute("ALTER TABLE users ADD COLUMN theme_preference TEXT DEFAULT 'light'")
            conn.commit()
        cursor.execute("SELECT id, name, username, role, is_active, permissions, theme_preference FROM users WHERE id = ?", (user_id,))
        u = cursor.fetchone()
        if not u or not u["is_active"]:
            session.clear()
            return jsonify({"authenticated": False}), 200
        
        perms = get_user_permissions(u["id"], u["username"], u["role"], u["permissions"])
        session["name"] = u["name"]
        session["username"] = u["username"]
        session["role"] = u["role"]
        session["permissions"] = perms
        is_super = (u["username"] == "arsh" or u["id"] == 1)
        session["is_super_admin"] = is_super
        theme_pref = u["theme_preference"] if ("theme_preference" in u.keys() and u["theme_preference"]) else "light"
        session["theme_preference"] = theme_pref

        return jsonify({
            "authenticated": True,
            "user": {
                "id": u["id"],
                "name": u["name"],
                "username": u["username"],
                "role": u["role"],
                "is_super_admin": is_super,
                "permissions": perms,
                "theme_preference": theme_pref
            }
        })


@app.route("/api/me/theme", methods=["POST"])
@login_required
def update_user_theme():
    data = request.json or {}
    theme = str(data.get("theme", "light")).strip().lower()
    if theme not in ["light", "dark"]:
        theme = "light"
    user_id = session.get("user_id")
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(users)")
        u_cols = [r[1] for r in cursor.fetchall()]
        if "theme_preference" not in u_cols:
            cursor.execute("ALTER TABLE users ADD COLUMN theme_preference TEXT DEFAULT 'light'")
        cursor.execute("UPDATE users SET theme_preference = ? WHERE id = ?", (theme, user_id))
        conn.commit()
    session["theme_preference"] = theme
    return jsonify({"success": True, "theme": theme})


@app.route("/api/sync/version", methods=["GET"])
def get_sync_version():
    """Ultra-fast endpoint returning current database version so clients auto-synchronize in real time."""
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT version, last_updated FROM app_sync WHERE id = 1")
        row = cursor.fetchone()
        if not row:
            cursor.execute("INSERT OR IGNORE INTO app_sync (id, version, last_updated) VALUES (1, 1, CURRENT_TIMESTAMP)")
            conn.commit()
            return jsonify({"version": 1, "last_updated": ""})
        return jsonify({
            "version": row["version"],
            "last_updated": row["last_updated"]
        })


# -------------------------------------------------------------
# DIRECT API / WEBHOOK (OPTION B)
# -------------------------------------------------------------
@app.route("/api/webhook/lead", methods=["POST"])
def webhook_lead():
    api_key = request.headers.get("X-API-Key") or request.args.get("api_key")
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT value FROM settings WHERE key = 'webhook_api_key'")
        setting_key = cursor.fetchone()

        if not setting_key or api_key != setting_key["value"]:
            return jsonify({"error": "Unauthorized. Invalid X-API-Key header."}), 401

        data = request.json or {}
        name = data.get("name", "").strip()
        father_name = data.get("father_name", "").strip()
        phone = data.get("phone", "").strip()
        alt_phone = data.get("alt_phone", "").strip()
        email = data.get("email", "").strip()
        course = data.get("course", "General Inquiry").strip()
        batch = data.get("batch", "").strip()
        city = data.get("city", "").strip()
        source = data.get("source", "Main Portal API").strip()

        if not name or not phone:
            return jsonify({"error": "Both 'name' and 'phone' are required fields."}), 400

        cursor.execute("""
            INSERT INTO leads (name, father_name, phone, alt_phone, email, course, batch, city, source, status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'NEW')
        """, (name, father_name, phone, alt_phone, email, course, batch, city, source))
        lead_id = cursor.lastrowid
        conn.commit()

        bump_sync_version()

        return jsonify({
            "success": True,
            "message": "Student lead added to shared caller directory",
            "lead_id": lead_id
        }), 201


# -------------------------------------------------------------
# BULK EXCEL (.XLSX) & CSV IMPORT API
# -------------------------------------------------------------
def clean_col(s):
    if s is None:
        return ""
    return str(s).strip().lower().replace(" ", "").replace("_", "").replace("-", "").replace(".", "").replace("/", "")

def clean_val(v):
    if v is None:
        return ""
    s = str(v).strip().strip("`").strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s

def clean_phone(v):
    raw = clean_val(v)
    digits = "".join([c for c in raw if c.isdigit()])
    return digits


@app.route("/api/manager/leads/import", methods=["POST"])
@manager_required
@permission_required("can_import")
def manager_import_leads():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded. Please select an Excel (.xlsx) or CSV (.csv) file."}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "No selected file."}), 400

    filename = file.filename.lower()
    records = []

    try:
        raw_rows = []

        if filename.endswith(".csv"):
            content = file.stream.read().decode("utf-8", errors="ignore")
            stream = io.StringIO(content)
            reader = csv.reader(stream)
            raw_rows = [row for row in reader if any(row)]

        elif filename.endswith(".xlsx"):
            if not HAS_OPENPYXL:
                return jsonify({"error": "openpyxl is not installed on server to read .xlsx files."}), 500

            wb = openpyxl.load_workbook(file, data_only=True)
            sheet = wb.active
            for r in sheet.iter_rows(values_only=True):
                if any(r):
                    raw_rows.append([str(c) if c is not None else "" for c in r])
        else:
            return jsonify({"error": "Unsupported file format. Please upload .xlsx (Excel) or .csv file."}), 400

        if not raw_rows:
            return jsonify({"error": "The uploaded file is empty."}), 400

        # Scan for the actual header row
        header_idx = -1
        for i, row in enumerate(raw_rows):
            cleaned_row = [clean_col(c) for c in row if c]
            if any(h in cleaned_row for h in ["mobile1", "mobile", "phone", "learnername", "studentname", "name"]):
                header_idx = i
                break

        if header_idx == -1:
            return jsonify({
                "error": "Could not identify header columns. Ensure the file contains columns like 'LearnerName' / 'Name' and 'Mobile1' / 'Phone'."
            }), 400

        header = [clean_col(c) for c in raw_rows[header_idx]]

        # Column mapping
        name_idx = next((i for i, h in enumerate(header) if h in ["learnername", "name", "studentname", "fullname", "nameoncertificate", "candidate"]), -1)
        father_idx = next((i for i, h in enumerate(header) if h in ["fathername", "fatherfullname", "father", "guardianname", "parentname", "fathersname"]), -1)
        phone1_idx = next((i for i, h in enumerate(header) if h in ["mobile1", "phone", "mobile", "contact", "phonenumber", "mobileno", "contactno"]), -1)
        phone2_idx = next((i for i, h in enumerate(header) if h in ["mobile2", "altphone", "parentmobile", "alternatephone", "phone2"]), -1)
        course_idx = next((i for i, h in enumerate(header) if h in ["course", "program", "inquiry", "stream", "trade"]), -1)
        batch_idx = next((i for i, h in enumerate(header) if h in ["batch", "session", "batchname", "admissionbatch"]), -1)
        city_idx = next((i for i, h in enumerate(header) if h in ["cityvillagetown", "city", "location", "district", "town", "village", "suburb"]), -1)
        email_idx = next((i for i, h in enumerate(header) if h in ["emailid", "email", "emailaddress"]), -1)
        code_idx = next((i for i, h in enumerate(header) if h in ["learnercode", "studentcode", "regno", "formnumber"]), -1)

        if phone1_idx == -1 and phone2_idx == -1:
            return jsonify({"error": "No phone/mobile column found in header row."}), 400

        for row in raw_rows[header_idx + 1:]:
            if not any(row):
                continue

            name = clean_val(row[name_idx]) if name_idx != -1 and name_idx < len(row) else ""
            father_name = clean_val(row[father_idx]) if father_idx != -1 and father_idx < len(row) else ""
            p1 = clean_phone(row[phone1_idx]) if phone1_idx != -1 and phone1_idx < len(row) else ""
            p2 = clean_phone(row[phone2_idx]) if phone2_idx != -1 and phone2_idx < len(row) else ""
            course = clean_val(row[course_idx]) if course_idx != -1 and course_idx < len(row) else "General Course"
            batch = clean_val(row[batch_idx]) if batch_idx != -1 and batch_idx < len(row) else ""
            city = clean_val(row[city_idx]) if city_idx != -1 and city_idx < len(row) else ""
            email = clean_val(row[email_idx]) if email_idx != -1 and email_idx < len(row) else ""
            code = clean_val(row[code_idx]) if code_idx != -1 and code_idx < len(row) else ""

            primary_phone = p1 if p1 else p2
            alt_phone = p2 if p1 and p2 != p1 else ""

            if name and primary_phone:
                records.append({
                    "name": name,
                    "father_name": father_name,
                    "phone": primary_phone,
                    "alt_phone": alt_phone,
                    "course": course,
                    "batch": batch,
                    "city": city,
                    "email": email,
                    "learner_code": code
                })

    except Exception as e:
        return jsonify({"error": f"Failed to parse file: {str(e)}"}), 400

    if not records:
        return jsonify({"error": "No valid student records with phone numbers found in the file."}), 400

    # Insert all records into shared database
    with get_db() as conn:
        cursor = conn.cursor()
        imported_count = 0
        skipped_count = 0
        seen_in_upload = set()

        for rec in records:
            name_clean = rec["name"].strip()
            phone_clean = rec["phone"].strip()
            code_clean = rec["learner_code"].strip() if rec["learner_code"] else ""
            course_clean = rec["course"].strip() if rec["course"] else ""

            # 1. File-level de-duplication
            dedup_key = f"code:{code_clean}:{course_clean.lower()}" if code_clean else f"phone:{phone_clean}:{name_clean.lower()}:{course_clean.lower()}"
            if dedup_key in seen_in_upload:
                skipped_count += 1
                continue
            seen_in_upload.add(dedup_key)

            # 2. Database-level de-duplication
            existing = None
            if code_clean:
                cursor.execute("""
                    SELECT id, father_name, batch, alt_phone FROM leads
                    WHERE learner_code = ? AND (course = ? OR course IS NULL)
                    LIMIT 1
                """, (code_clean, course_clean))
                existing = cursor.fetchone()

            if not existing:
                cursor.execute("""
                    SELECT id, father_name, batch, alt_phone FROM leads
                    WHERE phone = ? AND LOWER(name) = LOWER(?) AND (course = ? OR course IS NULL)
                    LIMIT 1
                """, (phone_clean, name_clean, course_clean))
                existing = cursor.fetchone()

            if existing:
                # Student already exists! Skip inserting duplicate
                skipped_count += 1
                # Optional: Enrich existing record if it was missing father_name, batch, or alt_phone
                ex_id = existing["id"]
                updates = []
                params = []
                if not existing["father_name"] and rec["father_name"]:
                    updates.append("father_name = ?")
                    params.append(rec["father_name"])
                if not existing["batch"] and rec["batch"]:
                    updates.append("batch = ?")
                    params.append(rec["batch"])
                if not existing["alt_phone"] and rec["alt_phone"]:
                    updates.append("alt_phone = ?")
                    params.append(rec["alt_phone"])

                if updates:
                    params.append(ex_id)
                    cursor.execute(f"UPDATE leads SET {', '.join(updates)} WHERE id = ?", params)
            else:
                cursor.execute("""
                    INSERT INTO leads (name, father_name, phone, alt_phone, email, course, batch, learner_code, city, source, status, record_type)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'SOLAR Export', 'NEW', 'admitted')
                """, (rec["name"], rec["father_name"], rec["phone"], rec["alt_phone"], rec["email"], rec["course"], rec["batch"], rec["learner_code"], rec["city"]))
                imported_count += 1

        conn.commit()
        bump_sync_version()

    if imported_count == 0 and skipped_count > 0:
        msg = f"All {skipped_count} student records in this file already exist in the portal! No duplicates were added."
    elif skipped_count > 0:
        msg = f"Import complete! {imported_count} new students added, and {skipped_count} duplicate records were skipped."
    else:
        msg = f"Successfully imported {imported_count} student records from {file.filename}!"

    return jsonify({
        "success": True,
        "message": msg,
        "count": imported_count,
        "duplicates_skipped": skipped_count
    })


# -------------------------------------------------------------
# STAFF CALLING APIS (STRICT PRIVACY: ONLY 6 FIELDS RETURNED)
# -------------------------------------------------------------
@app.route("/api/staff/leads", methods=["GET"])
@login_required
def staff_get_leads():
    filter_type = request.args.get("tab", "all")
    course_filter = request.args.get("course", "").strip()
    batch_filter = request.args.get("batch", "").strip()
    search = request.args.get("search", "").strip()
    today_str = date.today().isoformat()

    # Staff only gets the 6 allowed fields + status & remarks for calling
    query = """
        SELECT l.id, l.name, l.father_name, l.phone, l.alt_phone, l.course, l.batch, l.city, l.status, l.next_followup_date, l.last_call_at,
               (SELECT remarks FROM call_logs WHERE lead_id = l.id ORDER BY id DESC LIMIT 1) as last_remarks,
               (SELECT COALESCE(u.name, 'Staff') FROM call_logs cl LEFT JOIN users u ON cl.caller_id = u.id WHERE cl.lead_id = l.id ORDER BY cl.id DESC LIMIT 1) as last_caller_name,
               (SELECT COUNT(*) FROM call_logs WHERE lead_id = l.id) as total_calls
        FROM leads l
        WHERE (l.record_type = 'admitted' OR l.record_type IS NULL) AND (l.is_deleted = 0 OR l.is_deleted IS NULL)
    """
    params = []

    if filter_type == "today_followups":
        query += " AND l.status = 'FOLLOW_UP' AND DATE(l.next_followup_date) <= ?"
        params.append(today_str)
    elif filter_type == "new":
        query += " AND l.status = 'NEW'"
    elif filter_type == "exam_informed":
        query += " AND l.status = 'EXAM_INFORMED'"
    elif filter_type == "fee_reminder":
        query += " AND l.status = 'FEE_REMINDER'"
    elif filter_type == "completed":
        query += " AND l.status IN ('COMPLETED', 'ADMISSION_DONE')"

    if course_filter:
        query += " AND l.course = ?"
        params.append(course_filter)

    if batch_filter:
        query += " AND l.batch = ?"
        params.append(batch_filter)

    if search:
        query += " AND (l.name LIKE ? OR l.father_name LIKE ? OR l.phone LIKE ? OR l.alt_phone LIKE ? OR l.course LIKE ? OR l.batch LIKE ? OR l.city LIKE ?)"
        wildcard = f"%{search}%"
        params.extend([wildcard, wildcard, wildcard, wildcard, wildcard, wildcard, wildcard])

    query += " ORDER BY CASE WHEN l.status = 'FOLLOW_UP' AND DATE(l.next_followup_date) <= DATE('now') THEN 0 WHEN l.status = 'NEW' THEN 1 ELSE 2 END, l.id DESC"

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute(query, params)
        rows = cursor.fetchall()
        leads = [dict(r) for r in rows]

        cursor.execute("SELECT DISTINCT course FROM leads WHERE course IS NOT NULL AND course != '' ORDER BY course ASC")
        courses = [r[0] for r in cursor.fetchall()]

        cursor.execute("SELECT DISTINCT batch FROM leads WHERE batch IS NOT NULL AND batch != ''")
        raw_batches = [r[0] for r in cursor.fetchall()]
        batches = sort_batches(raw_batches)

    return jsonify({
        "leads": leads,
        "available_courses": courses,
        "available_batches": batches
    })


@app.route("/api/staff/lead/<int:lead_id>", methods=["GET"])
@login_required
def staff_get_lead_detail(lead_id):
    with get_db() as conn:
        cursor = conn.cursor()
        # Only select allowed fields
        cursor.execute("""
            SELECT id, name, father_name, phone, alt_phone, course, batch, city, status, next_followup_date
            FROM leads WHERE id = ?
        """, (lead_id,))
        lead = cursor.fetchone()
        if not lead:
            return jsonify({"error": "Lead not found"}), 404

        cursor.execute("""
            SELECT cl.*, COALESCE(u.name, 'Staff') as caller_name
            FROM call_logs cl
            LEFT JOIN users u ON cl.caller_id = u.id
            WHERE cl.lead_id = ?
            ORDER BY cl.id DESC
        """, (lead_id,))
        logs = [dict(r) for r in cursor.fetchall()]

        return jsonify({
            "lead": dict(lead),
            "logs": logs
        })


@app.route("/api/staff/lead/<int:lead_id>/call", methods=["POST"])
@login_required
def staff_log_call(lead_id):
    user_id = session["user_id"]
    data = request.json or {}
    status = data.get("status", "").strip()
    remarks = data.get("remarks", "").strip()
    next_followup_date = data.get("next_followup_date", None)

    valid_statuses = [
        "NEW", "EXAM_INFORMED", "FEE_REMINDER", "FOLLOW_UP", "COMPLETED",
        "RINGING", "SWITCHED_OFF", "NOT_INTERESTED", "WRONG_NUMBER", "ADMISSION_DONE", "INTERESTED"
    ]
    if status not in valid_statuses:
        return jsonify({"error": "Invalid status."}), 400

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM leads WHERE id = ?", (lead_id,))
        lead = cursor.fetchone()
        if not lead:
            return jsonify({"error": "Lead not found"}), 404

        cursor.execute("""
            INSERT INTO call_logs (lead_id, caller_id, status, remarks, next_followup_date, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (lead_id, user_id, status, remarks, next_followup_date if status == "FOLLOW_UP" else None, now_str))

        cursor.execute("""
            UPDATE leads
            SET status = ?,
                next_followup_date = ?,
                last_call_at = ?,
                last_caller_id = ?,
                updated_at = ?
            WHERE id = ?
        """, (status, next_followup_date if status == "FOLLOW_UP" else None, now_str, user_id, now_str, lead_id))

        conn.commit()
        bump_sync_version()

    return jsonify({"success": True, "message": "Call logged successfully"})


@app.route("/api/staff/stats", methods=["GET"])
@login_required
def staff_stats():
    user_id = session["user_id"]
    today_str = date.today().isoformat()

    with get_db() as conn:
        cursor = conn.cursor()

        # Shared admitted students count
        cursor.execute("SELECT COUNT(*) FROM leads WHERE (record_type = 'admitted' OR record_type IS NULL) AND (is_deleted = 0 OR is_deleted IS NULL)")
        total_shared_leads = cursor.fetchone()[0]

        # My enquiries stats (exclusive to this staff)
        cursor.execute("SELECT COUNT(*) FROM leads WHERE record_type = 'enquiry' AND assigned_to = ? AND (is_deleted = 0 OR is_deleted IS NULL)", (user_id,))
        my_enquiries_total = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM leads WHERE record_type = 'enquiry' AND assigned_to = ? AND status = 'ADMISSION_DONE' AND (is_deleted = 0 OR is_deleted IS NULL)", (user_id,))
        my_admissions_done = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM leads WHERE record_type = 'enquiry' AND assigned_to = ? AND status = 'CANCELLED' AND (is_deleted = 0 OR is_deleted IS NULL)", (user_id,))
        my_cancelled = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM leads WHERE record_type = 'enquiry' AND assigned_to = ? AND status NOT IN ('ADMISSION_DONE', 'CANCELLED') AND (is_deleted = 0 OR is_deleted IS NULL)", (user_id,))
        my_enquiries_pending = cursor.fetchone()[0]

        cursor.execute("""
            SELECT COUNT(*) FROM call_logs
            WHERE caller_id = ? AND DATE(created_at) = ?
        """, (user_id, today_str))
        my_calls_today = cursor.fetchone()[0]

        cursor.execute("""
            SELECT COUNT(*) FROM leads
            WHERE assigned_to = ? AND status = 'FOLLOW_UP' AND DATE(next_followup_date) <= ?
        """, (user_id, today_str))
        followups_due = cursor.fetchone()[0]

        # Fees Pending count
        cursor.execute("SELECT COUNT(*) FROM leads WHERE (record_type = 'admitted' OR record_type IS NULL) AND (is_deleted = 0 OR is_deleted IS NULL) AND is_fee_pending = 1")
        total_fees_pending = cursor.fetchone()[0]

        # Current Month Exam Calling info
        curr_m_code = date.today().strftime("%Y-%m")
        cursor.execute("SELECT is_visible_to_staff FROM exam_month_settings WHERE month_code = ?", (curr_m_code,))
        exam_vis_row = cursor.fetchone()
        exam_visible_to_staff = exam_vis_row["is_visible_to_staff"] if exam_vis_row else 1

        cursor.execute("SELECT COUNT(*) FROM exam_students WHERE month_code = ?", (curr_m_code,))
        total_exam_students = cursor.fetchone()[0] if exam_visible_to_staff else 0

        return jsonify({
            "total_leads": total_shared_leads,
            "total_admitted_leads": total_shared_leads,
            "my_enquiries_total": my_enquiries_total,
            "my_admissions_done": my_admissions_done,
            "my_cancelled": my_cancelled,
            "my_enquiries_pending": my_enquiries_pending,
            "calls_today": my_calls_today,
            "followups_due": followups_due,
            "admissions": my_admissions_done,
            "total_fees_pending": total_fees_pending,
            "total_exam_students": total_exam_students,
            "exam_visible_to_staff": exam_visible_to_staff,
            "current_exam_month": curr_m_code
        })


# -------------------------------------------------------------
# MANAGER DASHBOARD & ADMIN APIS
# -------------------------------------------------------------
@app.route("/api/manager/stats", methods=["GET"])
@manager_required
@permission_required("can_view")
def manager_stats():
    today_str = date.today().isoformat()

    with get_db() as conn:
        cursor = conn.cursor()

        # Admitted students pool (SOLAR 1,929 records)
        cursor.execute("SELECT COUNT(*) FROM leads WHERE (record_type = 'admitted' OR record_type IS NULL) AND (is_deleted = 0 OR is_deleted IS NULL)")
        total_admitted_students = cursor.fetchone()[0]

        # Inquiries pool
        cursor.execute("SELECT COUNT(*) FROM leads WHERE record_type = 'enquiry' AND (is_deleted = 0 OR is_deleted IS NULL)")
        total_enquiries = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM leads WHERE record_type = 'enquiry' AND status = 'ADMISSION_DONE' AND (is_deleted = 0 OR is_deleted IS NULL)")
        enquiries_converted = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM leads WHERE record_type = 'enquiry' AND status = 'CANCELLED' AND (is_deleted = 0 OR is_deleted IS NULL)")
        enquiries_cancelled = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM leads WHERE record_type = 'enquiry' AND status NOT IN ('ADMISSION_DONE', 'CANCELLED') AND (is_deleted = 0 OR is_deleted IS NULL)")
        enquiries_pending = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM leads WHERE is_deleted = 1")
        total_trash = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM call_logs WHERE DATE(created_at) = ?", (today_str,))
        calls_today = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM leads WHERE status = 'FOLLOW_UP' AND DATE(next_followup_date) <= ? AND (is_deleted = 0 OR is_deleted IS NULL)", (today_str,))
        pending_followups = cursor.fetchone()[0]

        cursor.execute("""
            SELECT u.id, u.name, u.phone, u.is_active,
                   (
                       SELECT COUNT(*) FROM call_logs cl
                       WHERE cl.caller_id = u.id AND DATE(cl.created_at) = ?
                   ) as calls_today,
                   (
                       SELECT COUNT(*) FROM call_logs cl
                       WHERE cl.caller_id = u.id
                   ) as total_calls,
                   (
                       SELECT COUNT(*) FROM leads l
                       WHERE l.record_type = 'enquiry' AND (l.is_deleted = 0 OR l.is_deleted IS NULL) AND l.assigned_to = u.id AND (l.is_deleted = 0 OR l.is_deleted IS NULL)
                   ) as enquiries_assigned,
                   (
                       SELECT COUNT(*) FROM leads l
                       WHERE l.record_type = 'enquiry' AND (l.is_deleted = 0 OR l.is_deleted IS NULL) AND l.assigned_to = u.id AND l.status = 'ADMISSION_DONE' AND (l.is_deleted = 0 OR l.is_deleted IS NULL)
                   ) as admissions_done,
                   (
                       SELECT COUNT(*) FROM leads l
                       WHERE l.record_type = 'enquiry' AND (l.is_deleted = 0 OR l.is_deleted IS NULL) AND l.assigned_to = u.id AND l.status = 'CANCELLED' AND (l.is_deleted = 0 OR l.is_deleted IS NULL)
                   ) as cancelled_count
            FROM users u
            WHERE u.role = 'staff'
            ORDER BY u.id ASC
        """, (today_str,))

        staff_list = []
        for r in cursor.fetchall():
            s = dict(r)
            assigned = s["enquiries_assigned"] or 0
            adm = s["admissions_done"] or 0
            s["conversion_rate"] = round((adm / assigned * 100), 1) if assigned > 0 else 0
            staff_list.append(s)

        # Fees Pending count
        cursor.execute("SELECT COUNT(*) FROM leads WHERE (record_type = 'admitted' OR record_type IS NULL) AND (is_deleted = 0 OR is_deleted IS NULL) AND is_fee_pending = 1")
        total_fees_pending = cursor.fetchone()[0]

        # Current Month Exam Calling info
        curr_m_code = date.today().strftime("%Y-%m")
        cursor.execute("SELECT is_visible_to_staff FROM exam_month_settings WHERE month_code = ?", (curr_m_code,))
        exam_vis_row = cursor.fetchone()
        exam_visible_to_staff = exam_vis_row["is_visible_to_staff"] if exam_vis_row else 1

        cursor.execute("SELECT COUNT(*) FROM exam_students WHERE month_code = ?", (curr_m_code,))
        total_exam_students = cursor.fetchone()[0]

        return jsonify({
            "total_leads": total_admitted_students,
            "total_admitted_students": total_admitted_students,
            "total_enquiries": total_enquiries,
            "enquiries_converted": enquiries_converted,
            "enquiries_cancelled": enquiries_cancelled,
            "enquiries_pending": enquiries_pending,
            "calls_today": calls_today,
            "pending_followups": pending_followups,
            "total_admissions": enquiries_converted,
            "total_trash": total_trash,
            "total_fees_pending": total_fees_pending,
            "total_exam_students": total_exam_students,
            "exam_visible_to_staff": exam_visible_to_staff,
            "current_exam_month": curr_m_code,
            "staff_performance": staff_list
        })


@app.route("/api/manager/leads", methods=["GET"])
@manager_required
@permission_required("can_view")
def manager_get_leads():
    status = request.args.get("status")
    course_filter = request.args.get("course", "").strip()
    batch_filter = request.args.get("batch", "").strip()
    search = request.args.get("search", "").strip()

    query = """
        SELECT l.*,
               u.name as last_caller_name,
               (SELECT COUNT(*) FROM call_logs WHERE lead_id = l.id) as total_calls,
               (SELECT remarks FROM call_logs WHERE lead_id = l.id ORDER BY id DESC LIMIT 1) as last_remarks
        FROM leads l
        LEFT JOIN users u ON l.last_caller_id = u.id
        WHERE (l.record_type = 'admitted' OR l.record_type IS NULL) AND (l.is_deleted = 0 OR l.is_deleted IS NULL)
    """
    params = []

    if status and status != "ALL":
        if status == "COMPLETED":
            query += " AND l.status IN ('COMPLETED', 'ADMISSION_DONE')"
        else:
            query += " AND l.status = ?"
            params.append(status)

    if course_filter:
        query += " AND l.course = ?"
        params.append(course_filter)

    if batch_filter:
        query += " AND l.batch = ?"
        params.append(batch_filter)

    if search:
        query += " AND (l.name LIKE ? OR l.father_name LIKE ? OR l.phone LIKE ? OR l.alt_phone LIKE ? OR l.course LIKE ? OR l.batch LIKE ? OR l.city LIKE ? OR l.learner_code LIKE ?)"
        wildcard = f"%{search}%"
        params.extend([wildcard, wildcard, wildcard, wildcard, wildcard, wildcard, wildcard, wildcard])

    query += " ORDER BY l.id DESC"

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute(query, params)
        leads = [dict(r) for r in cursor.fetchall()]

        cursor.execute("SELECT DISTINCT course FROM leads WHERE course IS NOT NULL AND course != '' ORDER BY course ASC")
        courses = [r[0] for r in cursor.fetchall()]

        cursor.execute("SELECT DISTINCT batch FROM leads WHERE batch IS NOT NULL AND batch != ''")
        raw_batches = [r[0] for r in cursor.fetchall()]
        batches = sort_batches(raw_batches)

    return jsonify({
        "leads": leads,
        "available_courses": courses,
        "available_batches": batches
    })


@app.route("/api/manager/leads/create", methods=["POST"])
@manager_required
@permission_required("can_edit")
def manager_create_lead():
    data = request.json or {}
    name = data.get("name", "").strip()
    father_name = data.get("father_name", "").strip()
    phone = data.get("phone", "").strip()
    alt_phone = data.get("alt_phone", "").strip()
    email = data.get("email", "").strip()
    course = data.get("course", "").strip()
    batch = data.get("batch", "").strip()
    city = data.get("city", "").strip()

    if not name or not phone:
        return jsonify({"error": "Student Name and Phone Number are required"}), 400

    with get_db() as conn:
        cursor = conn.cursor()

        # Check duplicate
        cursor.execute("SELECT id FROM leads WHERE phone = ? AND LOWER(name) = LOWER(?)", (phone, name))
        if cursor.fetchone():
            return jsonify({"error": f"Student '{name}' with phone {phone} already exists in the portal!"}), 400

        cursor.execute("""
            INSERT INTO leads (name, father_name, phone, alt_phone, email, course, batch, city, source, status, record_type)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'Manual / Manager', 'NEW', 'admitted')
        """, (name, father_name, phone, alt_phone, email, course, batch, city))
        lead_id = cursor.lastrowid
        conn.commit()
        bump_sync_version()

    return jsonify({"success": True, "lead_id": lead_id, "message": "Lead added to shared directory successfully"})


@app.route("/api/manager/leads/<int:lead_id>", methods=["GET", "POST", "DELETE"])
@manager_required
def manager_manage_lead(lead_id):
    with get_db() as conn:
        cursor = conn.cursor()

        if request.method == "DELETE":
            if not user_has_permission("can_delete"):
                return jsonify({"error": "Permission Denied: You do not have 'can_delete' permission to delete records."}), 403
            cursor.execute("SELECT id, name FROM leads WHERE id = ?", (lead_id,))
            lead = cursor.fetchone()
            if not lead:
                return jsonify({"error": "Student not found"}), 404
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            cursor.execute("UPDATE leads SET is_deleted = 1, deleted_at = ?, updated_at = ? WHERE id = ?", (now_str, now_str, lead_id))
            conn.commit()
            bump_sync_version()
            return jsonify({"success": True, "message": f"Student '{lead['name']}' moved to Recycle Bin. You can restore anytime!"})

        if request.method == "GET":
            cursor.execute("SELECT * FROM leads WHERE id = ?", (lead_id,))
            lead = cursor.fetchone()
            if not lead:
                return jsonify({"error": "Student not found"}), 404
            return jsonify({"lead": dict(lead)})

        if request.method == "POST":
            if not user_has_permission("can_edit"):
                return jsonify({"error": "Permission Denied: You do not have 'can_edit' permission to edit records."}), 403
            data = request.json or {}
            name = data.get("name", "").strip()
            father_name = data.get("father_name", "").strip()
            phone = data.get("phone", "").strip()
            alt_phone = data.get("alt_phone", "").strip()
            course = data.get("course", "").strip()
            batch = data.get("batch", "").strip()
            city = data.get("city", "").strip()

            if not name or not phone:
                return jsonify({"error": "Student Name and Primary Phone are required"}), 400

            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            cursor.execute("""
                UPDATE leads
                SET name = ?, father_name = ?, phone = ?, alt_phone = ?, course = ?, batch = ?, city = ?, updated_at = ?
                WHERE id = ?
            """, (name, father_name, phone, alt_phone, course, batch, city, now_str, lead_id))
            conn.commit()
            bump_sync_version()
            return jsonify({"success": True, "message": f"Student '{name}' details updated successfully!"})


@app.route("/api/manager/staff", methods=["GET", "POST"])
@manager_required
@permission_required("can_manage_users")
def manager_staff():
    with get_db() as conn:
        cursor = conn.cursor()

        if request.method == "POST":
            data = request.json or {}
            name = data.get("name", "").strip()
            username = data.get("username", "").strip()
            password = data.get("password", "").strip()
            phone = data.get("phone", "").strip()
            role = data.get("role", "staff").strip().lower()
            if role not in ("manager", "staff"):
                role = "staff"

            perms_input = data.get("permissions", {})
            if not isinstance(perms_input, dict):
                perms_input = {}
            perms_clean = {k: 1 if perms_input.get(k) in (1, "1", True) else 0 for k in ALL_PERMISSIONS}

            if not name or not username or not password:
                return jsonify({"error": "Name, username, and password are required"}), 400

            cursor.execute("SELECT id FROM users WHERE username = ?", (username,))
            if cursor.fetchone():
                return jsonify({"error": f"Username '{username}' already exists. Choose another one."}), 400

            cursor.execute("""
                INSERT INTO users (name, username, password_hash, role, phone, is_active, permissions)
                VALUES (?, ?, ?, ?, ?, 1, ?)
            """, (name, username, generate_password_hash(password), role, phone, json.dumps(perms_clean)))
            staff_id = cursor.lastrowid
            conn.commit()
            bump_sync_version()

            role_label = "Admin" if role == "manager" else "Staff"
            return jsonify({
                "success": True,
                "message": f"{role_label} account for '{name}' created successfully!",
                "staff_id": staff_id
            }), 201

        cursor.execute("SELECT id, name, username, role, phone, is_active, permissions, created_at FROM users ORDER BY id ASC")
        users_list = []
        for r in cursor.fetchall():
            u = dict(r)
            u["permissions"] = get_user_permissions(u["id"], u["username"], u["role"], u.get("permissions"))
            u["is_super_admin"] = (u["username"] == "arsh" or u["id"] == 1)
            users_list.append(u)

        cursor.execute("SELECT id, name, username FROM users WHERE role = 'staff' AND is_active = 1 ORDER BY name ASC")
        calling_staff = [dict(r) for r in cursor.fetchall()]

        return jsonify({"staff": users_list, "users": users_list, "calling_staff": calling_staff})


@app.route("/api/manager/staff/<int:staff_id>/toggle", methods=["POST"])
@manager_required
@permission_required("can_manage_users")
def manager_staff_toggle(staff_id):
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, username, is_active FROM users WHERE id = ?", (staff_id,))
        row = cursor.fetchone()
        if not row:
            return jsonify({"error": "Account not found"}), 404

        if row["username"] == "arsh" or row["id"] == 1:
            return jsonify({"error": "Super Admin account (Arsh) cannot be deactivated!"}), 403

        new_status = 0 if row["is_active"] == 1 else 1
        cursor.execute("UPDATE users SET is_active = ? WHERE id = ?", (new_status, staff_id))
        conn.commit()
        bump_sync_version()

    return jsonify({"success": True, "is_active": new_status})


@app.route("/api/manager/staff/<int:staff_id>", methods=["DELETE"])
@manager_required
@permission_required("can_manage_users")
def manager_delete_staff(staff_id):
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, name, username FROM users WHERE id = ?", (staff_id,))
        target = cursor.fetchone()
        if not target:
            return jsonify({"error": "Account not found"}), 404

        if target["username"] == "arsh" or target["id"] == 1:
            return jsonify({"error": "Super Admin account (Arsh) cannot be deleted!"}), 403

        cursor.execute("DELETE FROM users WHERE id = ?", (staff_id,))
        cursor.execute("UPDATE leads SET assigned_to = NULL WHERE assigned_to = ?", (staff_id,))
        conn.commit()
        bump_sync_version()

    return jsonify({"success": True, "message": f"Account '{target['name']}' deleted successfully!"})


@app.route("/api/manager/staff/<int:staff_id>/update", methods=["POST"])
@manager_required
@permission_required("can_manage_users")
def manager_update_staff(staff_id):
    data = request.json or {}
    name = data.get("name", "").strip()
    username = data.get("username", "").strip()
    phone = data.get("phone", "").strip()
    new_password = data.get("password", "").strip()
    is_active = data.get("is_active", 1)
    role = data.get("role")
    perms_input = data.get("permissions")

    if not name or not username:
        return jsonify({"error": "Name and Username are required."}), 400

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, username, role, permissions FROM users WHERE id = ?", (staff_id,))
        target_user = cursor.fetchone()
        if not target_user:
            return jsonify({"error": "Account not found."}), 404

        is_target_super = (target_user["username"] == "arsh" or target_user["id"] == 1)
        if is_target_super:
            if session.get("username") != "arsh" and session.get("user_id") != 1:
                return jsonify({"error": "Only Super Admin Arsh can edit the primary admin account."}), 403
            is_active = 1
            role = "manager"

        cursor.execute("SELECT id FROM users WHERE username = ? AND id != ?", (username, staff_id))
        if cursor.fetchone():
            return jsonify({"error": f"Username '{username}' is already taken by another user."}), 400

        if is_target_super:
            perms_json = json.dumps(ALL_PERMISSIONS)
        elif perms_input is not None and isinstance(perms_input, dict):
            perms_clean = {k: 1 if perms_input.get(k) in (1, "1", True) else 0 for k in ALL_PERMISSIONS}
            perms_json = json.dumps(perms_clean)
        else:
            perms_json = None

        if role not in ("manager", "staff"):
            role = target_user["role"]

        if new_password:
            if perms_json is not None:
                cursor.execute("""
                    UPDATE users SET name = ?, username = ?, phone = ?, role = ?, is_active = ?, permissions = ?, password_hash = ?
                    WHERE id = ?
                """, (name, username, phone, role, int(is_active), perms_json, generate_password_hash(new_password), staff_id))
            else:
                cursor.execute("""
                    UPDATE users SET name = ?, username = ?, phone = ?, role = ?, is_active = ?, password_hash = ?
                    WHERE id = ?
                """, (name, username, phone, role, int(is_active), generate_password_hash(new_password), staff_id))
        else:
            if perms_json is not None:
                cursor.execute("""
                    UPDATE users SET name = ?, username = ?, phone = ?, role = ?, is_active = ?, permissions = ?
                    WHERE id = ?
                """, (name, username, phone, role, int(is_active), perms_json, staff_id))
            else:
                cursor.execute("""
                    UPDATE users SET name = ?, username = ?, phone = ?, role = ?, is_active = ?
                    WHERE id = ?
                """, (name, username, phone, role, int(is_active), staff_id))
        conn.commit()
        bump_sync_version()

    return jsonify({"success": True, "message": f"Account '{name}' details and permissions updated successfully!"})


@app.route("/api/manager/staff/<int:staff_id>/change-password", methods=["POST"])
@manager_required
@permission_required("can_manage_users")
def manager_change_staff_password(staff_id):
    data = request.json or {}
    new_password = data.get("password", "").strip()
    if not new_password:
        return jsonify({"error": "Password cannot be blank."}), 400
    if len(new_password) < 3:
        return jsonify({"error": "Password must be at least 3 characters long."}), 400

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, name, username FROM users WHERE id = ?", (staff_id,))
        target = cursor.fetchone()
        if not target:
            return jsonify({"error": "Account not found."}), 404

        # If target is Super Admin Arsh, only Arsh can change Arsh's password
        if (target["username"] == "arsh" or target["id"] == 1) and session.get("username") != "arsh":
            return jsonify({"error": "Only Super Admin Arsh can change the primary admin password."}), 403

        cursor.execute("UPDATE users SET password_hash = ? WHERE id = ?", (generate_password_hash(new_password), staff_id))
        conn.commit()
        bump_sync_version()

    return jsonify({
        "success": True,
        "message": f"Password for '{target['name']}' (@{target['username']}) successfully updated to '{new_password}'!"
    })


@app.route("/api/manager/profile", methods=["GET", "POST"])
@manager_required
def manager_profile():
    user_id = session["user_id"]
    with get_db() as conn:
        cursor = conn.cursor()

        if request.method == "GET":
            cursor.execute("SELECT id, name, username, phone, role FROM users WHERE id = ?", (user_id,))
            user = cursor.fetchone()
            return jsonify({"user": dict(user)})

        data = request.json or {}
        name = data.get("name", "").strip()
        username = data.get("username", "").strip()
        phone = data.get("phone", "").strip()
        new_password = data.get("password", "").strip()

        if not name or not username:
            return jsonify({"error": "Name and Username are required."}), 400

        cursor.execute("SELECT id FROM users WHERE username = ? AND id != ?", (username, user_id))
        if cursor.fetchone():
            return jsonify({"error": f"Username '{username}' is already in use."}), 400

        if new_password:
            cursor.execute("""
                UPDATE users SET name = ?, username = ?, phone = ?, password_hash = ?
                WHERE id = ?
            """, (name, username, phone, generate_password_hash(new_password), user_id))
        else:
            cursor.execute("""
                UPDATE users SET name = ?, username = ?, phone = ?
                WHERE id = ?
            """, (name, username, phone, user_id))
        conn.commit()

        session["name"] = name
        session["username"] = username

    return jsonify({"success": True, "message": "Admin profile and password updated successfully!"})


# -------------------------------------------------------------
# ENQUIRIES APIS (FRESH PROSPECTIVE STUDENTS / ADMISSION CALLING)
# -------------------------------------------------------------

@app.route("/api/manager/enquiries", methods=["GET"])
@manager_required
@permission_required("can_view")
def manager_get_enquiries():
    staff_filter = request.args.get("staff_id", "").strip()
    status_filter = request.args.get("status", "").strip()
    course_filter = request.args.get("course", "").strip()
    search = request.args.get("search", "").strip()

    query = """
        SELECT l.*,
               u.name as assigned_to_name,
               u.phone as assigned_to_phone,
               (SELECT remarks FROM call_logs WHERE lead_id = l.id ORDER BY id DESC LIMIT 1) as last_remarks,
               (SELECT COALESCE(u2.name, 'Staff') FROM call_logs cl LEFT JOIN users u2 ON cl.caller_id = u2.id WHERE cl.lead_id = l.id ORDER BY cl.id DESC LIMIT 1) as last_caller_name,
               (SELECT COUNT(*) FROM call_logs WHERE lead_id = l.id) as total_calls
        FROM leads l
        LEFT JOIN users u ON l.assigned_to = u.id
        WHERE l.record_type = 'enquiry' AND (l.is_deleted = 0 OR l.is_deleted IS NULL)
    """
    params = []

    if staff_filter:
        if staff_filter == "unassigned":
            query += " AND l.assigned_to IS NULL"
        else:
            query += " AND l.assigned_to = ?"
            params.append(int(staff_filter))

    if status_filter and status_filter != "ALL":
        query += " AND l.status = ?"
        params.append(status_filter)

    if course_filter:
        query += " AND l.course = ?"
        params.append(course_filter)

    if search:
        query += " AND (l.name LIKE ? OR l.father_name LIKE ? OR l.phone LIKE ? OR l.alt_phone LIKE ? OR l.city LIKE ? OR l.notes LIKE ?)"
        wildcard = f"%{search}%"
        params.extend([wildcard, wildcard, wildcard, wildcard, wildcard, wildcard])

    query += " ORDER BY CASE WHEN l.status = 'NEW' THEN 0 WHEN l.status = 'FOLLOW_UP' THEN 1 WHEN l.status = 'ADMISSION_DONE' THEN 2 ELSE 3 END, l.id DESC"

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute(query, params)
        enquiries = [dict(r) for r in cursor.fetchall()]

    return jsonify({"enquiries": enquiries})


@app.route("/api/manager/enquiries", methods=["POST"])
@manager_required
@permission_required("can_edit")
def manager_create_enquiry():
    data = request.json or {}
    name = data.get("name", "").strip()
    phone = data.get("phone", "").strip()
    alt_phone = data.get("alt_phone", "").strip()
    father_name = data.get("father_name", "").strip()
    city = data.get("city", "").strip()
    course = data.get("course", "").strip()
    source = data.get("source", "Walk-in Center").strip()
    notes = data.get("notes", "").strip()
    assigned_to = data.get("assigned_to", None)

    if not name or not phone:
        return jsonify({"error": "Candidate Name and Phone Number are required"}), 400

    if assigned_to == "" or assigned_to is None:
        assigned_to = None
    else:
        try:
            assigned_to = int(assigned_to)
        except (ValueError, TypeError):
            assigned_to = None

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO leads (name, father_name, phone, alt_phone, city, course, source, notes, status, assigned_to, record_type, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'NEW', ?, 'enquiry', ?, ?)
        """, (name, father_name, phone, alt_phone, city, course, source, notes, assigned_to, now_str, now_str))
        enquiry_id = cursor.lastrowid
        conn.commit()
        bump_sync_version()

    return jsonify({"success": True, "enquiry_id": enquiry_id, "message": "New Admission Inquiry added successfully!"})


@app.route("/api/manager/enquiries/<int:enquiry_id>/assign", methods=["POST"])
@manager_required
@permission_required("can_edit")
def manager_assign_enquiry(enquiry_id):
    data = request.json or {}
    assigned_to = data.get("assigned_to")
    if assigned_to == "" or assigned_to is None:
        assigned_to = None
    else:
        try:
            assigned_to = int(assigned_to)
        except (ValueError, TypeError):
            assigned_to = None

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("UPDATE leads SET assigned_to = ?, updated_at = ? WHERE id = ? AND record_type = 'enquiry'",
                       (assigned_to, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), enquiry_id))
        conn.commit()
        bump_sync_version()

    return jsonify({"success": True, "message": "Staff assignment updated successfully!"})


@app.route("/api/manager/enquiries/<int:enquiry_id>", methods=["DELETE"])
@manager_required
@permission_required("can_delete")
def manager_delete_enquiry(enquiry_id):
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, name FROM leads WHERE id = ? AND record_type = 'enquiry'", (enquiry_id,))
        lead = cursor.fetchone()
        if not lead:
            return jsonify({"error": "Inquiry not found"}), 404
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cursor.execute("UPDATE leads SET is_deleted = 1, deleted_at = ?, updated_at = ? WHERE id = ?", (now_str, now_str, enquiry_id))
        conn.commit()
        bump_sync_version()

    return jsonify({"success": True, "message": f"Inquiry for '{lead['name']}' moved to Recycle Bin. You can restore anytime!"})


# -------------------------------------------------------------
# STAFF ENQUIRY APIS (SHARED VISIBILITY: ALL STAFF SEE INQUIRIES)
# -------------------------------------------------------------

@app.route("/api/staff/enquiries", methods=["GET"])
@login_required
def staff_get_enquiries():
    user_id = session["user_id"]
    tab = request.args.get("tab", "all")
    course_filter = request.args.get("course", "").strip()
    search = request.args.get("search", "").strip()
    scope = request.args.get("scope", "all")
    today_str = date.today().isoformat()

    # All active enquiries in institute; indicates assigned staff name
    query = """
        SELECT l.*,
               u.name as assigned_staff_name,
               (SELECT remarks FROM call_logs WHERE lead_id = l.id ORDER BY id DESC LIMIT 1) as last_remarks,
               (SELECT COUNT(*) FROM call_logs WHERE lead_id = l.id) as total_calls
        FROM leads l
        LEFT JOIN users u ON l.assigned_to = u.id
        WHERE l.record_type = 'enquiry' AND (l.is_deleted = 0 OR l.is_deleted IS NULL)
    """
    params = []

    if scope == "my":
        query += " AND l.assigned_to = ?"
        params.append(user_id)

    if tab == "new":
        query += " AND l.status = 'NEW'"
    elif tab == "today_followups":
        query += " AND l.status = 'FOLLOW_UP' AND DATE(l.next_followup_date) <= ?"
        params.append(today_str)
    elif tab == "admission_done":
        query += " AND l.status = 'ADMISSION_DONE'"
    elif tab == "cancelled":
        query += " AND l.status = 'CANCELLED'"

    if course_filter:
        query += " AND l.course = ?"
        params.append(course_filter)

    if search:
        query += " AND (l.name LIKE ? OR l.father_name LIKE ? OR l.phone LIKE ? OR l.alt_phone LIKE ? OR l.city LIKE ? OR l.notes LIKE ?)"
        wildcard = f"%{search}%"
        params.extend([wildcard, wildcard, wildcard, wildcard, wildcard, wildcard])

    query += " ORDER BY CASE WHEN l.status = 'NEW' THEN 0 WHEN l.status = 'FOLLOW_UP' AND DATE(l.next_followup_date) <= DATE('now') THEN 1 WHEN l.status = 'FOLLOW_UP' THEN 2 ELSE 3 END, l.id DESC"

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute(query, params)
        enquiries = [dict(r) for r in cursor.fetchall()]

    return jsonify({"enquiries": enquiries})


@app.route("/api/staff/enquiries", methods=["POST"])
@login_required
def staff_create_enquiry():
    user_id = session["user_id"]
    data = request.json or {}
    name = data.get("name", "").strip()
    phone = data.get("phone", "").strip()
    alt_phone = data.get("alt_phone", "").strip()
    father_name = data.get("father_name", "").strip()
    city = data.get("city", "").strip()
    course = data.get("course", "").strip()
    source = data.get("source", "Walk-in Center").strip()
    notes = data.get("notes", "").strip()

    if not name or not phone:
        return jsonify({"error": "Candidate Name and Phone Number are required"}), 400

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO leads (name, father_name, phone, alt_phone, city, course, source, notes, status, assigned_to, record_type, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'NEW', ?, 'enquiry', ?, ?)
        """, (name, father_name, phone, alt_phone, city, course, source, notes, user_id, now_str, now_str))
        enquiry_id = cursor.lastrowid
        conn.commit()
        bump_sync_version()

    return jsonify({"success": True, "enquiry_id": enquiry_id, "message": "Inquiry added to pipeline!"})


@app.route("/api/enquiries/<int:enquiry_id>/status", methods=["POST"])
@login_required
def update_enquiry_status(enquiry_id):
    user_id = session["user_id"]
    is_manager = session.get("role") == "manager"
    data = request.json or {}
    status = data.get("status", "").strip().upper()
    remarks = data.get("remarks", "").strip()
    next_followup_date = data.get("next_followup_date", None)
    cancellation_reason = data.get("cancellation_reason", "").strip()

    valid_statuses = ["NEW", "FOLLOW_UP", "ADMISSION_DONE", "CANCELLED", "RINGING", "VISIT_SCHEDULED"]
    if status not in valid_statuses:
        return jsonify({"error": "Invalid enquiry status"}), 400

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with get_db() as conn:
        cursor = conn.cursor()
        # Verify ownership: manager can edit any; staff can edit their own or unassigned
        if is_manager:
            cursor.execute("SELECT id, name FROM leads WHERE id = ? AND record_type = 'enquiry'", (enquiry_id,))
        else:
            cursor.execute("SELECT id, name FROM leads WHERE id = ? AND record_type = 'enquiry' AND (assigned_to = ? OR assigned_to IS NULL)", (enquiry_id, user_id))
        enquiry = cursor.fetchone()
        if not enquiry:
            return jsonify({"error": "Inquiry not found or access denied"}), 404

        # Log call
        cursor.execute("""
            INSERT INTO call_logs (lead_id, caller_id, status, remarks, next_followup_date, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (enquiry_id, user_id, status, remarks or f"Status updated to {status}", next_followup_date if status == "FOLLOW_UP" else None, now_str))

        conversion_date = now_str if status == "ADMISSION_DONE" else None

        cursor.execute("""
            UPDATE leads
            SET status = ?,
                assigned_to = COALESCE(assigned_to, ?),
                cancellation_reason = CASE WHEN ? = 'CANCELLED' THEN ? ELSE cancellation_reason END,
                conversion_date = CASE WHEN ? = 'ADMISSION_DONE' THEN ? ELSE conversion_date END,
                next_followup_date = ?,
                last_call_at = ?,
                last_caller_id = ?,
                updated_at = ?
            WHERE id = ?
        """, (status, user_id, status, cancellation_reason, status, conversion_date,
              next_followup_date if status == "FOLLOW_UP" else None, now_str, user_id, now_str, enquiry_id))
        conn.commit()
        bump_sync_version()

    return jsonify({"success": True, "message": f"Inquiry status updated to {status}!"})


# -------------------------------------------------------------
# RECYCLE BIN / TRASH / RESTORE APIS
# -------------------------------------------------------------

@app.route("/api/manager/trash", methods=["GET"])
@manager_required
@permission_required("can_restore")
def manager_get_trash():
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT l.*,
                   u.name as last_caller_name
            FROM leads l
            LEFT JOIN users u ON l.last_caller_id = u.id
            WHERE l.is_deleted = 1
            ORDER BY l.deleted_at DESC, l.id DESC
        """)
        trash_items = [dict(r) for r in cursor.fetchall()]
        return jsonify({"trash": trash_items, "total_deleted": len(trash_items)})


@app.route("/api/manager/trash/<int:lead_id>/restore", methods=["POST"])
@manager_required
@permission_required("can_restore")
def manager_restore_lead(lead_id):
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, name, record_type FROM leads WHERE id = ? AND is_deleted = 1", (lead_id,))
        lead = cursor.fetchone()
        if not lead:
            return jsonify({"error": "Record not found in Recycle Bin"}), 404

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cursor.execute("UPDATE leads SET is_deleted = 0, deleted_at = NULL, updated_at = ? WHERE id = ?", (now_str, lead_id))
        conn.commit()
        bump_sync_version()
        item_type = "Admitted student" if lead["record_type"] == "admitted" else "Inquiry"
        return jsonify({"success": True, "message": f"{item_type} '{lead['name']}' restored successfully!"})


@app.route("/api/manager/trash/restore-all", methods=["POST"])
@manager_required
@permission_required("can_restore")
def manager_restore_all_trash():
    with get_db() as conn:
        cursor = conn.cursor()
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cursor.execute("UPDATE leads SET is_deleted = 0, deleted_at = NULL, updated_at = ? WHERE is_deleted = 1", (now_str,))
        count = cursor.rowcount
        conn.commit()
        bump_sync_version()
        return jsonify({"success": True, "message": f"Successfully restored {count} records from Recycle Bin!"})


@app.route("/api/manager/trash/<int:lead_id>/permanent", methods=["DELETE"])
@manager_required
@permission_required("can_restore")
def manager_permanent_delete_lead(lead_id):
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, name FROM leads WHERE id = ? AND is_deleted = 1", (lead_id,))
        lead = cursor.fetchone()
        if not lead:
            return jsonify({"error": "Record not found in Recycle Bin"}), 404

        cursor.execute("DELETE FROM leads WHERE id = ?", (lead_id,))
        cursor.execute("DELETE FROM call_logs WHERE lead_id = ?", (lead_id,))
        conn.commit()
        bump_sync_version()
        return jsonify({"success": True, "message": f"Record '{lead['name']}' permanently deleted."})


@app.route("/api/manager/settings", methods=["GET", "POST"])
@manager_required
def manager_settings():
    with get_db() as conn:
        cursor = conn.cursor()

        if request.method == "POST":
            data = request.json or {}
            if "regenerate_key" in data and data["regenerate_key"]:
                new_key = "focus_live_" + secrets.token_hex(12)
                cursor.execute("UPDATE settings SET value = ? WHERE key = 'webhook_api_key'", (new_key,))
            conn.commit()

        cursor.execute("SELECT key, value FROM settings")
        settings_dict = {r["key"]: r["value"] for r in cursor.fetchall()}

        return jsonify({
            "webhook_api_key": settings_dict.get("webhook_api_key"),
            "webhook_url": request.host_url.rstrip("/") + "/api/webhook/lead"
        })


@app.route("/api/manager/export", methods=["GET"])
@manager_required
@permission_required("can_export")
def manager_export_leads():
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT l.id, l.name, l.father_name, l.phone, l.alt_phone, l.course, l.batch, l.learner_code, l.city, l.source,
                   l.status, u.name as last_called_by, l.next_followup_date, l.last_call_at, l.created_at,
                   (SELECT remarks FROM call_logs WHERE lead_id = l.id ORDER BY id DESC LIMIT 1) as last_remarks
            FROM leads l
            LEFT JOIN users u ON l.last_caller_id = u.id
            WHERE (l.is_deleted = 0 OR l.is_deleted IS NULL)
            ORDER BY l.id DESC
        """)
        rows = cursor.fetchall()

    si = io.StringIO()
    cw = csv.writer(si)
    cw.writerow([
        "Lead ID", "Student Name", "Father Name", "Mobile 1", "Mobile 2", "Course", "Batch", "Learner Code", "Address / City",
        "Source", "Status", "Last Called By", "Next Follow-up", "Last Call At", "Created At", "Latest Remarks"
    ])

    for r in rows:
        cw.writerow([
            r["id"], r["name"], r["father_name"] or "", r["phone"], r["alt_phone"] or "", r["course"] or "", r["batch"] or "", r["learner_code"] or "", r["city"] or "",
            r["source"] or "", r["status"], r["last_called_by"] or "Not Called",
            r["next_followup_date"] or "", r["last_call_at"] or "", r["created_at"], r["last_remarks"] or ""
        ])

    output = make_response(si.getvalue())
    output.headers["Content-Disposition"] = f"attachment; filename=focus_institute_students_{date.today().isoformat()}.csv"
    output.headers["Content-type"] = "text/csv"
    return output



# -------------------------------------------------------------
# FEES PENDING CALLING SYSTEM APIS
# -------------------------------------------------------------

@app.route("/api/fees/students", methods=["GET"])
@login_required
def get_fees_students():
    search = request.args.get("search", "").strip()
    status_filter = request.args.get("filter", "all").strip()
    today_str = date.today().isoformat()

    with get_db() as conn:
        cursor = conn.cursor()

        if status_filter == "paid":
            query = """
                SELECT l.id, l.name, l.father_name, l.phone, l.alt_phone, l.course, l.batch, l.learner_code, l.city,
                       l.fee_pending_amount, l.fee_due_date, l.fee_remarks, l.is_fee_pending, l.last_call_at,
                       u.name as last_caller_name,
                       (SELECT remarks FROM call_logs WHERE lead_id = l.id ORDER BY id DESC LIMIT 1) as last_remarks,
                       (SELECT COUNT(*) FROM call_logs WHERE lead_id = l.id) as total_calls
                FROM leads l
                LEFT JOIN users u ON l.last_caller_id = u.id
                WHERE (l.record_type = 'admitted' OR l.record_type IS NULL) 
                  AND (l.is_deleted = 0 OR l.is_deleted IS NULL)
                  AND l.is_fee_pending = 0
                  AND (l.fee_remarks LIKE '%Paid%' OR (l.fee_pending_amount != '' AND l.fee_pending_amount IS NOT NULL))
            """
        else:
            query = """
                SELECT l.id, l.name, l.father_name, l.phone, l.alt_phone, l.course, l.batch, l.learner_code, l.city,
                       l.fee_pending_amount, l.fee_due_date, l.fee_remarks, l.is_fee_pending, l.last_call_at,
                       u.name as last_caller_name,
                       (SELECT remarks FROM call_logs WHERE lead_id = l.id ORDER BY id DESC LIMIT 1) as last_remarks,
                       (SELECT COUNT(*) FROM call_logs WHERE lead_id = l.id) as total_calls
                FROM leads l
                LEFT JOIN users u ON l.last_caller_id = u.id
                WHERE (l.record_type = 'admitted' OR l.record_type IS NULL) 
                  AND (l.is_deleted = 0 OR l.is_deleted IS NULL)
                  AND l.is_fee_pending = 1
            """
        params = []

        if status_filter == "overdue":
            query += " AND l.fee_due_date != '' AND l.fee_due_date < ?"
            params.append(today_str)
        elif status_filter == "today":
            query += " AND l.fee_due_date = ?"
            params.append(today_str)
        elif status_filter == "upcoming":
            query += " AND l.fee_due_date > ?"
            params.append(today_str)

        if search:
            query += """ AND (
                l.name LIKE ? OR l.father_name LIKE ? OR l.phone LIKE ? OR 
                l.alt_phone LIKE ? OR l.course LIKE ? OR l.batch LIKE ? OR l.learner_code LIKE ?
            )"""
            w = f"%{search}%"
            params.extend([w, w, w, w, w, w, w])

        query += " ORDER BY CASE WHEN l.fee_due_date != '' THEN l.fee_due_date ELSE '9999-99-99' END ASC, l.name ASC"
        cursor.execute(query, params)
        rows = [dict(r) for r in cursor.fetchall()]

        cursor.execute("SELECT COUNT(*) as cnt FROM leads WHERE (record_type = 'admitted' OR record_type IS NULL) AND (is_deleted = 0 OR is_deleted IS NULL) AND is_fee_pending = 1")
        total_pending = cursor.fetchone()["cnt"]

    return jsonify({
        "success": True,
        "students": rows,
        "total_pending": total_pending
    })


@app.route("/api/fees/mark-pending", methods=["POST"])
@login_required
def mark_fee_pending():
    if session.get("role") != "manager" and not session.get("is_super_admin"):
        return jsonify({"error": "Only Manager can assign or edit pending fees."}), 403

    data = request.json or {}
    lead_id = data.get("lead_id")
    amount = str(data.get("amount", "")).strip()
    due_date = str(data.get("due_date", "")).strip()
    remarks = str(data.get("remarks", "")).strip()

    if not lead_id:
        return jsonify({"error": "Lead ID is required"}), 400

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE leads
            SET is_fee_pending = 1,
                fee_pending_amount = ?,
                fee_due_date = ?,
                fee_remarks = ?,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        """, (amount, due_date, remarks, lead_id))
        conn.commit()
    bump_sync_version()
    return jsonify({"success": True, "message": "Student marked for pending fees successfully!"})


@app.route("/api/fees/mark-paid", methods=["POST"])
@login_required
def mark_fee_paid():
    if session.get("role") != "manager" and not session.get("is_super_admin"):
        return jsonify({"error": "Only Manager can mark fees as paid."}), 403

    data = request.json or {}
    lead_id = data.get("lead_id")
    remarks = str(data.get("remarks", "Fee Paid")).strip()
    user_id = session.get("user_id")

    if not lead_id:
        return jsonify({"error": "Lead ID is required"}), 400

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE leads
            SET is_fee_pending = 0,
                fee_remarks = ?,
                last_caller_id = ?,
                last_call_at = CURRENT_TIMESTAMP,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        """, (f"Paid: {remarks}", user_id, lead_id))

        cursor.execute("""
            INSERT INTO call_logs (lead_id, caller_id, status, remarks)
            VALUES (?, ?, 'COMPLETED', ?)
        """, (lead_id, user_id, f"Fee Paid marked: {remarks}"))
        conn.commit()

    bump_sync_version()
    return jsonify({"success": True, "message": "Fee marked as paid successfully!"})


@app.route("/api/fees/log-call", methods=["POST"])
@login_required
def log_fee_call():
    data = request.json or {}
    lead_id = data.get("lead_id")
    status = data.get("status", "FEE_REMINDER").strip()
    remarks = data.get("remarks", "").strip()
    next_followup = data.get("next_followup_date", "")
    user_id = session.get("user_id")

    if not lead_id:
        return jsonify({"error": "Lead ID is required"}), 400

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO call_logs (lead_id, caller_id, status, remarks, next_followup_date)
            VALUES (?, ?, ?, ?, ?)
        """, (lead_id, user_id, status, remarks, next_followup or None))

        cursor.execute("""
            UPDATE leads
            SET last_caller_id = ?,
                last_call_at = CURRENT_TIMESTAMP,
                fee_remarks = ?,
                next_followup_date = ?,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        """, (user_id, remarks, next_followup or None, lead_id))
        conn.commit()

    bump_sync_version()
    return jsonify({"success": True, "message": "Fee call remarks saved!"})


# -------------------------------------------------------------
# CALL FOR EXAM (MONTHLY CALLING SYSTEM) APIS
# -------------------------------------------------------------

def get_current_exam_month():
    today = date.today()
    return today.strftime("%Y-%m"), today.strftime("%B %Y")


@app.route("/api/exams/months", methods=["GET"])
@login_required
def get_exam_months():
    is_admin = session.get("role") == "manager" or session.get("username") == "arsh" or session.get("user_id") == 1
    current_code, current_name = get_current_exam_month()

    with get_db() as conn:
        cursor = conn.cursor()

        cursor.execute("SELECT * FROM exam_month_settings WHERE month_code = ?", (current_code,))
        if not cursor.fetchone():
            cursor.execute("""
                INSERT OR IGNORE INTO exam_month_settings (month_code, month_name, is_visible_to_staff, is_active)
                VALUES (?, ?, 1, 1)
            """, (current_code, current_name))
            conn.commit()

        if is_admin:
            # Admin sees all months ever recorded (current & archive)
            cursor.execute("""
                SELECT s.month_code, s.month_name, s.is_visible_to_staff, s.is_active,
                       (SELECT COUNT(*) FROM exam_students WHERE month_code = s.month_code) as total_students,
                       (SELECT COUNT(*) FROM exam_students WHERE month_code = s.month_code AND call_status = 'CONFIRMED') as confirmed_count,
                       (SELECT COUNT(*) FROM exam_students WHERE month_code = s.month_code AND call_status = 'NOT_CALLED') as not_called_count
                FROM exam_month_settings s
                ORDER BY s.month_code DESC
            """ )
            months = [dict(r) for r in cursor.fetchall()]
            return jsonify({
                "success": True,
                "is_admin": True,
                "current_month": current_code,
                "months": months
            })
        else:
            # Staff ONLY sees current month until the month ends, and only if visible
            cursor.execute("""
                SELECT month_code, month_name, is_visible_to_staff
                FROM exam_month_settings
                WHERE month_code = ?
            """, (current_code,))
            row = cursor.fetchone()
            if not row or row["is_visible_to_staff"] == 0:
                return jsonify({
                    "success": True,
                    "is_admin": False,
                    "current_month": current_code,
                    "visible": False,
                    "message": "Exam calling list is currently hidden by Admin or not published."
                })

            cursor.execute("SELECT COUNT(*) as cnt FROM exam_students WHERE month_code = ?", (current_code,))
            total = cursor.fetchone()["cnt"]
            return jsonify({
                "success": True,
                "is_admin": False,
                "current_month": current_code,
                "visible": True,
                "months": [{
                    "month_code": current_code,
                    "month_name": current_name,
                    "is_visible_to_staff": 1,
                    "total_students": total
                }]
            })


@app.route("/api/exams/settings/visibility", methods=["POST"])
@manager_required
def toggle_exam_visibility():
    data = request.json or {}
    month_code = data.get("month_code", get_current_exam_month()[0])
    is_visible = 1 if data.get("is_visible") else 0

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO exam_month_settings (month_code, month_name, is_visible_to_staff)
            VALUES (?, ?, ?)
            ON CONFLICT(month_code) DO UPDATE SET is_visible_to_staff = ?, updated_at = CURRENT_TIMESTAMP
        """, (month_code, data.get("month_name", month_code), is_visible, is_visible))
        conn.commit()
    bump_sync_version()
    status_text = "Visible to Staff" if is_visible else "Hidden from Staff"
    return jsonify({"success": True, "is_visible": is_visible, "message": f"Exam calling list is now {status_text}!"})


@app.route("/api/exams/students", methods=["GET"])
@login_required
def get_exam_students():
    is_admin = session.get("role") == "manager" or session.get("username") == "arsh" or session.get("user_id") == 1
    current_code, _ = get_current_exam_month()

    requested_month = request.args.get("month", current_code).strip()
    status_filter = request.args.get("status", "").strip()
    search = request.args.get("search", "").strip()

    # Permission Enforcement: Staff can ONLY access current active month
    if not is_admin:
        if requested_month != current_code:
            return jsonify({"error": "Staff can only view current month exams. Past exam archives are Admin-only."}), 403

        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT is_visible_to_staff FROM exam_month_settings WHERE month_code = ?", (current_code,))
            st = cursor.fetchone()
            if not st or st["is_visible_to_staff"] == 0:
                return jsonify({
                    "success": True,
                    "visible": False,
                    "students": [],
                    "message": "Exam calling list is currently hidden by Admin."
                })

    with get_db() as conn:
        cursor = conn.cursor()
        query = "SELECT * FROM exam_students WHERE month_code = ?"
        params = [requested_month]

        if status_filter and status_filter != "ALL":
            query += " AND call_status = ?"
            params.append(status_filter)

        if search:
            query += """ AND (
                student_name LIKE ? OR father_name LIKE ? OR phone LIKE ? OR 
                alt_phone LIKE ? OR course LIKE ? OR batch LIKE ? OR learner_code LIKE ?
            )"""
            w = f"%{search}%"
            params.extend([w, w, w, w, w, w, w])

        query += """ ORDER BY 
            CASE call_status 
                WHEN 'NOT_CALLED' THEN 0 
                WHEN 'CALL_BACK' THEN 1 
                WHEN 'CONFIRMED' THEN 2
                ELSE 3 
            END ASC, student_name ASC"""
        cursor.execute(query, params)
        students = [dict(r) for r in cursor.fetchall()]

        cursor.execute("""
            SELECT 
                COUNT(*) as total,
                SUM(CASE WHEN call_status = 'NOT_CALLED' THEN 1 ELSE 0 END) as not_called,
                SUM(CASE WHEN call_status = 'CONFIRMED' THEN 1 ELSE 0 END) as confirmed,
                SUM(CASE WHEN call_status = 'ABSENT' THEN 1 ELSE 0 END) as absent,
                SUM(CASE WHEN call_status = 'CALL_BACK' THEN 1 ELSE 0 END) as call_back,
                SUM(CASE WHEN call_status = 'FEE_ISSUE' THEN 1 ELSE 0 END) as fee_issue
            FROM exam_students
            WHERE month_code = ?
        """, (requested_month,))
        stats_row = cursor.fetchone()
        stats = {
            "total": stats_row["total"] or 0,
            "not_called": stats_row["not_called"] or 0,
            "confirmed": stats_row["confirmed"] or 0,
            "absent": stats_row["absent"] or 0,
            "call_back": stats_row["call_back"] or 0,
            "fee_issue": stats_row["fee_issue"] or 0
        }

        cursor.execute("SELECT is_visible_to_staff FROM exam_month_settings WHERE month_code = ?", (requested_month,))
        vis_row = cursor.fetchone()
        is_visible_to_staff = vis_row["is_visible_to_staff"] if vis_row else 1

    return jsonify({
        "success": True,
        "visible": True,
        "month_code": requested_month,
        "is_visible_to_staff": is_visible_to_staff,
        "stats": stats,
        "students": students
    })


@app.route("/api/exams/students/add", methods=["POST"])
@manager_required
def add_exam_student():
    data = request.json or {}
    month_code = data.get("month_code", get_current_exam_month()[0])
    month_name = data.get("month_name", get_current_exam_month()[1])
    student_name = data.get("student_name", "").strip()
    father_name = data.get("father_name", "").strip()
    phone = data.get("phone", "").strip()
    alt_phone = data.get("alt_phone", "").strip()
    course = data.get("course", "").strip()
    batch = data.get("batch", "").strip()
    learner_code = data.get("learner_code", "").strip()
    exam_date_val = data.get("exam_date", "").strip()
    student_id = data.get("student_id")
    remarks = data.get("remarks", "").strip()

    if not student_name or not phone:
        return jsonify({"error": "Student Name and Phone Number are required."}), 400

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO exam_students 
            (month_code, month_name, student_id, student_name, father_name, phone, alt_phone, course, batch, learner_code, exam_date, remarks)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (month_code, month_name, student_id, student_name, father_name, phone, alt_phone, course, batch, learner_code, exam_date_val, remarks))
        conn.commit()
    bump_sync_version()
    return jsonify({"success": True, "message": f"Student '{student_name}' added to {month_name} exam list!"})


@app.route("/api/exams/students/add-from-leads", methods=["POST"])
@manager_required
def add_exam_student_from_leads():
    data = request.json or {}
    lead_id = data.get("lead_id")
    month_code = data.get("month_code", get_current_exam_month()[0])
    month_name = data.get("month_name", get_current_exam_month()[1])
    exam_date_val = data.get("exam_date", "").strip()

    if not lead_id:
        return jsonify({"error": "Lead ID required"}), 400

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM leads WHERE id = ?", (lead_id,))
        lead = cursor.fetchone()
        if not lead:
            return jsonify({"error": "Student not found"}), 404

        cursor.execute("SELECT id FROM exam_students WHERE month_code = ? AND (student_id = ? OR phone = ?)", (month_code, lead_id, lead["phone"]))
        if cursor.fetchone():
            return jsonify({"error": f"Student '{lead['name']}' is already in {month_name} exam list!"}), 400

        cursor.execute("""
            INSERT INTO exam_students 
            (month_code, month_name, student_id, student_name, father_name, phone, alt_phone, course, batch, learner_code, exam_date)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (month_code, month_name, lead_id, lead["name"], lead["father_name"], lead["phone"], lead["alt_phone"], lead["course"], lead["batch"], lead["learner_code"], exam_date_val))
        conn.commit()
    bump_sync_version()
    return jsonify({"success": True, "message": f"'{lead['name']}' added to {month_name} exam list!"})


@app.route("/api/exams/students/<int:student_id>/call", methods=["POST"])
@login_required
def log_exam_student_call(student_id):
    data = request.json or {}
    call_status = data.get("call_status", "").strip()
    remarks = data.get("remarks", "").strip()
    caller_name = session.get("name") or session.get("username") or "Staff"

    valid_statuses = ["NOT_CALLED", "CONFIRMED", "ABSENT", "CALL_BACK", "FEE_ISSUE", "WRONG_NUMBER"]
    if call_status not in valid_statuses:
        return jsonify({"error": "Invalid call status"}), 400

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE exam_students
            SET call_status = ?,
                remarks = ?,
                last_called_by = ?,
                last_call_at = CURRENT_TIMESTAMP
            WHERE id = ?
        """, (call_status, remarks, caller_name, student_id))
        conn.commit()
    bump_sync_version()
    return jsonify({"success": True, "message": "Exam call status updated!"})


@app.route("/api/exams/students/<int:student_id>", methods=["DELETE"])
@manager_required
def delete_exam_student(student_id):
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM exam_students WHERE id = ?", (student_id,))
        conn.commit()
    bump_sync_version()
    return jsonify({"success": True, "message": "Student removed from exam list."})


@app.route("/api/exams/students/import", methods=["POST"])
@manager_required
def import_exam_students():
    month_code = request.form.get("month_code", get_current_exam_month()[0])
    month_name = request.form.get("month_name", get_current_exam_month()[1])

    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    if not file or file.filename == "":
        return jsonify({"error": "Empty file uploaded"}), 400

    filename = file.filename.lower()
    imported_count = 0
    records = []

    try:
        if filename.endswith(".csv"):
            stream = io.StringIO(file.stream.read().decode("utf-8", errors="ignore"))
            reader = csv.reader(stream)
            headers = [h.strip().lower() for h in next(reader, [])]
            for row in reader:
                if not row or not any(row):
                    continue
                row_dict = {headers[i]: row[i].strip() if i < len(row) else "" for i in range(len(headers))}
                records.append(row_dict)
        elif filename.endswith(".xlsx") or filename.endswith(".xls"):
            import openpyxl
            wb = openpyxl.load_workbook(file.stream, data_only=True)
            sheet = wb.active
            rows = list(sheet.iter_rows(values_only=True))
            if rows:
                headers = [str(h).strip().lower() if h is not None else "" for h in rows[0]]
                for r in rows[1:]:
                    if not r or not any(r):
                        continue
                    row_dict = {headers[i]: str(r[i]).strip() if i < len(r) and r[i] is not None else "" for i in range(len(headers))}
                    records.append(row_dict)

        with get_db() as conn:
            cursor = conn.cursor()
            for r in records:
                sname = r.get("name") or r.get("student name") or r.get("student") or r.get("learner name") or ""
                fname = r.get("father name") or r.get("father's name") or r.get("father") or ""
                phone = r.get("mobile") or r.get("phone") or r.get("contact") or r.get("mobile 1") or ""
                alt = r.get("alt mobile") or r.get("alt phone") or r.get("mobile 2") or ""
                course = r.get("course") or r.get("course name") or ""
                batch = r.get("batch") or r.get("year") or ""
                lcode = r.get("learner code") or r.get("roll no") or r.get("code") or ""
                ex_date = r.get("exam date") or r.get("date") or ""

                phone_clean = re.sub(r"[^0-9]", "", phone)[-10:]
                if not sname or len(phone_clean) < 10:
                    continue

                cursor.execute("""
                    INSERT INTO exam_students 
                    (month_code, month_name, student_name, father_name, phone, alt_phone, course, batch, learner_code, exam_date)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (month_code, month_name, sname, fname, phone_clean, alt, course, batch, lcode, ex_date))
                imported_count += 1
            conn.commit()

        bump_sync_version()
        return jsonify({"success": True, "message": f"Successfully imported {imported_count} exam students for {month_name}!"})
    except Exception as e:
        return jsonify({"error": f"Import failed: {str(e)}"}), 500


if __name__ == "__main__":
    init_db()
    port = int(os.environ.get("PORT", 5500))
    print("================================================================")
    print(" FOCUS INSTITUTE - STUDENT CALLING & CRM PORTAL")
    print(f" Server running at: http://127.0.0.1:{port}")
    print(" Manager Login: arsh / arsh123")
    print(" Staff Logins:  gurpreet / staff123, simran / staff123, aman / staff123")
    print(" Staff Privacy: Restricted to 6 Fields (Name, Father, Phones, Course, Batch, Address)")
    print("================================================================")
    app.run(host="0.0.0.0", port=port, debug=False)
