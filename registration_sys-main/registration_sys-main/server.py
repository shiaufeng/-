import os
import re
import csv
import io
import json
from datetime import datetime
from urllib.parse import urlparse, quote

from flask import Flask, request, jsonify, session, send_from_directory, Response
from flask_cors import CORS
import pymysql
from pymysql.cursors import DictCursor

app = Flask(__name__, static_folder='.', static_url_path='')
app.secret_key = os.getenv('SECRET_KEY', 'smart-ark-dev-secret')
CORS(app, supports_credentials=True)

DEFAULT_ADMIN = os.getenv('ADMIN_USERNAME', 'admin')
DEFAULT_PASSWORD = os.getenv('ADMIN_PASSWORD', 'admin123')
DEFAULT_SHEET = os.getenv('ADMIN_DEFAULT_EVENT') or os.getenv('ADMIN_DEFAULT_EVENTS') or '活動報到名單'
CHECKED_STATUSES = ('checked_in', '已報到', '替代', 'done')

# ============================================================
# DB helpers
# ============================================================

def db_params():
    url = os.getenv('DATABASE_URL') or os.getenv('MYSQL_URL')
    if url:
        u = urlparse(url)
        return dict(
            host=u.hostname,
            port=u.port or 3306,
            user=u.username,
            password=u.password,
            database=(u.path or '').lstrip('/') or os.getenv('MYSQLDATABASE'),
            charset='utf8mb4',
            cursorclass=DictCursor,
            autocommit=False,
        )
    return dict(
        host=os.getenv('MYSQLHOST') or os.getenv('DB_HOST') or 'localhost',
        port=int(os.getenv('MYSQLPORT') or os.getenv('DB_PORT') or 3306),
        user=os.getenv('MYSQLUSER') or os.getenv('DB_USER') or 'root',
        password=os.getenv('MYSQLPASSWORD') or os.getenv('DB_PASSWORD') or '',
        database=os.getenv('MYSQLDATABASE') or os.getenv('DB_NAME') or 'railway',
        charset='utf8mb4',
        cursorclass=DictCursor,
        autocommit=False,
    )


def get_db_connection():
    return pymysql.connect(**db_params())


def column_exists(cur, table, col):
    cur.execute(
        """
        SELECT COUNT(*) AS c
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = DATABASE()
          AND TABLE_NAME = %s
          AND COLUMN_NAME = %s
        """,
        (table, col),
    )
    return int((cur.fetchone() or {}).get('c') or 0) > 0


def table_columns(cur, table):
    cur.execute(
        """
        SELECT COLUMN_NAME
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = DATABASE()
          AND TABLE_NAME = %s
        """,
        (table,),
    )
    return {r['COLUMN_NAME'] for r in cur.fetchall()}


def add_col(cur, table, col, spec):
    if not column_exists(cur, table, col):
        cur.execute(f"ALTER TABLE `{table}` ADD COLUMN `{col}` {spec}")


def try_sql(cur, sql, args=None):
    try:
        cur.execute(sql, args or ())
    except Exception:
        pass


def ensure_core_tables(conn):
    with conn.cursor() as cur:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS admins (
                id INT AUTO_INCREMENT PRIMARY KEY,
                username VARCHAR(120) UNIQUE NOT NULL,
                password VARCHAR(255) NOT NULL,
                allowed_events LONGTEXT,
                current_event VARCHAR(255),
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """
        )

        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS event_configs (
                id INT AUTO_INCREMENT PRIMARY KEY,
                admin_username VARCHAR(120) NOT NULL DEFAULT 'admin',
                google_sheet_name VARCHAR(255) NOT NULL DEFAULT '活動報到名單',
                event_title VARCHAR(255),
                event_subtitle VARCHAR(255),
                date_start VARCHAR(80),
                date_end VARCHAR(80),
                brand_name VARCHAR(255),
                logo_url LONGTEXT,
                banner_image_url LONGTEXT,
                map_image_url LONGTEXT,
                products LONGTEXT,
                product_categories LONGTEXT,
                industry_mappings LONGTEXT,
                agenda LONGTEXT,
                event_config LONGTEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                updated_at DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                UNIQUE KEY uq_event_config (admin_username, google_sheet_name)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """
        )

        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS event_registrations (
                id INT AUTO_INCREMENT PRIMARY KEY,
                admin_username VARCHAR(120) NOT NULL DEFAULT 'admin',
                google_sheet_name VARCHAR(255) NOT NULL DEFAULT '活動報到名單',
                admin_user VARCHAR(120) DEFAULT 'admin',
                event_key VARCHAR(255) DEFAULT '活動報到名單',
                name VARCHAR(255),
                phone VARCHAR(100),
                email VARCHAR(255),
                company VARCHAR(255),
                company_name VARCHAR(255),
                job_title VARCHAR(255),
                region VARCHAR(255),
                training_level VARCHAR(255),
                seat VARCHAR(100),
                seating_chart VARCHAR(100),
                status VARCHAR(40) DEFAULT 'pending',
                is_original TINYINT(1) DEFAULT 1,
                proxy_name VARCHAR(255),
                proxy_phone VARCHAR(100),
                checked_in_at DATETIME NULL,
                checkin_time DATETIME NULL,
                portrait_consent TINYINT(1) NULL,
                portrait_consent_status VARCHAR(40),
                portrait_consent_time DATETIME NULL,
                special_notes LONGTEXT,
                note LONGTEXT,
                raw_data LONGTEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                updated_at DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                KEY idx_event (admin_username, google_sheet_name),
                KEY idx_status (status)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """
        )

        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS agenda_items (
                id INT AUTO_INCREMENT PRIMARY KEY,
                admin_username VARCHAR(120) NOT NULL DEFAULT 'admin',
                google_sheet_name VARCHAR(255) NOT NULL DEFAULT '活動報到名單',
                sort_order INT DEFAULT 0,
                time VARCHAR(60),
                event VARCHAR(255),
                title VARCHAR(255),
                description LONGTEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                KEY idx_agenda (admin_username, google_sheet_name)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """
        )

        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS industry_mappings (
                id INT AUTO_INCREMENT PRIMARY KEY,
                admin_username VARCHAR(120) NOT NULL DEFAULT 'admin',
                google_sheet_name VARCHAR(255) NOT NULL DEFAULT '活動報到名單',
                company_name VARCHAR(255),
                keyword VARCHAR(255),
                category VARCHAR(255),
                industry VARCHAR(255),
                sort_order INT DEFAULT 0,
                KEY idx_industry (admin_username, google_sheet_name)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """
        )

        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS exhibitors (
                id INT AUTO_INCREMENT PRIMARY KEY,
                admin_username VARCHAR(120) NOT NULL DEFAULT 'admin',
                google_sheet_name VARCHAR(255) NOT NULL DEFAULT '活動報到名單',
                name VARCHAR(255),
                company_name VARCHAR(255),
                industry VARCHAR(255),
                image_url LONGTEXT,
                logo LONGTEXT,
                website LONGTEXT,
                contact VARCHAR(255),
                description LONGTEXT,
                sort_order INT DEFAULT 0,
                KEY idx_exhibitor (admin_username, google_sheet_name)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """
        )

        cfg_cols = {
            'admin_username': "VARCHAR(120) NOT NULL DEFAULT 'admin'",
            'google_sheet_name': "VARCHAR(255) NOT NULL DEFAULT '活動報到名單'",
            'event_title': 'VARCHAR(255)',
            'event_subtitle': 'VARCHAR(255)',
            'event_date_start': 'VARCHAR(80)',
            'event_date_end': 'VARCHAR(80)',
            'card1_icon': 'VARCHAR(40)',
            'card1_title': 'VARCHAR(255)',
            'card1_subtitle': 'VARCHAR(255)',
            'card1_description': 'LONGTEXT',
            'card1_url': 'LONGTEXT',
            'card2_icon': 'VARCHAR(40)',
            'card2_title': 'VARCHAR(255)',
            'card2_subtitle': 'VARCHAR(255)',
            'card2_description': 'LONGTEXT',
            'card2_url': 'LONGTEXT',
            'card3_icon': 'VARCHAR(40)',
            'card3_title': 'VARCHAR(255)',
            'card3_subtitle': 'VARCHAR(255)',
            'card3_description': 'LONGTEXT',
            'card3_url': 'LONGTEXT',
            'card4_icon': 'VARCHAR(40)',
            'card4_title': 'VARCHAR(255)',
            'card4_subtitle': 'VARCHAR(255)',
            'card4_description': 'LONGTEXT',
            'card4_url': 'LONGTEXT',
            'products': 'LONGTEXT',
            'product_categories': 'LONGTEXT',
            'industry_mappings': 'LONGTEXT',
            'agenda': 'LONGTEXT',
            'event_config': 'LONGTEXT',
        }
        for col, spec in cfg_cols.items():
            add_col(cur, 'event_configs', col, spec)

        cur.execute("SELECT id FROM admins WHERE username=%s", (DEFAULT_ADMIN,))
        if not cur.fetchone():
            cur.execute(
                "INSERT INTO admins (username,password,allowed_events,current_event) VALUES (%s,%s,%s,%s)",
                (DEFAULT_ADMIN, DEFAULT_PASSWORD, DEFAULT_SHEET, DEFAULT_SHEET),
            )
    conn.commit()


# ============================================================
# Generic helpers
# ============================================================

def get_payload():
    return request.get_json(silent=True) if request.is_json else None


def event_args():
    data = get_payload() or {}
    admin = (
        request.args.get('admin')
        or request.form.get('admin')
        or data.get('admin')
        or data.get('admin_username')
        or session.get('username')
        or DEFAULT_ADMIN
    )
    sheet = (
        request.args.get('sheet')
        or request.args.get('google_sheet_name')
        or request.args.get('event_key')
        or request.form.get('sheet')
        or data.get('sheet')
        or data.get('google_sheet_name')
        or session.get('current_admin_sheet')
        or DEFAULT_SHEET
    )
    return str(admin).strip() or DEFAULT_ADMIN, str(sheet).strip() or DEFAULT_SHEET


def q_event():
    return event_args()


def json_loads(value, default=None):
    if default is None:
        default = []
    if value is None or value == '':
        return default
    if isinstance(value, (dict, list)):
        return value
    try:
        return json.loads(value)
    except Exception:
        return default


def json_dumps(value):
    return json.dumps(value if value is not None else [], ensure_ascii=False)


def pick(row, keys):
    for k in keys:
        if k in row and str(row[k]).strip():
            return str(row[k]).strip()
    return ''


def clean_text(v):
    return str(v or '').strip()


def search_norm(v):
    text = str(v or '').lower().strip()
    text = re.sub(r'[\s\-_/\.、，,。．·・:：;；\(\)（）\[\]【】{}「」『』\+]+', '', text)
    return text


def search_tokens(v):
    raw = str(v or '').strip()
    tokens = [search_norm(x) for x in re.split(r'[\s、，,]+', raw) if search_norm(x)]
    whole = search_norm(raw)
    if whole and whole not in tokens:
        tokens.insert(0, whole)
    return tokens


def search_match_score(row, query, method='company'):
    q = search_norm(query)
    tokens = search_tokens(query)
    company = search_norm(row.get('company') or row.get('company_name'))
    name = search_norm(row.get('name'))
    phone = search_norm(row.get('phone'))
    email = search_norm(row.get('email'))
    job = search_norm(row.get('job_title'))
    hay_company = company
    hay_all = ''.join([company, name, phone, email, job])
    if not q and not tokens:
        return 0
    if method in ['company', 'company_name', 'unit']:
        if q and q in hay_company:
            return 100
        if tokens and all(t in hay_company for t in tokens):
            return 95
        if q and q in hay_all:
            return 80
        if tokens and all(t in hay_all for t in tokens):
            return 75
        if tokens and any(t in hay_company for t in tokens):
            return 50
    else:
        if q and q in hay_all:
            return 90
        if tokens and all(t in hay_all for t in tokens):
            return 80
    return 0


def table_sort_key(label):
    s = clean_text(label)
    nums = re.findall(r'\d+', s)
    return (0, int(nums[0])) if nums else (1, s)


def normalize_table_label(value):
    v = clean_text(value)
    if not v:
        return ''
    v = v.replace('桌', '').replace('第', '').strip()
    return v or clean_text(value)


def status_checked(status):
    return str(status or '').lower() in ['checked_in', '已報到', '替代', 'done']


def normalize_registration(row):
    return {
        'name': pick(row, ['姓名', 'name', 'Name', '名字', '貴賓姓名']),
        'phone': pick(row, ['手機', '電話', 'phone', 'Phone', '行動電話', '手機號碼']),
        'email': pick(row, ['Email', 'email', 'E-mail', '信箱', '電子郵件']),
        'company': pick(row, ['公司', '公司名稱', '服務單位', '單位', 'company', 'Company']),
        'job_title': pick(row, ['職稱', 'title', 'job_title', '職位', 'position', 'Position']),
        'region': pick(row, ['地區', '區域', 'region', 'Region']),
        'training_level': pick(row, ['職階', '層級', 'training_level', 'level']),
        'seat': normalize_table_label(pick(row, ['桌號', '座位', '桌次', 'seat', 'Seat', 'seating_chart'])),
        'special_notes': pick(row, ['備註', 'notes', 'note', 'Remarks']),
        'raw_data': json_dumps(row),
    }


def portrait_status_from_row(row):
    raw = clean_text((row or {}).get('portrait_consent_status'))
    if raw:
        return raw
    v = (row or {}).get('portrait_consent')
    if v in [1, True, '1', 'true', 'True', '同意', 'yes', 'Yes']:
        return '同意'
    if v in [0, False, '0', 'false', 'False', '不同意', 'no', 'No']:
        return '不同意'
    return '未填'


def public_user(row):
    r = dict(row or {})
    company = r.get('company') or r.get('company_name') or ''
    seat = r.get('seat') or r.get('seating_chart') or ''
    checked_time = r.get('checked_in_at') or r.get('checkin_time')
    portrait_status = portrait_status_from_row(r)
    result = {
        **r,
        'company': company,
        'company_name': company,
        'seat': seat,
        'table': seat,
        'seating_chart': seat,
        'job_title': r.get('job_title') or '',
        'portrait_consent_status': portrait_status,
        'portrait_consent': True if portrait_status == '同意' else (False if portrait_status == '不同意' else None),
        'checked_in_at': checked_time.strftime('%Y-%m-%d %H:%M:%S') if hasattr(checked_time, 'strftime') else (checked_time or ''),
        'checkin_time': checked_time.strftime('%Y-%m-%d %H:%M:%S') if hasattr(checked_time, 'strftime') else (checked_time or ''),
        'checkedInAt': checked_time.strftime('%H:%M:%S') if hasattr(checked_time, 'strftime') else (checked_time or ''),
        'special_notes': r.get('special_notes') or r.get('note') or '',
        'note': r.get('special_notes') or r.get('note') or '',
    }
    return result


def ensure_config(conn, admin, sheet):
    with conn.cursor() as cur:
        cur.execute(
            "SELECT id FROM event_configs WHERE admin_username=%s AND google_sheet_name=%s LIMIT 1",
            (admin, sheet),
        )
        if not cur.fetchone():
            cur.execute(
                """
                INSERT INTO event_configs (admin_username, google_sheet_name, event_title, event_subtitle, brand_name)
                VALUES (%s,%s,%s,%s,%s)
                """,
                (admin, sheet, sheet, '世代共榮的數位聚合', '智匯方舟'),
            )
    conn.commit()


def get_config_row(conn, admin, sheet):
    ensure_config(conn, admin, sheet)
    with conn.cursor() as cur:
        cur.execute("SELECT * FROM event_configs WHERE admin_username=%s AND google_sheet_name=%s LIMIT 1", (admin, sheet))
        return cur.fetchone() or {}


def load_agenda(conn, admin, sheet):
    with conn.cursor() as cur:
        cur.execute(
            "SELECT time, COALESCE(NULLIF(title,''), event) AS title, event, description, sort_order FROM agenda_items WHERE admin_username=%s AND google_sheet_name=%s ORDER BY sort_order ASC, id ASC",
            (admin, sheet),
        )
        rows = cur.fetchall()
    if rows:
        return [dict(r) for r in rows]
    cfg = get_config_row(conn, admin, sheet)
    return json_loads(cfg.get('agenda'), [])


def save_agenda(conn, admin, sheet, rows):
    rows = rows or []
    ensure_config(conn, admin, sheet)
    with conn.cursor() as cur:
        cur.execute("DELETE FROM agenda_items WHERE admin_username=%s AND google_sheet_name=%s", (admin, sheet))
        for i, item in enumerate(rows):
            time = clean_text(item.get('time'))
            title = clean_text(item.get('title') or item.get('event'))
            desc = clean_text(item.get('description'))
            if not (time or title or desc):
                continue
            cur.execute(
                """
                INSERT INTO agenda_items (admin_username, google_sheet_name, sort_order, time, event, title, description)
                VALUES (%s,%s,%s,%s,%s,%s,%s)
                """,
                (admin, sheet, int(item.get('sort_order', i) or i), time, title, title, desc),
            )
        cur.execute("UPDATE event_configs SET agenda=%s WHERE admin_username=%s AND google_sheet_name=%s", (json_dumps(rows), admin, sheet))
    conn.commit()


def load_industry_mappings(conn, admin, sheet):
    with conn.cursor() as cur:
        cur.execute(
            "SELECT company_name, keyword, category, industry, sort_order FROM industry_mappings WHERE admin_username=%s AND google_sheet_name=%s ORDER BY sort_order ASC, id ASC",
            (admin, sheet),
        )
        rows = cur.fetchall()
    if rows:
        out = []
        for r in rows:
            keyword = r.get('keyword') or r.get('company_name') or ''
            category = r.get('category') or r.get('industry') or '其他'
            out.append({**r, 'keyword': keyword, 'category': category, 'company': r.get('company_name') or keyword})
        return out
    cfg = get_config_row(conn, admin, sheet)
    return json_loads(cfg.get('industry_mappings'), [])


def save_industry_mappings(conn, admin, sheet, rows):
    rows = rows or []
    ensure_config(conn, admin, sheet)
    with conn.cursor() as cur:
        cur.execute("DELETE FROM industry_mappings WHERE admin_username=%s AND google_sheet_name=%s", (admin, sheet))
        for i, item in enumerate(rows):
            company = clean_text(item.get('company_name') or item.get('company') or item.get('keyword'))
            keyword = clean_text(item.get('keyword') or company)
            category = clean_text(item.get('category') or item.get('industry')) or '其他'
            if not (company or keyword or category):
                continue
            cur.execute(
                """
                INSERT INTO industry_mappings (admin_username, google_sheet_name, company_name, keyword, category, industry, sort_order)
                VALUES (%s,%s,%s,%s,%s,%s,%s)
                """,
                (admin, sheet, company, keyword, category, category, int(item.get('sort_order', i) or i)),
            )
        cur.execute("UPDATE event_configs SET industry_mappings=%s WHERE admin_username=%s AND google_sheet_name=%s", (json_dumps(rows), admin, sheet))
    conn.commit()


def load_exhibitors(conn, admin, sheet):
    with conn.cursor() as cur:
        cur.execute(
            "SELECT name, company_name, industry, image_url, logo, website, contact, description, sort_order FROM exhibitors WHERE admin_username=%s AND google_sheet_name=%s ORDER BY sort_order ASC, id ASC",
            (admin, sheet),
        )
        rows = cur.fetchall()
    out = []
    for r in rows:
        name = r.get('name') or r.get('company_name') or ''
        image = r.get('image_url') or r.get('logo') or ''
        out.append({**r, 'name': name, 'company_name': name, 'image_url': image, 'logo': image})
    return out


def save_exhibitors(conn, admin, sheet, rows):
    rows = rows or []
    with conn.cursor() as cur:
        cur.execute("DELETE FROM exhibitors WHERE admin_username=%s AND google_sheet_name=%s", (admin, sheet))
        for i, item in enumerate(rows):
            name = clean_text(item.get('name') or item.get('company_name'))
            if not name:
                continue
            image = clean_text(item.get('image_url') or item.get('logo'))
            cur.execute(
                """
                INSERT INTO exhibitors (admin_username, google_sheet_name, name, company_name, industry, image_url, logo, website, contact, description, sort_order)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                """,
                (
                    admin, sheet, name, name,
                    clean_text(item.get('industry')),
                    image, image,
                    clean_text(item.get('website')),
                    clean_text(item.get('contact')),
                    clean_text(item.get('description')),
                    int(item.get('sort_order', i) or i),
                ),
            )
    conn.commit()


def load_products_from_config(row):
    products = json_loads((row or {}).get('products'), [])
    if isinstance(products, dict):
        products = products.get('items') or products.get('products') or []
    return products if isinstance(products, list) else []


def save_config(conn, admin, sheet, payload):
    payload = dict(payload or {})
    ensure_config(conn, admin, sheet)
    cols = None
    with conn.cursor() as cur:
        cols = table_columns(cur, 'event_configs')
        direct = {}
        extra = {}
        for k, v in payload.items():
            if k in ['products', 'industry_mappings', 'agenda', 'product_categories']:
                direct[k] = json_dumps(v)
            elif k in cols and k not in ['id', 'created_at', 'updated_at', 'success_card_config', 'success_info_cards_config', 'dashboard_agenda_config']:
                direct[k] = v
            else:
                extra[k] = v
        if extra:
            row = get_config_row(conn, admin, sheet)
            old = json_loads(row.get('event_config'), {})
            if not isinstance(old, dict):
                old = {}
            old.update(extra)
            direct['event_config'] = json_dumps(old)
        if direct:
            set_sql = ', '.join([f"`{k}`=%s" for k in direct])
            cur.execute(f"UPDATE event_configs SET {set_sql} WHERE admin_username=%s AND google_sheet_name=%s", list(direct.values()) + [admin, sheet])
    conn.commit()


def serialize_config(row, conn=None, admin=None, sheet=None):
    cfg = dict(row or {})
    extra = json_loads(cfg.get('event_config'), {})
    if isinstance(extra, dict):
        cfg.update({k: v for k, v in extra.items() if k not in cfg or cfg.get(k) in [None, '']})
    cfg['google_sheet_name'] = cfg.get('google_sheet_name') or sheet or DEFAULT_SHEET
    cfg['admin_username'] = cfg.get('admin_username') or admin or DEFAULT_ADMIN
    cfg['event_title'] = cfg.get('event_title') or cfg.get('exp_event_title') or cfg['google_sheet_name']
    cfg['event_subtitle'] = cfg.get('event_subtitle') or cfg.get('exp_event_subtitle') or '世代共榮的數位聚合'
    cfg['date_start'] = cfg.get('date_start') or cfg.get('event_date_start') or cfg.get('exp_event_date_start') or ''
    cfg['date_end'] = cfg.get('date_end') or cfg.get('event_date_end') or cfg.get('exp_event_date_end') or ''
    cfg['brand_name'] = cfg.get('brand_name') or cfg.get('exp_brand_name') or '智匯方舟'
    cfg['products'] = load_products_from_config(cfg)
    cfg['product_categories'] = json_loads(cfg.get('product_categories'), ['課程', '諮詢', '聯盟', '紀念品'])
    cfg['industry_mappings'] = json_loads(cfg.get('industry_mappings'), [])
    cfg['agenda'] = json_loads(cfg.get('agenda'), [])
    if conn and admin and sheet:
        table_agenda = load_agenda(conn, admin, sheet)
        table_mapping = load_industry_mappings(conn, admin, sheet)
        table_exhibitors = load_exhibitors(conn, admin, sheet)
        if table_agenda:
            cfg['agenda'] = table_agenda
        if table_mapping:
            cfg['industry_mappings'] = table_mapping
        cfg['exhibitors'] = table_exhibitors
    for k in list(cfg.keys()):
        if isinstance(cfg[k], datetime):
            cfg[k] = cfg[k].strftime('%Y-%m-%d %H:%M:%S')
    return cfg


def get_logs(conn, admin, sheet, limit=None, checked_only=False):
    sql = """
        SELECT * FROM event_registrations
        WHERE admin_username=%s AND google_sheet_name=%s
    """
    args = [admin, sheet]
    if checked_only:
        sql += " AND status IN ('checked_in','已報到','替代','done')"
    sql += " ORDER BY COALESCE(checked_in_at, checkin_time, created_at) DESC, id DESC"
    if limit:
        sql += " LIMIT %s"
        args.append(int(limit))
    with conn.cursor() as cur:
        cur.execute(sql, args)
        return [public_user(r) for r in cur.fetchall()]


# ============================================================
# Pages
# ============================================================
@app.route('/')
def index():
    return send_from_directory('.', '活動報到系統.html')


@app.route('/admin')
def admin_page():
    return send_from_directory('.', 'admin.html')


@app.route('/dashboard')
def dashboard_page():
    return send_from_directory('.', 'dashboard.html')


@app.route('/<path:path>')
def static_files(path):
    return send_from_directory('.', path)


# ============================================================
# Auth / health
# ============================================================
@app.route('/api/health')
def health():
    return jsonify(success=True, status='ok', time=datetime.now().strftime('%Y-%m-%d %H:%M:%S'))

@app.route('/api/login', methods=['POST'])
def api_login():
    conn = None
    try:
        data = get_payload() or {}
        username = clean_text(data.get('username'))
        password = clean_text(data.get('password'))
        conn = get_db_connection()
        ensure_core_tables(conn)
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM admins WHERE TRIM(username)=%s AND TRIM(password)=%s LIMIT 1", (username, password))
            admin = cur.fetchone()
        if not admin:
            return jsonify(success=False, message='帳密錯誤'), 401
        allowed = [x.strip() for x in (admin.get('allowed_events') or DEFAULT_SHEET).split(',') if x.strip()] or [DEFAULT_SHEET]
        current = admin.get('current_event') if admin.get('current_event') in allowed else allowed[0]
        session['admin_logged_in'] = True
        session['username'] = admin['username']
        session['allowed_sheets'] = allowed
        session['current_admin_sheet'] = current
        return jsonify(success=True, username=admin['username'], allowed_sheets=allowed, sheets=allowed, current_sheet=current)
    except Exception as e:
        return jsonify(success=False, message=f'登入 API 失敗：{e}'), 500
    finally:
        if conn:
            conn.close()

@app.route('/api/logout')
def api_logout():
    session.clear()
    return jsonify(success=True)

# ============================================================
# Sheets / event selection
# ============================================================
@app.route('/api/current_sheet')
def current_sheet():
    admin = clean_text(request.args.get('admin')) or session.get('username') or DEFAULT_ADMIN
    conn = None
    try:
        conn = get_db_connection()
        ensure_core_tables(conn)
        with conn.cursor() as cur:
            cur.execute("SELECT current_event, allowed_events FROM admins WHERE username=%s LIMIT 1", (admin,))
            row = cur.fetchone() or {}
        allowed = [x.strip() for x in (row.get('allowed_events') or DEFAULT_SHEET).split(',') if x.strip()] or [DEFAULT_SHEET]
        current = row.get('current_event') or session.get('current_admin_sheet') or allowed[0]
        if current not in allowed:
            current = allowed[0]
        return jsonify(success=True, current_sheet=current, sheet=current, sheets=allowed)
    except Exception as e:
        fallback = session.get('current_admin_sheet') or DEFAULT_SHEET
        return jsonify(success=True, current_sheet=fallback, sheet=fallback, warning=str(e))
    finally:
        if conn:
            conn.close()


@app.route('/api/sheets/list')
def sheets_list():
    admin = clean_text(request.args.get('admin')) or session.get('username') or DEFAULT_ADMIN
    conn = None
    try:
        conn = get_db_connection()
        ensure_core_tables(conn)
        with conn.cursor() as cur:
            cur.execute("SELECT allowed_events, current_event FROM admins WHERE username=%s LIMIT 1", (admin,))
            row = cur.fetchone() or {}
            allowed = [x.strip() for x in (row.get('allowed_events') or DEFAULT_SHEET).split(',') if x.strip()] or [DEFAULT_SHEET]
            cur.execute("SELECT DISTINCT google_sheet_name FROM event_registrations WHERE admin_username=%s", (admin,))
            from_data = [r['google_sheet_name'] for r in cur.fetchall() if r.get('google_sheet_name')]
            cur.execute("SELECT DISTINCT google_sheet_name FROM event_configs WHERE admin_username=%s", (admin,))
            from_cfg = [r['google_sheet_name'] for r in cur.fetchall() if r.get('google_sheet_name')]
        sheets = []
        for s in allowed + from_cfg + from_data:
            if s and s not in sheets:
                sheets.append(s)
        current = row.get('current_event') or session.get('current_admin_sheet') or (sheets[0] if sheets else DEFAULT_SHEET)
        return jsonify(success=True, sheets=sheets or [DEFAULT_SHEET], allowed_sheets=sheets or [DEFAULT_SHEET], current_sheet=current)
    except Exception as e:
        return jsonify(success=False, message=str(e), sheets=[]), 500
    finally:
        if conn:
            conn.close()


@app.route('/api/session/sheet', methods=['POST'])
def session_sheet():
    conn = None
    try:
        data = get_payload() or {}
        admin = clean_text(data.get('admin')) or session.get('username') or DEFAULT_ADMIN
        sheet = clean_text(data.get('sheet') or data.get('google_sheet_name') or data.get('event_key')) or DEFAULT_SHEET
        conn = get_db_connection()
        ensure_core_tables(conn)
        ensure_config(conn, admin, sheet)
        with conn.cursor() as cur:
            cur.execute("SELECT allowed_events FROM admins WHERE username=%s LIMIT 1", (admin,))
            row = cur.fetchone()
            if row:
                allowed = [x.strip() for x in (row.get('allowed_events') or '').split(',') if x.strip()]
                if sheet not in allowed:
                    allowed.append(sheet)
                cur.execute("UPDATE admins SET current_event=%s, allowed_events=%s WHERE username=%s", (sheet, ','.join(allowed), admin))
            else:
                cur.execute("INSERT INTO admins (username,password,allowed_events,current_event) VALUES (%s,%s,%s,%s)", (admin, DEFAULT_PASSWORD, sheet, sheet))
        conn.commit()
        session['username'] = admin
        session['current_admin_sheet'] = sheet
        return jsonify(success=True, current_sheet=sheet, sheet=sheet)
    except Exception as e:
        if conn:
            conn.rollback()
        return jsonify(success=False, message=str(e)), 500
    finally:
        if conn:
            conn.close()

# ============================================================
# Config / Navigator
# ============================================================
@app.route('/api/config', methods=['GET', 'PUT', 'POST'])
def api_config():
    admin, sheet = q_event()
    conn = None
    try:
        conn = get_db_connection()
        ensure_core_tables(conn)
        if request.method in ['PUT', 'POST']:
            save_config(conn, admin, sheet, get_payload() or {})
        row = get_config_row(conn, admin, sheet)
        return jsonify(serialize_config(row, conn, admin, sheet))
    except Exception as e:
        return jsonify(success=False, message=str(e)), 500
    finally:
        if conn:
            conn.close()


@app.route('/api/event-config', methods=['GET'])
def api_event_config():
    admin, sheet = q_event()
    conn = None
    try:
        conn = get_db_connection()
        ensure_core_tables(conn)
        row = get_config_row(conn, admin, sheet)
        return jsonify(success=True, config=serialize_config(row, conn, admin, sheet))
    except Exception as e:
        return jsonify(success=False, message=str(e), config={}), 500
    finally:
        if conn:
            conn.close()


@app.route('/api/admin/event-config', methods=['PUT', 'POST'])
@app.route('/api/admin/config', methods=['PUT', 'POST'])
def api_admin_event_config():
    admin, sheet = q_event()
    conn = None
    try:
        conn = get_db_connection()
        ensure_core_tables(conn)
        save_config(conn, admin, sheet, get_payload() or {})
        row = get_config_row(conn, admin, sheet)
        return jsonify(success=True, message='設定已儲存', config=serialize_config(row, conn, admin, sheet))
    except Exception as e:
        if conn:
            conn.rollback()
        return jsonify(success=False, message=str(e)), 500
    finally:
        if conn:
            conn.close()


# ============================================================
# Products
# ============================================================
@app.route('/api/products', methods=['GET', 'PUT', 'POST'])
def api_products():
    admin, sheet = q_event()
    conn = None
    try:
        conn = get_db_connection()
        ensure_core_tables(conn)
        if request.method in ['PUT', 'POST']:
            data = get_payload() or {}
            products = data.get('products') if isinstance(data, dict) else []
            if products is None and isinstance(data, list):
                products = data
            save_config(conn, admin, sheet, {'products': products or []})
        cfg = get_config_row(conn, admin, sheet)
        return jsonify(success=True, products=load_products_from_config(cfg))
    except Exception as e:
        return jsonify(success=False, message=str(e), products=[]), 500
    finally:
        if conn:
            conn.close()

# ============================================================
# Agenda / schedule
# ============================================================
@app.route('/api/agenda', methods=['GET', 'PUT', 'POST'])
@app.route('/api/schedule', methods=['GET', 'PUT', 'POST'])
def api_agenda():
    admin, sheet = q_event()
    conn = None
    try:
        conn = get_db_connection()
        ensure_core_tables(conn)
        if request.method in ['PUT', 'POST']:
            data = get_payload() or {}
            rows = data.get('agenda') or data.get('schedule') or data.get('items') or (data if isinstance(data, list) else [])
            save_agenda(conn, admin, sheet, rows)
        rows = load_agenda(conn, admin, sheet)
        return jsonify(success=True, agenda=rows, schedule=rows, items=rows)
    except Exception as e:
        if conn:
            conn.rollback()
        return jsonify(success=False, message=str(e), agenda=[], schedule=[]), 500
    finally:
        if conn:
            conn.close()


@app.route('/api/admin/schedule', methods=['PUT', 'POST'])
def api_admin_schedule():
    return api_agenda()


# ============================================================
# Industry mappings / exhibitors / companies
# ============================================================
@app.route('/api/industry-mappings', methods=['GET', 'PUT', 'POST'])
def api_industry_mappings():
    admin, sheet = q_event()
    conn = None
    try:
        conn = get_db_connection()
        ensure_core_tables(conn)
        if request.method in ['PUT', 'POST']:
            data = get_payload() or {}
            rows = data.get('mappings') or data.get('industry_mappings') or data.get('items') or (data if isinstance(data, list) else [])
            save_industry_mappings(conn, admin, sheet, rows)
        rows = load_industry_mappings(conn, admin, sheet)
        return jsonify(success=True, mappings=rows, industry_mappings=rows)
    except Exception as e:
        if conn:
            conn.rollback()
        return jsonify(success=False, message=str(e), mappings=[]), 500
    finally:
        if conn:
            conn.close()


@app.route('/api/exhibitors', methods=['GET', 'PUT', 'POST'])
def api_exhibitors():
    admin, sheet = q_event()
    conn = None
    try:
        conn = get_db_connection()
        ensure_core_tables(conn)
        if request.method in ['PUT', 'POST']:
            data = get_payload() or {}
            rows = data.get('exhibitors') or data.get('items') or (data if isinstance(data, list) else [])
            save_exhibitors(conn, admin, sheet, rows)
        rows = load_exhibitors(conn, admin, sheet)
        return jsonify(success=True, exhibitors=rows, data=rows)
    except Exception as e:
        if conn:
            conn.rollback()
        return jsonify(success=False, message=str(e)), 500
    finally:
        if conn:
            conn.close()


@app.route('/api/companies')
def api_companies():
    admin, sheet = q_event()
    conn = None
    try:
        conn = get_db_connection()
        ensure_core_tables(conn)
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT DISTINCT COALESCE(NULLIF(company,''), company_name) AS company
                FROM event_registrations
                WHERE admin_username=%s AND google_sheet_name=%s
                  AND COALESCE(NULLIF(company,''), company_name, '') <> ''
                ORDER BY company
                """,
                (admin, sheet),
            )
            rows = [r['company'] for r in cur.fetchall() if r.get('company')]
        return jsonify(success=True, companies=rows)
    except Exception as e:
        return jsonify(success=False, message=str(e), companies=[]), 500
    finally:
        if conn:
            conn.close()

# ============================================================
# Registrations / search / check-in
# ============================================================
@app.route('/api/search/<method>')
def api_search(method):
    admin, sheet = q_event()
    method = clean_text(method).lower()
    value = clean_text(
        request.args.get(method)
        or request.args.get('q')
        or request.args.get('value')
        or request.args.get('keyword')
        or request.args.get('query')
        or request.args.get('company')
        or request.args.get('name')
        or request.args.get('phone')
    )
    conn = None
    try:
        conn = get_db_connection()
        ensure_core_tables(conn)
        if not value:
            return jsonify(success=True, data=[], results=[])

        if method in ['company', 'company_name', 'unit']:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT * FROM event_registrations
                    WHERE admin_username=%s AND google_sheet_name=%s
                    ORDER BY id ASC
                    LIMIT 5000
                    """,
                    (admin, sheet),
                )
                candidates = cur.fetchall()
            scored = []
            for row in candidates:
                score = search_match_score(row, value, method)
                if score > 0:
                    scored.append((score, row))
            scored.sort(key=lambda x: (-x[0], clean_text(x[1].get('company') or x[1].get('company_name')), clean_text(x[1].get('name'))))
            rows = [public_user(r) for _, r in scored[:120]]
            return jsonify(success=True, data=rows, results=rows, count=len(rows))

        like = f"%{value}%"
        if method in ['phone', 'mobile', 'tel']:
            cond = "phone LIKE %s"
            args = [like]
        else:
            cond = "name LIKE %s"
            args = [like]
        with conn.cursor() as cur:
            cur.execute(
                f"""
                SELECT * FROM event_registrations
                WHERE admin_username=%s AND google_sheet_name=%s AND {cond}
                ORDER BY id ASC
                LIMIT 120
                """,
                [admin, sheet] + args,
            )
            rows = [public_user(r) for r in cur.fetchall()]
        return jsonify(success=True, data=rows, results=rows, count=len(rows))
    except Exception as e:
        return jsonify(success=False, message=str(e), data=[], results=[]), 500
    finally:
        if conn:
            conn.close()


@app.route('/api/registrations/add', methods=['POST'])
def api_registration_add():
    conn = None
    try:
        data = get_payload() or {}
        admin = clean_text(data.get('admin')) or session.get('username') or DEFAULT_ADMIN
        sheet = clean_text(data.get('sheet') or data.get('google_sheet_name')) or session.get('current_admin_sheet') or DEFAULT_SHEET
        portrait_bool = bool(data.get('portrait_consent'))
        portrait_status = clean_text(data.get('portrait_consent_status')) or ('同意' if portrait_bool else '不同意')
        now = datetime.now()
        conn = get_db_connection()
        ensure_core_tables(conn)
        ensure_config(conn, admin, sheet)
        with conn.cursor() as cur:
            # 具名排除 id，全面精準寫入
            cur.execute(
                """
                INSERT INTO event_registrations
                (admin_username, google_sheet_name, admin_user, event_key, name, phone, email, company, company_name, job_title, seat, seating_chart, status, is_original, checked_in_at, checkin_time, portrait_consent, portrait_consent_status, portrait_consent_time, special_notes, note, raw_data)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                """,
                (
                    admin, sheet, admin, sheet,
                    clean_text(data.get('name')),
                    clean_text(data.get('phone')),
                    clean_text(data.get('email')),
                    clean_text(data.get('company')),
                    clean_text(data.get('company')),
                    clean_text(data.get('job_title')),
                    clean_text(data.get('seat')) or '現場安排',
                    clean_text(data.get('seat')) or '現場安排',
                    'checked_in', 1, now, now,
                    1 if portrait_bool else 0,
                    portrait_status,
                    now,
                    clean_text(data.get('special_notes') or data.get('note')),
                    clean_text(data.get('special_notes') or data.get('note')),
                    json_dumps({k: v for k, v in data.items()}),
                ),
            )
            rid = cur.lastrowid
            cur.execute("SELECT * FROM event_registrations WHERE id=%s", (rid,))
            row = public_user(cur.fetchone())
        conn.commit()
        return jsonify(success=True, id=rid, data=row)
    except Exception as e:
        if conn:
            conn.rollback()
        return jsonify(success=False, message=str(e)), 500
    finally:
        if conn:
            conn.close()


@app.route('/api/checkin/<int:rid>', methods=['POST'])
def api_checkin(rid):
    conn = None
    try:
        data = get_payload() or {}
        admin, sheet = event_args()
        is_original = data.get('is_original', True)
        proxy = data.get('proxy_info') or data.get('proxy') or {}
        portrait_bool = bool(data.get('portrait_consent'))
        portrait_status = clean_text(data.get('portrait_consent_status') or data.get('image_rights_status')) or ('同意' if portrait_bool else '不同意')
        status = 'checked_in' if is_original else '替代'
        now = datetime.now()
        conn = get_db_connection()
        ensure_core_tables(conn)
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM event_registrations WHERE id=%s AND admin_username=%s AND google_sheet_name=%s LIMIT 1", (rid, admin, sheet))
            old = cur.fetchone()
            if not old:
                cur.execute("SELECT * FROM event_registrations WHERE id=%s LIMIT 1", (rid,))
                old = cur.fetchone()
            if not old:
                return jsonify(success=False, message='找不到報到資料'), 404
            cur.execute(
                """
                UPDATE event_registrations
                SET status=%s,
                    is_original=%s,
                    proxy_name=%s,
                    proxy_phone=%s,
                    checked_in_at=%s,
                    checkin_time=%s,
                    portrait_consent=%s,
                    portrait_consent_status=%s,
                    portrait_consent_time=%s
                WHERE id=%s
                """,
                (
                    status,
                    1 if is_original else 0,
                    clean_text(proxy.get('name')) if not is_original else '',
                    clean_text(proxy.get('phone')) if not is_original else '',
                    now,
                    now,
                    1 if portrait_bool else 0,
                    portrait_status,
                    now,
                    rid,
                ),
            )
            cur.execute("SELECT * FROM event_registrations WHERE id=%s", (rid,))
            updated = public_user(cur.fetchone())
        conn.commit()
        return jsonify(success=True, data=updated)
    except Exception as e:
        if conn:
            conn.rollback()
        return jsonify(success=False, message=str(e)), 500
    finally:
        if conn:
            conn.close()


@app.route('/api/company-search')
def api_company_search_alias():
    return api_search('company')


@app.route('/api/search')
def api_search_query_alias():
    method = clean_text(request.args.get('method') or request.args.get('type') or 'name')
    return api_search(method)


# ============================================================
# CSV import / export / delete
# ============================================================
@app.route('/api/sheets/import_csv', methods=['POST'])
def import_csv_api():
    admin, sheet = q_event()
    conn = None
    try:
        upload = request.files.get('file') or request.files.get('csv') or request.files.get('upload')
        if not upload:
            return jsonify(success=False, message='找不到 CSV 檔案'), 400
        raw = upload.read()
        text = None
        last_error = None
        for enc_name in ['utf-8-sig', 'utf-8', 'cp950', 'big5', 'big5hkscs', 'latin1']:
            try:
                text = raw.decode(enc_name)
                break
            except Exception as err:
                last_error = err
        if text is None:
            return jsonify(success=False, message=f'CSV 編碼解析失敗：{last_error}'), 400
        sample = text[:4096]
        try:
            dialect = csv.Sniffer().sniff(sample)
        except Exception:
            dialect = csv.excel
        reader = csv.DictReader(io.StringIO(text), dialect=dialect)
        if not reader.fieldnames:
            return jsonify(success=False, message='CSV 沒有標題列，請確認第一列是欄位名稱'), 400
        rows = [normalize_registration(r) for r in reader]
        conn = get_db_connection()
        ensure_core_tables(conn)
        ensure_config(conn, admin, sheet)
        with conn.cursor() as cur:
            cur.execute("DELETE FROM event_registrations WHERE admin_username=%s AND google_sheet_name=%s", (admin, sheet))
            for r in rows:
                if not (r['name'] or r['phone'] or r['company']):
                    continue
                # 完全避免在具名 INSERT 欄位中放入 id 欄位名，徹底解決 1054 錯誤
                cur.execute(
                    """
                    INSERT INTO event_registrations
                    (admin_username, google_sheet_name, admin_user, event_key, name, phone, email, company, company_name, job_title, region, training_level, seat, seating_chart, status, special_notes, note, raw_data)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,'pending',%s,%s,%s)
                    """,
                    (
                        admin, sheet, admin, sheet,
                        r['name'], r['phone'], r['email'], r['company'], r['company'], r['job_title'],
                        r['region'], r['training_level'], r['seat'], r['seat'], r['special_notes'], r['special_notes'], r['raw_data'],
                    ),
                )
        conn.commit()
        return jsonify(success=True, message=f'CSV 匯入完成：{len(rows)} 筆', count=len(rows))
    except Exception as e:
        if conn:
            conn.rollback()
        return jsonify(success=False, message=str(e)), 500
    finally:
        if conn:
            conn.close()


@app.route('/api/sheets/delete_data', methods=['DELETE', 'POST'])
def delete_sheet_data_api():
    admin, sheet = q_event()
    conn = None
    try:
        conn = get_db_connection()
        ensure_core_tables(conn)
        with conn.cursor() as cur:
            cur.execute("DELETE FROM event_registrations WHERE admin_username=%s AND google_sheet_name=%s", (admin, sheet))
        conn.commit()
        return jsonify(success=True, message=f'已成功清空場次「{sheet}」的所有旅客名單資料')
    except Exception as e:
        if conn:
            conn.rollback()
        return jsonify(success=False, message=str(e)), 500
    finally:
        if conn:
            conn.close()


@app.route('/api/sheets/export_csv', methods=['GET'])
def export_csv_api():
    admin, sheet = q_event()
    conn = None
    try:
        conn = get_db_connection()
        ensure_core_tables(conn)
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT * FROM event_registrations
                WHERE admin_username=%s AND google_sheet_name=%s
                ORDER BY id ASC
                """,
                (admin, sheet),
            )
            rows = [public_user(r) for r in cur.fetchall()]
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow(['表單名稱', sheet])
        writer.writerow(['產生時間', datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
        writer.writerow([])
        writer.writerow(['姓名', '手機', '公司/單位', 'Email', '職稱', '桌號/座位', '報到狀態', '報到時間', '肖像權狀態', '代理人姓名', '代理人手機', '備註'])
        for r in rows:
            writer.writerow([
                r.get('name', ''), r.get('phone', ''), r.get('company', ''), r.get('email', ''), r.get('job_title', ''),
                r.get('seat', ''), r.get('status', ''), r.get('checkin_time', ''), r.get('portrait_consent_status', ''),
                r.get('proxy_name', ''), r.get('proxy_phone', ''), r.get('special_notes', ''),
            ])
        data = output.getvalue().encode('utf-8-sig')
        filename = re.sub(r'[\\/:*?"<>|\s]+', '_', sheet).strip('_') or 'registrations'
        return Response(
            data,
            mimetype='text/csv; charset=utf-8',
            headers={'Content-Disposition': f"attachment; filename*=UTF-8''{quote(filename)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"},
        )
    except Exception as e:
        return jsonify(success=False, message=str(e)), 500
    finally:
        if conn:
            conn.close()

# ============================================================
# Stats
# ============================================================
@app.route('/api/dashboard_stats')
def dashboard_stats():
    admin, sheet = q_event()
    conn = None
    try:
        conn = get_db_connection()
        ensure_core_tables(conn)
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT
                    COUNT(*) AS total,
                    SUM(CASE WHEN status IN ('checked_in','已報到','替代','done') THEN 1 ELSE 0 END) AS checked
                FROM event_registrations
                WHERE admin_username=%s AND google_sheet_name=%s
                """,
                (admin, sheet),
            )
            s = cur.fetchone() or {}
            total = int(s.get('total') or 0)
            checked = int(s.get('checked') or 0)

            cur.execute(
                """
                SELECT id, name, phone, email, company, company_name, job_title,
                       COALESCE(NULLIF(seat,''), seating_chart, '未分桌') AS seat,
                       status, checked_in_at, checkin_time, portrait_consent_status, special_notes
                FROM event_registrations
                WHERE admin_username=%s AND google_sheet_name=%s
                ORDER BY id ASC
                """,
                (admin, sheet),
            )
            table_detail_map = {}
            for raw_row in cur.fetchall():
                row = public_user(raw_row)
                seat_val = clean_text(row.get('seat') or '未分桌')
                if not seat_val:
                    seat_val = '未分桌'
                
                table_val = seat_val.split('-')[0].strip() if '-' in seat_val else seat_val
                table_detail_map.setdefault(table_val, []).append(row)

            table_stats = []
            table_details = []
            for table_val, members in sorted(table_detail_map.items(), key=lambda kv: table_sort_key(kv[0])):
                checked_count = sum(1 for m in members if status_checked(m.get('status')))
                total_count = len(members)
                table_stats.append({
                    'table': table_val,
                    'seat': table_val,
                    'total': total_count,
                    'checked_in': checked_count,
                    'percent': round((checked_count / total_count) * 100, 1) if total_count else 0
                })
                table_details.append({
                    'table': table_val,
                    'seat': table_val,
                    'members': members,
                    'checked_in': checked_count,
                    'total': total_count
                })

        logs = get_logs(conn, admin, sheet, 25, checked_only=True)
        industry_logs = get_logs(conn, admin, sheet, None, checked_only=True)
        return jsonify(success=True, stats={
            'total': total,
            'checked_in': checked,
            'not_checked_in': max(total - checked, 0),
            'logs': logs,
            'industry_logs': [{'company': r.get('company', ''), 'name': r.get('name', '')} for r in industry_logs],
            'table_stats': table_stats,
            'table_details': table_details,
        })
    except Exception as e:
        return jsonify(success=False, message=str(e), stats={}), 500
    finally:
        if conn:
            conn.close()

if __name__ == '__main__':
    port = int(os.getenv('PORT', 10000))
    app.run(host='0.0.0.0', port=port, debug=os.getenv('FLASK_DEBUG') == '1')
