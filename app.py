from flask import Flask, render_template, request, redirect, url_for, flash, session, g, jsonify
import os, json, csv, requests
import psycopg as psycopg2
import fitz #PyMuPDF
import openpyxl
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash
from functools import wraps
from pathlib import Path
from dotenv import load_dotenv
from celery import Celery

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', '2409')

# ========== RENDER POSTGRES ==========


DATABASE_URL = os.getenv('DATABASE_URL')

def get_db():
    db = getattr(g, '_database', None)
    if db is None:
        db = g._database = psycopg2.connect(DATABASE_URL, cursor_factory=psycopg2.extras.DictCursor)
    return db

@app.teardown_appcontext
def close_connection(exception):
    db = getattr(g, '_database', None)
    if db is not None:
        db.close()

def init_db():
    db = psycopg2.connect(DATABASE_URL)
    cur = db.cursor()
    cur.execute('''
    CREATE TABLE IF NOT EXISTS users (
        id SERIAL PRIMARY KEY,
        email TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        school_name TEXT NOT NULL,
        sender_id TEXT NOT NULL,
        paycode TEXT,
        sms_credits INTEGER DEFAULT 1000,
        plan TEXT DEFAULT 'pro',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS messages (
        id SERIAL PRIMARY KEY,
        user_id INTEGER NOT NULL,
        student_name TEXT,
        phone TEXT,
        message TEXT,
        msg_id TEXT,
        status TEXT DEFAULT 'queued',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (user_id) REFERENCES users (id)
    );
    CREATE TABLE IF NOT EXISTS tickets (
        id SERIAL PRIMARY KEY,
        user_id INTEGER NOT NULL,
        subject TEXT,
        body TEXT,
        status TEXT DEFAULT 'open',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (user_id) REFERENCES users (id)
    );
    ''')
    db.commit()
    db.close()

# ========== CELERY ==========
celery = Celery(app.name, broker=os.getenv('REDIS_URL'), backend=os.getenv('REDIS_URL'))

@celery.task(bind=True, max_retries=3)
def send_sms_task(self, user_id, name, phone, balance, due, sender_id, paycode):
    db = psycopg2.connect(DATABASE_URL)
    cur = db.cursor()

    # Format phone: 0772... -> 263772...
    phone = ''.join(filter(str.isdigit, str(phone)))
    if phone.startswith('0'):
        phone = '263' + phone[1:]

    msg = f"{sender_id}: {name} fees ${balance} due {due}. Pay {paycode}. Ignore if paid."

    url = "https://api.ping.co.zw/v1/sms/send"
    headers = {"Authorization": f"Bearer {os.getenv('PING_API_KEY')}"}
    payload = {"to": phone, "message": msg, "from": sender_id}

    try:
        r = requests.post(url, headers=headers, json=payload, timeout=10)
        resp = r.json()
        success = r.status_code == 200 and resp.get('status') == 'success'
        msg_id = resp.get('messageId', '')
        status = 'sent' if success else 'failed'
    except Exception as e:
        msg_id, status = '', 'failed'
        if self.request.retries < self.max_retries:
            raise self.retry(exc=e, countdown=60)

    cur.execute(
        'INSERT INTO messages (user_id, student_name, phone, message, msg_id, status) VALUES (%s,%s,%s,%s)',
        (user_id, name, phone, msg, msg_id, status)
    )
    if status == 'sent':
        cur.execute('UPDATE users SET sms_credits = sms_credits - 1 WHERE id = %s', (user_id,))
    db.commit()
    db.close()
    return status

# ========== FILE PARSING ==========
def parse_pdf(filepath):
    doc = fitz.open(filepath)
    rows, headers = [], []
    for page in doc:
        tables = page.find_tables()
        if tables.tables:
            data = tables.tables[0].extract()
            if not data:
                continue
            headers = [str(h or '').lower().strip() for h in data[0]]
            for row in data[1:]:
                if any(cell and str(cell).strip() for cell in row):
                    clean_row = [str(c or '').strip() for c in row]
                    rows.append(dict(zip(headers, clean_row)))
            break
    doc.close()
    if not headers:
        raise ValueError("No table found in PDF. Export as Excel if this fails.")
    return rows, headers

def parse_excel(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    rows = []
    headers = [str(cell.value or '').lower().strip() for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    for row in ws.iter_rows(min_row=2):
        values = [str(cell.value or '').strip() for cell in row]
        if any(values):
            rows.append(dict(zip(headers, values)))
    return rows, headers

def parse_csv(filepath):
    with open(filepath, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        headers = [h.lower().strip() for h in reader.fieldnames]
        rows = [{k.lower().strip(): str(v or '').strip() for k, v in row.items()} for row in reader]
    return rows, headers

def parse_any_file(filepath):
    ext = Path(filepath).suffix.lower()
    if ext == '.pdf':
        return parse_pdf(filepath)
    elif ext in ['.xlsx', '.xls']:
        return parse_excel(filepath)
    elif ext == '.csv':
        return parse_csv(filepath)
    else:
        raise ValueError(f"Unsupported file type: {ext}. Use PDF, Excel, or CSV.")

def auto_find_cols(headers):
    name_keys = ['name', 'student', 'learner', 'pupil']
    phone_keys = ['phone', 'mobile', 'cell', 'contact', 'number']
    bal_keys = ['balance', 'due', 'owing', 'amount', 'fee', 'outstanding']

    name_c = next((h for h in headers if any(k in h for k in name_keys)), None)
    phone_c = next((h for h in headers if any(k in h for k in phone_keys)), None)
    bal_c = next((h for h in headers if any(k in h for k in bal_keys)), None)

    if not all([name_c, phone_c, bal_c]):
        raise ValueError(f"Missing columns. Found: {headers}. Need name, phone, balance columns.")
    return name_c, phone_c, bal_c

# ========== AUTH ==========
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

def get_current_user():
    if 'user_id' not in session:
        return None
    db = get_db()
    cur = db.cursor()
    cur.execute('SELECT * FROM users WHERE id = %s', (session['user_id'],))
    return cur.fetchone()

@app.context_processor
def inject_user():
    return dict(current_user=get_current_user())

# ========== PING WEBHOOK ==========
@app.route('/webhook', methods=['POST'])
def ping_webhook():
    data = request.get_json()
    msg_id = data.get('messageId')
    status = data.get('status')
    if msg_id:
        db = get_db()
        cur = db.cursor()
        cur.execute("UPDATE messages SET status=%s WHERE msg_id=%s", (status, msg_id))
        db.commit()
    return '', 200

# ========== PWA ROUTES ==========
@app.route('/manifest.json')
def manifest():
    return jsonify({
        "name": "FeeRemind Pro",
        "short_name": "FeeRemind",
        "start_url": "/dashboard",
        "display": "standalone",
        "background_color": "#0f172a",
        "theme_color": "#2563eb",
        "icons": [{"src": "/static/icon-192.png", "sizes": "192x192"}]
    })

@app.route('/sw.js')
def sw():
    return app.send_static_file('sw.js')

# ========== AUTH ROUTES ==========
@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        email = request.form['email']
        password = request.form['password']
        school_name = request.form['school_name']
        sender_id = request.form['sender_id']
        paycode = request.form['paycode']

        db = get_db()
        cur = db.cursor()
        try:
            cur.execute(
                'INSERT INTO users (email, password_hash, school_name, sender_id, paycode) VALUES (%s,%s,%s,%s)',
                (email, generate_password_hash(password), school_name, sender_id, paycode)
            )
            db.commit()
            flash('Account created. Login now.', 'success')
            return redirect(url_for('login'))
        except psycopg2.IntegrityError:
            flash('Email already exists', 'error')
    return render_template('register.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form['email']
        password = request.form['password']
        db = get_db()
        cur = db.cursor()
        cur.execute('SELECT * FROM users WHERE email = %s', (email,))
        user = cur.fetchone()
        if user and check_password_hash(user['password_hash'], password):
            session['user_id'] = user['id']
            return redirect(url_for('dashboard'))
        flash('Invalid credentials', 'error')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# ========== APP ROUTES ==========
@app.route('/')
def index():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route('/dashboard')
@login_required
def dashboard():
    user = get_current_user()
    db = get_db()
    cur = db.cursor()
    cur.execute('''
        SELECT COUNT(*) as total,
               SUM(CASE WHEN status='Delivered' THEN 1 ELSE 0 END) as delivered,
               SUM(CASE WHEN status='Failed' THEN 1 ELSE 0 END) as failed,
               SUM(CASE WHEN status='sent' THEN 1 ELSE 0 END) as sent
        FROM messages WHERE user_id = %s AND date(created_at) = CURRENT_DATE
    ''', (user['id'],))
    stats = cur.fetchone()

    cur.execute('''
        SELECT date(created_at) as date, COUNT(*) as count,
               SUM(CASE WHEN status='Delivered' THEN 1 ELSE 0 END) as delivered
        FROM messages WHERE user_id = %s
        GROUP BY date(created_at) ORDER BY date DESC LIMIT 7
    ''', (user['id'],))
    uploads = cur.fetchall()

    return render_template('dashboard.html', user=user, stats=stats, uploads=uploads, school_name=user['school_name'])

@app.route('/upload', methods=['GET', 'POST'])
@login_required
def upload():
    user = get_current_user()
    if user['sms_credits'] <= 0:
        flash('SMS credits exhausted. Contact support.', 'error')
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        file = request.files['file']
        if not file.filename:
            flash('No file selected', 'error')
            return redirect(url_for('upload'))

        path = f"/tmp/{file.filename}"
        file.save(path)

        try:
            rows, headers = parse_any_file(path)
            name_c, phone_c, bal_c = auto_find_cols(headers)
        except Exception as e:
            flash(f'File error: {str(e)}', 'error')
            return redirect(url_for('upload'))

        queued = 0
        for r in rows:
            try:
                bal = float(str(r.get(bal_c, '0')).replace('$','').replace(',','') or 0)
            except:
                continue

            if bal > 0 and user['sms_credits'] > queued:
                phone = str(r.get(phone_c, ''))
                name = str(r.get(name_c, 'Parent'))
                due = str(r.get('due', r.get('due_date', r.get('duedate', 'ASAP'))))

                send_sms_task.delay(user['id'], name, phone, bal, due, user['sender_id'], user['paycode'])
                queued += 1

        os.remove(path)
        flash(f'Queued {queued} SMS. Sending in background.', 'success')
        return redirect(url_for('dashboard'))

    return render_template('upload.html', user=user, school_name=user['school_name'])

@app.route('/settings', methods=['GET', 'POST'])
@login_required
def settings():
    user = get_current_user()
    if request.method == 'POST':
        db = get_db()
        cur = db.cursor()
        cur.execute(
            'UPDATE users SET school_name=%s, sender_id=%s, paycode=%s WHERE id=%s',
            (request.form['school_name'], request.form['sender_id'], request.form['paycode'], user['id'])
        )
        db.commit()
        flash('Settings saved', 'success')
        return redirect(url_for('settings'))

    return render_template('settings.html', user=user, school_name=user['school_name'])

@app.route('/api/status')
@login_required
def api_status():
    user = get_current_user()
    db = get_db()
    cur = db.cursor()
    cur.execute('''
        SELECT status, COUNT(*) FROM messages
        WHERE user_id=%s AND date(created_at)=CURRENT_DATE GROUP BY status
    ''', (user['id'],))
    data = {row[0]: row[1] for row in cur.fetchall()}
    return jsonify(data)

@app.route('/help', methods=['GET', 'POST'])
@login_required
def help():
    user = get_current_user()
    if request.method == 'POST':
        db = get_db()
        cur = db.cursor()
        cur.execute('INSERT INTO tickets (user_id, subject, body) VALUES (%s,%s)',
                   (user['id'], request.form['subject'], request.form['body']))
        db.commit()
        flash('Ticket submitted. We reply within 4 hours.', 'success')
    return render_template('help.html', user=user, school_name=user['school_name'])

if __name__ == '__main__':
    init_db()
    app.run(debug=False, host='0.0.0.0')