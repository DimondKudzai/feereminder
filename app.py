import os, csv, time
from pathlib import Path
from datetime import datetime
from functools import wraps
from flask_admin import Admin
from flask_admin.contrib.sqla import ModelView
from flask_login import LoginManager
import fitz
import openpyxl
import requests
from flask import Flask, render_template, request, redirect, url_for, flash, session, jsonify
from sqlalchemy import func, case
from werkzeug.security import generate_password_hash, check_password_hash
from flask_sqlalchemy import SQLAlchemy
from dotenv import load_dotenv


load_dotenv()

# ====================== APP & DB SETUP ======================


app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', '2409')
app.config['SQLALCHEMY_DATABASE_URI'] = os.getenv('DATABASE_URL')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
# Optional but recommended for Render/Neon
app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
    "pool_pre_ping": True,       # checks connections are alive
    "pool_recycle": 280,         # recycle idle connections before timeout
    "connect_args": {"sslmode": "require"}  # ensures SSL
}


db = SQLAlchemy(app)

# ====================== MODELS ======================
class User(db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.Text, unique=True, nullable=False)
    password_hash = db.Column(db.Text, nullable=False)
    school_name = db.Column(db.Text, nullable=False)
    sender_id = db.Column(db.Text, nullable=False)
    message_theme = db.Column(db.Text, default='Tution')
    paycode = db.Column(db.Text)
    sms_credits = db.Column(db.Integer, default=500)
    plan = db.Column(db.Text, default='pro')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    messages = db.relationship('Message', backref='user', lazy=True)
    tickets = db.relationship('Ticket', backref='user', lazy=True)

class Message(db.Model):
    __tablename__ = 'messages'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    student_name = db.Column(db.Text)
    phone = db.Column(db.Text)
    message = db.Column(db.Text)
    msg_id = db.Column(db.Text)
    status = db.Column(db.Text, default='queued')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class Ticket(db.Model):
    __tablename__ = 'tickets'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    subject = db.Column(db.Text)
    body = db.Column(db.Text)
    status = db.Column(db.Text, default='open')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class SmsBatch(db.Model):
    __tablename__ = 'sms_batches'

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'))
    batch_number = db.Column(db.Text, unique=True)
    batch_id = db.Column(db.Text)
    status = db.Column(db.Text, default="PROCESSING")
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

with app.app_context():
    db.create_all()
    
    
    
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'  # optional but recommended

    
@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))
    
class MyModelView(ModelView):
    def is_accessible(self):
        return current_user.is_authenticated and current_user.id == 1 or current_user.id == 2

admin = Admin(app, name='Fee Reminder')
admin.add_view(MyModelView(User, db.session))
admin.add_view(MyModelView(Message, db.session))
admin.add_view(MyModelView(Ticket, db.session))
    
    
# ====================== SMS SENDER ======================
#Esolutions
def send_sms_sync(user_id, recipients):

    import uuid

    user = User.query.get(user_id)

    esol_user = os.getenv('ESOL_API_USER')
    esol_pass = os.getenv('ESOL_API_PASS')

    if not esol_user or not esol_pass:
        raise ValueError("Missing eSolutions credentials")


    bulk_url = "https://mobile.esolutions.co.zw/bmg/api/bulk"


    batch_number = "B" + str(uuid.uuid4().hex[:12]).upper()


    messages = []


    for index, recipient in enumerate(recipients):

        phone = ''.join(filter(str.isdigit, str(recipient['phone'])))


        if phone.startswith('0'):
            phone = '263' + phone[1:]

        elif not phone.startswith('263'):
            phone = '263' + phone


        symbol = "USD $" if recipient['currency']=="USD" else "ZWG $"


        msg = (
            f"{user.school_name}: Fees reminder for "
            f"{recipient['name']}. "
            f"Balance: {symbol}{recipient['balance']:.2f}. "
            f"Query Admin."
        )


        reference = f"{user_id}-{uuid.uuid4().hex[:8]}"


        messages.append({

            "originator": user.sender_id,

            "destination": phone,

            "messageText": msg,

            "messageReference": reference

        })


        db.session.add(
            Message(
                user_id=user_id,
                student_name=recipient['name'],
                phone=phone,
                message=msg,
                msg_id=reference,
                status="PROCESSING"
            )
        )


    payload = {

        "batchNumber": batch_number,

        "messages": messages

    }


    try:

        r = requests.post(

            bulk_url,

            auth=(esol_user, esol_pass),

            headers={
                "Content-Type":"application/json"
            },

            json=payload,

            timeout=30

        )


        print("ESOL RESPONSE:")
        print(r.status_code)
        print(r.text)


        response = r.json()


        batch = SmsBatch(

            user_id=user_id,

            batch_number=batch_number,

            batch_id=response.get("batchId"),

            status=response.get("status","UNKNOWN")

        )


        db.session.add(batch)


        db.session.commit()


        return {

            "batch": batch_number,

            "status": response.get("status"),

            "total": len(messages)

        }



    except Exception as e:

        db.session.rollback()

        print("SMS ERROR:", e)

        raise
        
        
# ============ check batch ====================

@app.route('/check_batch/<batch_number>')
@login_required
def check_batch(batch_number):

    esol_user = os.getenv('ESOL_API_USER')
    esol_pass = os.getenv('ESOL_API_PASS')


    url = f"https://mobile.esolutions.co.zw/bmg/api/bulk/{batch_number}"


    r = requests.get(
        url,
        auth=(esol_user, esol_pass),
        timeout=20
    )


    data = r.json()


    batch = SmsBatch.query.filter_by(
        batch_number=batch_number
    ).first()


    if batch:

        batch.status = data.get("status")

        db.session.commit()


    return jsonify(data)
        
        

 # ====================== FILE PARSING ======================
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
    headers = [str(cell.value or '').lower().strip() for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    rows = []
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

# ====================== AUTH DECORATORS ======================
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

@app.context_processor
def inject_user():
    if 'user_id' in session:
        return dict(current_user=User.query.get(session['user_id']))
    return dict(current_user=None)

# ====================== WEBHOOK ======================
@app.route('/webhook', methods=['POST'])
def ping_webhook():
    data = request.get_json()
    msg_id = data.get('messageId')
    status = data.get('status')
    if msg_id:
        Message.query.filter_by(msg_id=msg_id).update({'status': status})
        db.session.commit()
    return '', 200

# ====================== PWA ======================
@app.route('/manifest.json')
def manifest():
    return jsonify({
        "name": "FeeRemind Pro",
        "short_name": "FeeRemind",
        "start_url": "/dashboard",
        "display": "standalone",
        "background_color": "#0f172a",
        "theme_color": "#2563eb",
        "icons": [{"src": "/static/ic-1.png", "sizes": "192x192"}]
    })

@app.route('/sw.js')
def sw():
    return app.send_static_file('sw.js')

# ====================== AUTH ROUTES ======================
@app.route('/register', methods=['GET', 'POST'])
def register():
    if session.get('user_id', 0) > 2:
        return redirect('/dashboard')
    if request.method == 'POST':
        email = request.form['email'].strip()
        password = request.form['password'].strip()
        school_name = request.form['school_name'].strip()
        sender_id = request.form['sender_id'].strip()
        paycode = request.form['paycode'].strip()

        if User.query.filter_by(email=email).first():
	            flash('Email already exists', 'error')
            return render_template('register.html')

        user = User(
            email=email,
            password_hash=generate_password_hash(password),
            school_name=school_name,
            sender_id=sender_id,
            paycode=paycode
        )
        db.session.add(user)
        db.session.commit()
        flash('Account created. Login now.', 'success')
        return redirect(url_for('login'))
    return render_template('register.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form['email']
        password = request.form['password']
        user = User.query.filter_by(email=email).first()
        if user and check_password_hash(user.password_hash, password):
            session['user_id'] = user.id
            return redirect(url_for('dashboard'))
        flash('Invalid credentials', 'error')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# ====================== APP ROUTES ======================
@app.route('/')
def index():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))



@app.route('/dashboard')
@login_required
def dashboard():
    user = User.query.get(session['user_id'])
    today = datetime.utcnow().date()

    # ====== Simplified Stats ======
    stats = db.session.query(
        func.count(Message.id).label('total'),
        func.count(case((Message.status == 'Delivered', 1))).label('delivered'),
        func.count(case((Message.status == 'Failed', 1))).label('failed'),
        func.count(case((Message.status == 'sent', 1))).label('sent')
    ).filter(
        Message.user_id == user.id,
        func.date(Message.created_at) == today
    ).first()

    # ====== Recent Uploads ======
    uploads = db.session.query(
        func.date(Message.created_at).label('date'),
        func.count(Message.id).label('count'),
        func.count(case((Message.status == 'Delivered', 1))).label('delivered')
    ).filter(
        Message.user_id == user.id
    ).group_by(
        func.date(Message.created_at)
    ).order_by(
        func.date(Message.created_at).desc()
    ).limit(7).all()

    return render_template(
        'dashboard.html',
        user=user,
        stats=stats,
        uploads=uploads,
        school_name=user.school_name
    )
    
    
@app.route('/upload', methods=['GET', 'POST'])
@login_required
def upload():
    user = User.query.get(session['user_id'])
    if user.sms_credits <= 0:
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
            # try to auto-find currency column too
            curr_c = next((h for h in headers if h.lower() in ['currency','curr','ccy']), None)
        except Exception as e:
            flash(f'File error: {str(e)}', 'error')
            os.remove(path)
            return redirect(url_for('upload'))

        recipients = []
        for r in rows:
            try:
                raw_bal = str(r.get(bal_c, '0'))
                # detect currency symbol from raw balance string
                currency = 'USD'  # default
                if 'ZWG' in raw_bal.upper() or 'ZW$' in raw_bal:
                    currency = 'ZWG'
                elif 'USD' in raw_bal.upper() or '$' in raw_bal and 'ZWG' not in raw_bal.upper():
                    currency = 'USD'

                bal = float(raw_bal.replace('$','').replace('USD','').replace('ZWG','').replace('ZW$','').replace(',','').strip() or 0)
            except:
                continue

            if bal > 0 and user.sms_credits > len(recipients):
                recipients.append({
                    "name": str(r.get(name_c, 'Parent')),
                    "phone": str(r.get(phone_c, '')),
                    "balance": bal,
                    "currency": currency, 
                    "due": str(r.get('due', r.get('due_date', r.get('duedate', 'ASAP')))),
                    "paycode": user.paycode or ''
                })

        os.remove(path)

        MAX_ROWS_PER_UPLOAD = 100
        if len(recipients) > MAX_ROWS_PER_UPLOAD:
            flash(f'Too many valid rows: {len(recipients)}. Max allowed is {MAX_ROWS_PER_UPLOAD}. Split your file and try again.', 'error')
            return redirect(url_for('upload'))

        if recipients:
            result = send_sms_sync(user.id, recipients)
            flash(
            f"Batch created: {result['batch']} Status: {result['status']}",
            "success"
            )
        else:
            flash('No valid recipients found', 'error')

        return redirect(url_for('dashboard'))

    return render_template('upload.html', user=user, school_name=user.school_name)
      
    
@app.route('/settings', methods=['GET', 'POST'])
@login_required
def settings():
    user = User.query.get(session['user_id'])
    if request.method == 'POST':
        user.school_name = request.form['school_name'].strip()
        user.password_hash = generate_password_hash(request.form['password'].strip())
        user.message_theme = request.form['theme'].strip()
        if request.form['password'].strip() != request.form['confirm_pass'].strip():
            flash('Passwords do not match', 'error')
            return redirect(url_for('settings'))
        db.session.commit()
        flash('Settings saved', 'success')
        return redirect(url_for('settings'))
    return render_template('settings.html', user=user, school_name=user.school_name)


@app.route('/api/status')
@login_required
def api_status():
    user = User.query.get(session['user_id'])
    today = datetime.utcnow().date()
    data = dict(db.session.query(Message.status, db.func.count(Message.id))
               .filter(Message.user_id == user.id, db.func.date(Message.created_at) == today)
               .group_by(Message.status).all())
    return jsonify(data)

@app.route('/help', methods=['GET', 'POST'])
@login_required
def help():
    user = User.query.get(session['user_id'])
    if request.method == 'POST':
        ticket = Ticket(user_id=user.id, subject=request.form['subject'], body=request.form['body'])
        db.session.add(ticket)
        db.session.commit()
        flash('Ticket submitted. We reply within 4 hours.', 'success')
    return render_template('help.html', user=user, school_name=user.school_name)

# ====================== RUN APP ======================
if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))