from flask import Flask, render_template, request, redirect, url_for, session, send_file, flash
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
import sqlite3, os, uuid
from scripts.nikuy_kokhaviyot import process_xls

app = Flask(__name__)
app.secret_key = 'zman_emet_secret_2024_change_this'

DB = 'platform.db'
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

SCRIPTS = {
    'nikuy': {
        'id': 'nikuy',
        'name': 'ניקוי כוכביות',
        'desc': 'מסיר * ו-? מדוח נוכחות חודשי',
        'accept': '.xls,.xlsx',
        'icon': '🧹'
    }
}

def get_db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db() as db:
        db.execute('''CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            full_name TEXT,
            is_admin INTEGER DEFAULT 0,
            active INTEGER DEFAULT 1
        )''')
        db.execute('''CREATE TABLE IF NOT EXISTS permissions (
            user_id INTEGER,
            script_id TEXT,
            PRIMARY KEY (user_id, script_id)
        )''')
        # Create default admin if not exists
        existing = db.execute("SELECT id FROM users WHERE username='admin'").fetchone()
        if not existing:
            db.execute("INSERT INTO users (username, password, full_name, is_admin) VALUES (?, ?, ?, 1)",
                ('admin', generate_password_hash('admin123'), 'מנהל מערכת'))
        db.commit()

init_db()

def login_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('is_admin'):
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated

# ─── AUTH ───────────────────────────────────────────────────────────────────

@app.route('/', methods=['GET', 'POST'])
def login():
    if 'user_id' in session:
        return redirect(url_for('admin' if session.get('is_admin') else 'dashboard'))
    error = None
    if request.method == 'POST':
        username = request.form['username'].strip()
        password = request.form['password']
        with get_db() as db:
            user = db.execute("SELECT * FROM users WHERE username=? AND active=1", (username,)).fetchone()
        if user and check_password_hash(user['password'], password):
            session['user_id']  = user['id']
            session['username'] = user['username']
            session['name']     = user['full_name']
            session['is_admin'] = bool(user['is_admin'])
            return redirect(url_for('admin' if user['is_admin'] else 'dashboard'))
        error = 'שם משתמש או סיסמה שגויים'
    return render_template('login.html', error=error)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# ─── USER DASHBOARD ─────────────────────────────────────────────────────────

@app.route('/dashboard')
@login_required
def dashboard():
    if session.get('is_admin'):
        return redirect(url_for('admin'))
    with get_db() as db:
        perms = db.execute("SELECT script_id FROM permissions WHERE user_id=?",
                           (session['user_id'],)).fetchall()
    allowed = [SCRIPTS[p['script_id']] for p in perms if p['script_id'] in SCRIPTS]
    return render_template('dashboard.html', scripts=allowed)

@app.route('/run/<script_id>', methods=['GET', 'POST'])
@login_required
def run_script(script_id):
    if session.get('is_admin'):
        return redirect(url_for('admin'))
    # Check permission
    with get_db() as db:
        perm = db.execute("SELECT 1 FROM permissions WHERE user_id=? AND script_id=?",
                          (session['user_id'], script_id)).fetchone()
    if not perm or script_id not in SCRIPTS:
        flash('אין לך הרשאה לסקריפט זה')
        return redirect(url_for('dashboard'))

    script = SCRIPTS[script_id]
    result = None
    error = None

    if request.method == 'POST':
        file = request.files.get('file')
        if not file or file.filename == '':
            error = 'לא נבחר קובץ'
        else:
            uid = str(uuid.uuid4())[:8]
            filename = secure_filename(file.filename)
            in_path  = os.path.join(UPLOAD_FOLDER, f'{uid}_{filename}')
            out_name = filename.rsplit('.', 1)[0] + '_ללא_כוכביות.xlsx'
            out_path = os.path.join(OUTPUT_FOLDER, f'{uid}_{out_name}')
            file.save(in_path)
            try:
                process_xls(in_path, out_path)
                result = f'{uid}_{out_name}'
            except Exception as e:
                error = f'שגיאה בעיבוד: {str(e)}'
            finally:
                try: os.remove(in_path)
                except: pass

    return render_template('run.html', script=script, result=result, error=error)

@app.route('/download/<filename>')
@login_required
def download(filename):
    path = os.path.join(OUTPUT_FOLDER, filename)
    if not os.path.exists(path):
        flash('הקובץ לא נמצא')
        return redirect(url_for('dashboard'))
    return send_file(path, as_attachment=True,
                     download_name=filename.split('_', 1)[-1] if '_' in filename else filename)

# ─── ADMIN ──────────────────────────────────────────────────────────────────

@app.route('/admin')
@login_required
@admin_required
def admin():
    with get_db() as db:
        users = db.execute("SELECT * FROM users WHERE is_admin=0").fetchall()
        perms = db.execute("SELECT * FROM permissions").fetchall()
    user_perms = {}
    for p in perms:
        user_perms.setdefault(p['user_id'], set()).add(p['script_id'])
    return render_template('admin.html', users=users, scripts=SCRIPTS, user_perms=user_perms)

@app.route('/admin/add_user', methods=['POST'])
@login_required
@admin_required
def add_user():
    username  = request.form['username'].strip()
    password  = request.form['password']
    full_name = request.form['full_name'].strip()
    try:
        with get_db() as db:
            db.execute("INSERT INTO users (username, password, full_name) VALUES (?, ?, ?)",
                       (username, generate_password_hash(password), full_name))
            db.commit()
        flash(f'משתמש {full_name} נוצר בהצלחה')
    except sqlite3.IntegrityError:
        flash('שם משתמש כבר קיים')
    return redirect(url_for('admin'))

@app.route('/admin/delete_user/<int:uid>')
@login_required
@admin_required
def delete_user(uid):
    with get_db() as db:
        db.execute("DELETE FROM users WHERE id=?", (uid,))
        db.execute("DELETE FROM permissions WHERE user_id=?", (uid,))
        db.commit()
    flash('משתמש נמחק')
    return redirect(url_for('admin'))

@app.route('/admin/set_password/<int:uid>', methods=['POST'])
@login_required
@admin_required
def set_password(uid):
    new_pass = request.form['new_password']
    with get_db() as db:
        db.execute("UPDATE users SET password=? WHERE id=?",
                   (generate_password_hash(new_pass), uid))
        db.commit()
    flash('סיסמה עודכנה')
    return redirect(url_for('admin'))

@app.route('/admin/permissions/<int:uid>', methods=['POST'])
@login_required
@admin_required
def set_permissions(uid):
    selected = request.form.getlist('scripts')
    with get_db() as db:
        db.execute("DELETE FROM permissions WHERE user_id=?", (uid,))
        for s in selected:
            if s in SCRIPTS:
                db.execute("INSERT OR IGNORE INTO permissions (user_id, script_id) VALUES (?, ?)", (uid, s))
        db.commit()
    flash('הרשאות עודכנו')
    return redirect(url_for('admin'))

if __name__ == '__main__':
    app.run(debug=True)
