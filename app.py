from flask import Flask, render_template, request, redirect, url_for, session, send_file, flash
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
import sqlite3, os, uuid

import xlrd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

H_ALIGN = {1:'left',2:'center',3:'right',4:'fill',5:'justify',6:'centerContinuous',7:'distributed'}
V_ALIGN = {0:'top',1:'center',2:'bottom',3:'justify',4:'distributed'}
NO_BORDER = Border(left=Side(border_style=None), right=Side(border_style=None),
                   top=Side(border_style=None), bottom=Side(border_style=None))

def clean(val):
    if isinstance(val, str):
        return val.replace('*','').replace('?','').strip()
    return val

def idx_to_hex(cmap, idx):
    if idx in (0,64,65,32767,None): return None
    rgb = cmap.get(idx)
    if rgb and None not in rgb:
        return '{:02X}{:02X}{:02X}'.format(*rgb)
    return None

def process_xls(input_path, output_path):
    wb_in  = xlrd.open_workbook(input_path, formatting_info=True)
    cmap   = wb_in.colour_map
    wb_out = Workbook()
    wb_out.remove(wb_out.active)
    for s_idx in range(wb_in.nsheets):
        ws_in  = wb_in.sheet_by_index(s_idx)
        ws_out = wb_out.create_sheet(title=ws_in.name)
        ws_out.sheet_view.rightToLeft = True
        ws_out.sheet_view.showGridLines = False
        for c in range(ws_in.ncols):
            ltr = get_column_letter(c+1)
            ci  = ws_in.colinfo_map.get(c)
            w   = (ci.width/256.0) if (ci and ci.width) else 1.0
            ws_out.column_dimensions[ltr].width = max(w, 0.5)
        for r in range(ws_in.nrows):
            ri = ws_in.rowinfo_map.get(r)
            orig_pts = (ri.height/20.0) if (ri and ri.height) else 12.75
            has_content = any(v != '' for v in ws_in.row_values(r))
            if has_content:
                new_h = orig_pts if orig_pts >= 25 else (18 if orig_pts >= 15 else 15)
            else:
                new_h = 2 if orig_pts <= 8 else (orig_pts if orig_pts <= 50 else 12)
            ws_out.row_dimensions[r+1].height = new_h
        for r in range(ws_in.nrows):
            for c in range(ws_in.ncols):
                val = clean(ws_in.cell_value(r, c))
                xf  = wb_in.xf_list[ws_in.cell_xf_index(r, c)]
                fi  = wb_in.font_list[xf.font_index]
                cell = ws_out.cell(row=r+1, column=c+1, value=val)
                fc = idx_to_hex(cmap, fi.colour_index)
                cell.font = Font(bold=bool(fi.bold), italic=bool(fi.italic),
                                 size=fi.height/20, name=fi.name or 'Arial',
                                 color=fc or '000000')
                bg = idx_to_hex(cmap, xf.background.pattern_colour_index)
                if bg and xf.background.fill_pattern != 0:
                    cell.fill = PatternFill(fill_type='solid', fgColor=bg)
                cell.alignment = Alignment(
                    horizontal=H_ALIGN.get(xf.alignment.hor_align,'general'),
                    vertical=V_ALIGN.get(xf.alignment.vert_align,'bottom'),
                    wrapText=bool(xf.alignment.text_wrapped),
                    shrinkToFit=bool(xf.alignment.shrink_to_fit),
                    readingOrder=2)
                cell.border = NO_BORDER
        for r1,r2,c1,c2 in ws_in.merged_cells:
            if r2>r1 or c2>c1:
                try:
                    ws_out.merge_cells(start_row=r1+1, start_column=c1+1,
                                       end_row=r2, end_column=c2)
                except: pass
    for ws in wb_out.worksheets:
        ws.column_dimensions['G'].width  += 4.5
        ws.column_dimensions['S'].width  += 3.0
        ws.column_dimensions['AU'].width += 3.0
    wb_out.save(output_path)

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'zman_emet_secret_2024')

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
    with get_db() as db:
        perm = db.execute("SELECT 1 FROM permissions WHERE user_id=? AND script_id=?",
                          (session['user_id'], script_id)).fetchone()
    if not perm or script_id not in SCRIPTS:
        flash('אין לך הרשאה לסקריפט זה')
        return redirect(url_for('dashboard'))
    script = SCRIPTS[script_id]
    result = None
    error  = None
    if request.method == 'POST':
        file = request.files.get('file')
        if not file or file.filename == '':
            error = 'לא נבחר קובץ'
        else:
            uid      = str(uuid.uuid4())[:8]
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
    display_name = filename.split('_', 1)[-1] if '_' in filename else filename
    return send_file(path, as_attachment=True, download_name=display_name)

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
