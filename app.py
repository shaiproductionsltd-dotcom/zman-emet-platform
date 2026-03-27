from flask import Flask, request, redirect, session, send_file
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
import sqlite3, os, uuid

import xlrd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

H_ALIGN = {1:'left',2:'center',3:'right',4:'fill',5:'justify',6:'centerContinuous',7:'distributed'}
V_ALIGN = {0:'top',1:'center',2:'bottom',3:'justify',4:'distributed'}
NO_BORDER = Border(left=Side(border_style=None),right=Side(border_style=None),
                   top=Side(border_style=None),bottom=Side(border_style=None))

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
    wb_in = xlrd.open_workbook(input_path, formatting_info=True)
    cmap  = wb_in.colour_map
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
                val  = clean(ws_in.cell_value(r, c))
                xf   = wb_in.xf_list[ws_in.cell_xf_index(r, c)]
                fi   = wb_in.font_list[xf.font_index]
                cell = ws_out.cell(row=r+1, column=c+1, value=val)
                fc   = idx_to_hex(cmap, fi.colour_index)
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

CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Segoe UI', Arial, sans-serif; background: #f0f4ff; min-height: 100vh; direction: rtl; }
.topbar { background: #1e3a8a; color: white; padding: 0 2rem; height: 58px; display: flex; align-items: center; justify-content: space-between; }
.topbar h1 { font-size: 17px; font-weight: 700; }
.topbar a { color: #93c5fd; font-size: 13px; text-decoration: none; }
.wrap { max-width: 900px; margin: 2rem auto; padding: 0 1rem; }
.login-wrap { max-width: 400px; margin: 5rem auto; padding: 0 1rem; }
.card { background: white; border-radius: 16px; box-shadow: 0 4px 24px rgba(37,99,235,.1); padding: 2rem; margin-bottom: 1.5rem; }
.card h2 { font-size: 16px; font-weight: 700; color: #1e3a8a; margin-bottom: 1rem; padding-bottom: .75rem; border-bottom: 1.5px solid #e0e7ff; }
label.field-label { font-size: 13px; font-weight: 600; color: #374151; margin-bottom: 5px; display: block; }
input[type=text], input[type=password] { padding: 9px 12px; border: 1.5px solid #e2e8f0; border-radius: 8px; font-size: 13px; font-family: inherit; outline: none; width: 100%; margin-bottom: .75rem; }
input:focus { border-color: #2563eb; }
.btn { padding: 10px 20px; border: none; border-radius: 8px; font-size: 13px; font-weight: 600; cursor: pointer; font-family: inherit; }
.btn-blue { background: #2563eb; color: white; }
.btn-blue:hover { background: #1d4ed8; }
.btn-red { background: #fee2e2; color: #dc2626; }
.btn-gray { background: #f1f5f9; color: #475569; }
.flash { background: #f0fdf4; border: 1px solid #86efac; color: #15803d; border-radius: 8px; padding: 10px 14px; font-size: 13px; margin-bottom: 1rem; }
.flash-err { background: #fef2f2; border: 1px solid #fecaca; color: #dc2626; border-radius: 8px; padding: 10px 14px; font-size: 13px; margin-bottom: 1rem; }
table { width: 100%; border-collapse: collapse; font-size: 13px; }
th { text-align: right; padding: 10px 12px; background: #f8fafc; color: #64748b; font-weight: 600; border-bottom: 1.5px solid #e2e8f0; }
td { padding: 12px; border-bottom: 1px solid #f1f5f9; vertical-align: middle; }
.badge { display: inline-block; padding: 3px 10px; border-radius: 99px; font-size: 11px; font-weight: 600; background: #f1f5f9; color: #64748b; }
.form-row { display: flex; gap: 10px; flex-wrap: wrap; align-items: flex-end; }
.form-group { flex: 1; min-width: 130px; }
.drop-zone { border: 2px dashed #c7d7f5; border-radius: 14px; padding: 2rem; text-align: center; cursor: pointer; background: #fafcff; margin-bottom: 1rem; }
.drop-zone:hover { border-color: #2563eb; background: #eff6ff; }
.success-box { padding: 1.25rem; background: #f0fdf4; border: 1.5px solid #86efac; border-radius: 13px; text-align: center; margin-top: 1rem; }
.dl-btn { display: inline-block; padding: 11px 28px; background: #16a34a; color: white; border-radius: 9px; font-size: 14px; font-weight: 700; text-decoration: none; }
.modal-bg { display: none; position: fixed; inset: 0; background: rgba(0,0,0,.4); z-index: 100; align-items: center; justify-content: center; }
.modal-box { background: white; border-radius: 16px; padding: 1.75rem; width: 320px; }
"""

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
            active INTEGER DEFAULT 1)''')
        db.execute('''CREATE TABLE IF NOT EXISTS permissions (
            user_id INTEGER, script_id TEXT,
            PRIMARY KEY (user_id, script_id))''')
        if not db.execute("SELECT id FROM users WHERE username='admin'").fetchone():
            db.execute("INSERT INTO users(username,password,full_name,is_admin)VALUES(?,?,?,1)",
                ('admin', generate_password_hash('admin123'), 'מנהל מערכת'))
        db.commit()

init_db()

def add_flash(msg):
    session.setdefault('msgs', []).append(msg)

def pop_flashes():
    msgs = session.pop('msgs', [])
    return ''.join('<div class="flash">' + m + '</div>' for m in msgs)

def render(title, body, nav=True):
    topbar = ''
    if nav:
        name = session.get('name', '')
        topbar = (
            '<div class="topbar">'
            '<h1>&#9201; זמן אמת</h1>'
            '<div style="display:flex;gap:16px;align-items:center">'
            '<span style="font-size:13px;color:#93c5fd">שלום, ' + name + '</span>'
            '<a href="/logout">יציאה</a>'
            '</div></div>'
        )
    wrap_cls = 'wrap' if nav else 'login-wrap'
    return (
        '<!DOCTYPE html><html dir="rtl" lang="he">'
        '<head><meta charset="UTF-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1">'
        '<title>' + title + ' | זמן אמת</title>'
        '<style>' + CSS + '</style></head>'
        '<body>' + topbar +
        '<div class="' + wrap_cls + '">' + pop_flashes() + body + '</div>'
        '</body></html>'
    )

def login_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect('/')
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('is_admin'):
            return redirect('/dashboard')
        return f(*args, **kwargs)
    return decorated


@app.route('/', methods=['GET', 'POST'])
def login():
    if 'user_id' in session:
        return redirect('/admin' if session.get('is_admin') else '/dashboard')
    error = ''
    if request.method == 'POST':
        u = request.form['username'].strip()
        p = request.form['password']
        with get_db() as db:
            user = db.execute("SELECT * FROM users WHERE username=? AND active=1", (u,)).fetchone()
        if user and check_password_hash(user['password'], p):
            session.update({
                'user_id': user['id'],
                'username': user['username'],
                'name': user['full_name'],
                'is_admin': bool(user['is_admin'])
            })
            return redirect('/admin' if user['is_admin'] else '/dashboard')
        error = '<div class="flash-err">שם משתמש או סיסמה שגויים</div>'

    body = (
        '<div class="card" style="padding:2rem">'
        '<div style="text-align:center;margin-bottom:1.5rem">'
        '<div style="font-size:40px">&#9201;</div>'
        '<h1 style="font-size:20px;font-weight:700;color:#1e3a8a;margin-top:8px">זמן אמת</h1>'
        '<p style="font-size:12px;color:#888;margin-top:3px">מערכת לניהול נוכחות ושכר</p>'
        '</div>' + error +
        '<form method="POST">'
        '<label class="field-label">שם משתמש</label>'
        '<input type="text" name="username" required autofocus>'
        '<label class="field-label">סיסמה</label>'
        '<input type="password" name="password" required>'
        '<button type="submit" class="btn btn-blue" style="width:100%;padding:12px;font-size:15px;margin-top:.5rem">כניסה למערכת</button>'
        '</form>'
        '<p style="text-align:center;margin-top:1.5rem;font-size:11px;color:#bbb">&#169; זמן אמת</p>'
        '</div>'
    )
    return render('כניסה', body, nav=False)


@app.route('/logout')
def logout():
    session.clear()
    return redirect('/')


@app.route('/dashboard')
@login_required
def dashboard():
    if session.get('is_admin'):
        return redirect('/admin')
    with get_db() as db:
        perms = db.execute(
            "SELECT script_id FROM permissions WHERE user_id=?",
            (session['user_id'],)
        ).fetchall()
    allowed = [SCRIPTS[p['script_id']] for p in perms if p['script_id'] in SCRIPTS]

    cards = ''
    for s in allowed:
        cards += (
            '<a href="/run/' + s['id'] + '" style="background:white;border-radius:16px;'
            'box-shadow:0 2px 16px rgba(0,0,0,.06);padding:1.5rem;text-decoration:none;display:block">'
            '<div style="font-size:36px;margin-bottom:.75rem">' + s['icon'] + '</div>'
            '<div style="font-size:15px;font-weight:700;color:#1e3a8a;margin-bottom:4px">' + s['name'] + '</div>'
            '<div style="font-size:12px;color:#64748b">' + s['desc'] + '</div>'
            '</a>'
        )

    if not allowed:
        cards = (
            '<div style="text-align:center;padding:3rem;color:#94a3b8">'
            '<div style="font-size:48px;margin-bottom:1rem">&#128274;</div>'
            '<div>אין כלים זמינים עדיין</div>'
            '</div>'
        )

    body = (
        '<h2 style="font-size:22px;font-weight:700;color:#1e3a8a;margin-bottom:.4rem">'
        'שלום, ' + session['name'] + ' &#128075;</h2>'
        '<p style="font-size:14px;color:#64748b;margin-bottom:2rem">הכלים הזמינים עבורך:</p>'
        '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:1rem">'
        + cards + '</div>'
    )
    return render('הכלים שלי', body)


@app.route('/run/<script_id>', methods=['GET', 'POST'])
@login_required
def run_script(script_id):
    if session.get('is_admin'):
        return redirect('/admin')
    with get_db() as db:
        perm = db.execute(
            "SELECT 1 FROM permissions WHERE user_id=? AND script_id=?",
            (session['user_id'], script_id)
        ).fetchone()
    if not perm or script_id not in SCRIPTS:
        add_flash('אין לך הרשאה לסקריפט זה')
        return redirect('/dashboard')

    scr    = SCRIPTS[script_id]
    result = None
    error  = ''

    if request.method == 'POST':
        f = request.files.get('file')
        if not f or f.filename == '':
            error = '<div class="flash-err">לא נבחר קובץ</div>'
        else:
            uid = str(uuid.uuid4())[:8]
            fn  = secure_filename(f.filename)
            inp = os.path.join(UPLOAD_FOLDER, uid + '_' + fn)
            onm = fn.rsplit('.', 1)[0] + '_ללא_כוכביות.xlsx'
            out = os.path.join(OUTPUT_FOLDER, uid + '_' + onm)
            f.save(inp)
            try:
                process_xls(inp, out)
                result = uid + '_' + onm
            except Exception as e:
                error = '<div class="flash-err">שגיאה בעיבוד: ' + str(e) + '</div>'
            finally:
                try:
                    os.remove(inp)
                except: pass

    if result:
        content = (
            '<div class="success-box">'
            '<div style="font-size:32px;margin-bottom:6px">&#9989;</div>'
            '<div style="font-size:16px;font-weight:700;color:#15803d;margin-bottom:10px">הקובץ מוכן!</div>'
            '<a href="/download/' + result + '" class="dl-btn">&#8681; הורד קובץ נקי</a>'
            '<br><br><a href="/run/' + script_id + '" style="font-size:13px;color:#2563eb">עבד קובץ נוסף</a>'
            '</div>'
        )
    else:
        content = (
            error +
            '<form method="POST" enctype="multipart/form-data">'
            '<div class="drop-zone" onclick="document.getElementById(\'fi\').click()">'
            '<input type="file" name="file" id="fi" accept="' + scr['accept'] + '" style="display:none"'
            ' onchange="document.getElementById(\'lbl\').textContent=this.files[0].name;'
            'document.getElementById(\'gb\').disabled=false">'
            '<div style="font-size:32px;margin-bottom:8px">&#128194;</div>'
            '<div style="font-size:15px;font-weight:600;color:#1e40af;margin-bottom:4px">לחץ לבחירת קובץ</div>'
            '<div style="font-size:12px;color:#94a3b8" id="lbl">' + scr['accept'] + '</div>'
            '</div>'
            '<button type="submit" class="btn btn-blue" id="gb" disabled'
            ' style="width:100%;padding:13px;font-size:15px;font-weight:700">'
            + scr['icon'] + ' הפעל</button>'
            '</form>'
        )

    body = (
        '<a href="/dashboard" style="color:#2563eb;font-size:13px;text-decoration:none;display:block;margin-bottom:1rem">&#8592; חזרה לכלים</a>'
        '<div class="card">'
        '<div style="font-size:40px;margin-bottom:.5rem">' + scr['icon'] + '</div>'
        '<div style="font-size:20px;font-weight:700;color:#1e3a8a;margin-bottom:4px">' + scr['name'] + '</div>'
        '<div style="font-size:13px;color:#64748b;margin-bottom:1.75rem">' + scr['desc'] + '</div>'
        + content +
        '</div>'
    )
    return render(scr['name'], body)


@app.route('/download/<filename>')
@login_required
def download(filename):
    path = os.path.join(OUTPUT_FOLDER, filename)
    if not os.path.exists(path):
        add_flash('הקובץ לא נמצא')
        return redirect('/dashboard')
    dn = filename.split('_', 1)[-1] if '_' in filename else filename
    return send_file(path, as_attachment=True, download_name=dn)


@app.route('/admin')
@login_required
@admin_required
def admin():
    with get_db() as db:
        users = db.execute("SELECT * FROM users WHERE is_admin=0").fetchall()
        perms = db.execute("SELECT * FROM permissions").fetchall()
    up = {}
    for p in perms:
        up.setdefault(p['user_id'], set()).add(p['script_id'])

    rows = ''
    for u in users:
        uid  = u['id']
        uname = u['username']
        uname_full = u['full_name']

        checks = ''
        for sid, s in SCRIPTS.items():
            checked = 'checked' if (uid in up and sid in up[uid]) else ''
            checks += (
                '<label style="display:flex;align-items:center;gap:5px;font-size:13px;margin-left:10px">'
                '<input type="checkbox" name="scripts" value="' + sid + '" ' + checked + '>'
                + s['icon'] + ' ' + s['name'] +
                '</label>'
            )

        perm_form = (
            '<form method="POST" action="/admin/permissions/' + str(uid) + '" style="display:inline">'
            '<div style="display:flex;flex-wrap:wrap">' + checks + '</div>'
            '<button type="submit" class="btn btn-gray" style="margin-top:6px;font-size:12px;padding:5px 12px">שמור</button>'
            '</form>'
        )

        pass_btn = (
            '<button class="btn btn-gray" style="font-size:12px;padding:5px 12px"'
            ' onclick="openPass(' + str(uid) + ',\'' + uname_full + '\')">'
            'שנה סיסמה</button>'
        )

        del_link = (
            '<a href="/admin/delete/' + str(uid) + '" '
            'onclick="return confirm(\'למחוק?\');" '
            'class="btn btn-red" style="text-decoration:none;font-size:12px;padding:5px 12px">מחק</a>'
        )

        rows += (
            '<tr>'
            '<td><strong>' + uname_full + '</strong></td>'
            '<td><span class="badge">' + uname + '</span></td>'
            '<td>' + perm_form + '</td>'
            '<td>' + pass_btn + '</td>'
            '<td>' + del_link + '</td>'
            '</tr>'
        )

    table = (
        '<table><thead><tr>'
        '<th>שם</th><th>משתמש</th><th>הרשאות</th><th>סיסמה</th><th>מחק</th>'
        '</tr></thead><tbody>' + rows + '</tbody></table>'
    ) if users else '<p style="color:#94a3b8;text-align:center;padding:2rem">אין לקוחות עדיין</p>'

    body = (
        '<div class="card">'
        '<h2>&#10133; הוספת לקוח חדש</h2>'
        '<form method="POST" action="/admin/add_user">'
        '<div class="form-row">'
        '<div class="form-group"><label class="field-label">שם מלא</label>'
        '<input type="text" name="full_name" placeholder="שם הלקוח" required style="margin-bottom:0"></div>'
        '<div class="form-group"><label class="field-label">שם משתמש</label>'
        '<input type="text" name="username" placeholder="לכניסה למערכת" required style="margin-bottom:0"></div>'
        '<div class="form-group"><label class="field-label">סיסמה</label>'
        '<input type="password" name="password" placeholder="סיסמה ראשונית" required style="margin-bottom:0"></div>'
        '<button type="submit" class="btn btn-blue" style="height:40px;align-self:flex-end">הוסף</button>'
        '</div></form></div>'
        '<div class="card"><h2>&#128101; לקוחות במערכת</h2>' + table + '</div>'
        '<div class="modal-bg" id="passModal">'
        '<div class="modal-box">'
        '<h3 style="font-size:15px;font-weight:700;margin-bottom:1rem;color:#1e3a8a">'
        'שינוי סיסמה &#8212; <span id="pname"></span></h3>'
        '<form method="POST" id="pform">'
        '<input type="password" name="new_password" placeholder="סיסמה חדשה" required>'
        '<div style="display:flex;gap:8px;margin-top:.5rem;justify-content:flex-end">'
        '<button type="button" class="btn btn-gray" onclick="closePass()">ביטול</button>'
        '<button type="submit" class="btn btn-blue">עדכן</button>'
        '</div></form></div></div>'
        '<script>'
        'function openPass(id,name){'
        'document.getElementById("pname").textContent=name;'
        'document.getElementById("pform").action="/admin/setpass/"+id;'
        'document.getElementById("passModal").style.display="flex";}'
        'function closePass(){'
        'document.getElementById("passModal").style.display="none";}'
        '</script>'
    )
    return render('ניהול', body)


@app.route('/admin/add_user', methods=['POST'])
@login_required
@admin_required
def add_user():
    u = request.form['username'].strip()
    p = request.form['password']
    n = request.form['full_name'].strip()
    try:
        with get_db() as db:
            db.execute(
                "INSERT INTO users(username,password,full_name)VALUES(?,?,?)",
                (u, generate_password_hash(p), n)
            )
            db.commit()
        add_flash('משתמש ' + n + ' נוצר בהצלחה')
    except sqlite3.IntegrityError:
        add_flash('שם משתמש כבר קיים')
    return redirect('/admin')


@app.route('/admin/delete/<int:uid>')
@login_required
@admin_required
def delete_user(uid):
    with get_db() as db:
        db.execute("DELETE FROM users WHERE id=?", (uid,))
        db.execute("DELETE FROM permissions WHERE user_id=?", (uid,))
        db.commit()
    add_flash('משתמש נמחק')
    return redirect('/admin')


@app.route('/admin/setpass/<int:uid>', methods=['POST'])
@login_required
@admin_required
def set_password(uid):
    with get_db() as db:
        db.execute(
            "UPDATE users SET password=? WHERE id=?",
            (generate_password_hash(request.form['new_password']), uid)
        )
        db.commit()
    add_flash('סיסמה עודכנה')
    return redirect('/admin')


@app.route('/admin/permissions/<int:uid>', methods=['POST'])
@login_required
@admin_required
def set_permissions(uid):
    selected = request.form.getlist('scripts')
    with get_db() as db:
        db.execute("DELETE FROM permissions WHERE user_id=?", (uid,))
        for s in selected:
            if s in SCRIPTS:
                db.execute(
                    "INSERT OR IGNORE INTO permissions(user_id,script_id)VALUES(?,?)",
                    (uid, s)
                )
        db.commit()
    add_flash('הרשאות עודכנו')
    return redirect('/admin')


if __name__ == '__main__':
    app.run(debug=True)
