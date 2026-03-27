<!DOCTYPE html>
<html dir="rtl" lang="he">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>כניסה למערכת | זמן אמת</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; background: #f0f4ff; min-height: 100vh; display: flex; align-items: center; justify-content: center; }
  .card { background: #fff; border-radius: 20px; box-shadow: 0 8px 40px rgba(37,99,235,0.12); padding: 2.5rem 2rem; width: 100%; max-width: 400px; }
  .logo { text-align: center; margin-bottom: 1.75rem; }
  .logo-icon { font-size: 40px; }
  .logo h1 { font-size: 20px; font-weight: 700; color: #1e3a8a; margin-top: 8px; }
  .logo p { font-size: 12px; color: #94a3b8; margin-top: 3px; }
  label { display: block; font-size: 13px; font-weight: 600; color: #374151; margin-bottom: 5px; }
  input { width: 100%; padding: 11px 14px; border: 1.5px solid #e2e8f0; border-radius: 10px; font-size: 14px; margin-bottom: 1rem; outline: none; transition: border 0.2s; font-family: inherit; }
  input:focus { border-color: #2563eb; }
  .btn { width: 100%; padding: 12px; background: #2563eb; color: #fff; border: none; border-radius: 10px; font-size: 15px; font-weight: 700; cursor: pointer; transition: background 0.15s; font-family: inherit; }
  .btn:hover { background: #1d4ed8; }
  .error { background: #fef2f2; border: 1px solid #fecaca; color: #dc2626; border-radius: 8px; padding: 10px 14px; font-size: 13px; margin-bottom: 1rem; }
  .footer { text-align: center; margin-top: 1.5rem; font-size: 11px; color: #cbd5e1; }
</style>
</head>
<body>
<div class="card">
  <div class="logo">
    <div class="logo-icon">⏱</div>
    <h1>זמן אמת</h1>
    <p>מערכת לניהול נוכחות ושכר</p>
  </div>
  {% if error %}
  <div class="error">{{ error }}</div>
  {% endif %}
  <form method="POST">
    <label>שם משתמש</label>
    <input type="text" name="username" autocomplete="username" required autofocus>
    <label>סיסמה</label>
    <input type="password" name="password" autocomplete="current-password" required>
    <button class="btn" type="submit">כניסה למערכת</button>
  </form>
  <div class="footer">© זמן אמת – כל הזכויות שמורות</div>
</div>
</body>
</html>
