import os
import re
import hashlib
import tempfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from functools import wraps
from flask import (
    Flask, render_template_string, request, redirect,
    url_for, session, jsonify, get_flashed_messages, flash
)
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func

try:
    import openpyxl
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'itvms_super_secret_2026_change_me')
app.permanent_session_lifetime = timedelta(days=30)

DB_USER = os.environ.get('DB_USER', 'itvms_user')
DB_PASS = os.environ.get('DB_PASS', 'StrongPass123!')
DB_HOST = os.environ.get('DB_HOST', 'localhost')
DB_NAME = os.environ.get('DB_NAME', 'itvms')

if os.environ.get('DB_ENGINE', 'sqlite') == 'mysql':
    app.config['SQLALCHEMY_DATABASE_URI'] = (
        f'mysql+pymysql://{DB_USER}:{DB_PASS}@{DB_HOST}/{DB_NAME}?charset=utf8mb4'
    )
    app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
        'pool_size': 10, 'pool_recycle': 3600, 'pool_pre_ping': True
    }
else:
    _db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'itvms_local.db')
    app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + _db_path
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

BRAND = 'Tech Expert ASTANA'
ADMIN_USER = 'bakyt'
ADMIN_HASH = hashlib.sha256('123123'.encode()).hexdigest()

class Workshop(db.Model):
    __tablename__ = 'workshops'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)

class Order(db.Model):
    __tablename__ = 'orders'
    id = db.Column(db.Integer, primary_key=True)
    client_code = db.Column(db.String(100), nullable=False)
    due_date = db.Column(db.String(20))
    created_at = db.Column(db.String(50), nullable=False)
    source_file = db.Column(db.String(255))
    workshop_name = db.Column(db.String(100), db.ForeignKey('workshops.name', ondelete='CASCADE'), nullable=False)

class OrderItem(db.Model):
    __tablename__ = 'order_items'
    id = db.Column(db.Integer, primary_key=True)
    order_id = db.Column(db.Integer, db.ForeignKey('orders.id', ondelete='CASCADE'), nullable=False)
    item_code = db.Column(db.String(100), nullable=False)
    item_name = db.Column(db.String(255), nullable=False)
    planned_qty = db.Column(db.Integer, nullable=False)
    scanned_qty = db.Column(db.Integer, default=0)
    designation = db.Column(db.String(100), nullable=True, default='')

with app.app_context():
    db.create_all()
    default_hash = hashlib.sha256('123123'.encode()).hexdigest()
    if Workshop.query.count() == 0:
        for wname in ['Основной цех', 'Цех №2', 'Цех №3']:
            db.session.add(Workshop(name=wname, password_hash=default_hash))
        db.session.commit()

def _clean_code(raw):
    cleaned = re.sub(r'[^0-9.]', '', str(raw).strip()).strip('.')
    return cleaned if cleaned else str(raw).strip()

def parse_excel(filepath):
    if not OPENPYXL_OK:
        raise RuntimeError('Модуль openpyxl не установлен на сервере. Выполните: pip install openpyxl')
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
    except Exception as e:
        raise ValueError(f'Не удалось прочитать Excel файл: {e}')
    ws = wb.active
    items, skipped = [], 0
    for row in ws.iter_rows(min_row=1):
        try:
            if len(row) < 6:
                skipped += 1
                continue
            code = row[1].value
            name = row[2].value
            qty = row[5].value
            if not code and not name:
                skipped += 1
                continue
            if not code or not name or qty is None:
                skipped += 1
                continue
            designation = row[3].value if row[3].value else ''
            items.append({
                'code': _clean_code(str(code)), 
                'name': str(name).strip(),
                'qty': max(1, int(float(str(qty)))), 
                'designation': str(designation).strip()
            })
        except:
            skipped += 1
    if not items:
        raise ValueError(f'В файле не обнаружены валидные данные (Столбцы: B=Код, C=Название, F=Кол-во). Пропущено строк: {skipped}')
    return items

def _xml_text(elem, tag):
    node = elem.find(tag)
    return node.text.strip() if (node is not None and node.text) else ''

def _xml_qty(elem):
    try:
        return max(int(float(_xml_text(elem, 'Количество') or '1')), 1)
    except ValueError:
        return 1

def parse_bazis(root):
    izd = root.find('.//Изделие')
    if izd is None:
        return []
    agg, seq = {}, []
    def walk(elem, mult):
        for child in elem:
            if child.tag == 'Объект':
                if _xml_text(child, 'ТипОбъекта') != 'Панель':
                    continue
                code = _xml_text(child, 'Обозначение') or _xml_text(child, 'КодДетали').lstrip('_')
                if not code:
                    continue
                name = _xml_text(child, 'Наименование')
                if code not in agg:
                    agg[code] = {'qty': 0, 'names': []}
                    seq.append(code)
                agg[code]['qty'] += _xml_qty(child) * mult
                agg[code]['designation'] = code
                if name and name not in agg[code]['names']:
                    agg[code]['names'].append(name)
            elif child.tag == 'Сборка':
                walk(child, mult * _xml_qty(child))
            else:
                walk(child, mult)
    walk(izd, _xml_qty(izd))
    return [{'code': c, 'name': (' / '.join(agg[c]['names']) or c)[:240], 'qty': agg[c]['qty'], 'designation': agg[c].get('designation', '')} for c in seq]

def parse_xml(filepath):
    try:
        tree = ET.parse(filepath)
    except ET.ParseError as e:
        raise ValueError(f'Ошибка структуры XML: {e}')
    root = tree.getroot()
    items = parse_bazis(root)
    if items:
        return items

    details = root.findall('.//Detail') + root.findall('.//Assembly')
    items = []
    if details:
        for elem in details:
            pos = (elem.findtext('Position') or '').strip()
            desig = (elem.findtext('Designation') or '').strip()
            name = (elem.findtext('Name') or '').strip()
            if desig.upper() in ('БЕЗ_ОБОЗН', 'БЕЗ ОБОЗН', ''):
                continue
            if not pos or not name:
                continue
            code = _clean_code(f'{pos}.{desig}' if desig else pos)
            items.append({'code': code, 'name': name, 'qty': 1, 'designation': desig})
    else:
        for elem in (root.findall('.//item') or list(root)):
            code = (elem.get('code') or (elem.find('code') is not None and elem.find('code').text) or '').strip()
            name = (elem.get('name') or (elem.find('name') is not None and elem.find('name').text) or '').strip()
            qty = elem.get('qty') or (elem.find('qty') is not None and elem.find('qty').text) or '1'
            if not code or not name: 
                continue
            items.append({'code': _clean_code(code), 'name': name, 'qty': max(1, int(float(qty))), 'designation': ''})
    if not items:
        raise ValueError('XML-файл не содержит поддерживаемых структур спецификаций деталей')
    return items

def parse_file(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    if ext in ('.xlsx', '.xls'): 
        return parse_excel(filepath)
    if ext == '.xml':            
        return parse_xml(filepath)
    raise ValueError(f'Формат файла {ext} не поддерживается системой')

def orders_for(workshop_name):
    rows = Order.query.filter_by(workshop_name=workshop_name).order_by(Order.id.desc()).all()
    result = []
    for o in rows:
        total = db.session.query(func.sum(OrderItem.planned_qty)).filter_by(order_id=o.id).scalar() or 0
        scanned = db.session.query(func.sum(OrderItem.scanned_qty)).filter_by(order_id=o.id).scalar() or 0
        pct = int(scanned / total * 100) if total else 0
        result.append({
            'id': o.id, 'client_code': o.client_code,
            'due_date': o.due_date or '—', 'created_at': o.created_at,
            'source_file': o.source_file or '—',
            'total': total, 'scanned': scanned, 'pct': pct
        })
    return result

def admin_ok():  
    return 'admin' in session

CSS = """
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,400;0,9..40,500;0,9..40,600;1,9..40,400&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
:root{
  --ink:#0b0e14;--ink2:#1c2130;--ink3:#2d3450;
  --paper:#f4f6fd;--card:#ffffff;
  --blue:#1f4fd8;--blue2:#3a6fe8;--blue-gl:rgba(31,79,216,.11);
  --green:#16a34a;--green-bg:#f0fdf4;
  --amber:#b45309;--amber-bg:#fffbeb;
  --red:#dc2626;--red-bg:#fef2f2;
  --muted:#697592;--border:#dde3f0;
  --r:16px;
}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
html{scroll-behavior:smooth}
body{background:var(--paper);background-image:radial-gradient(ellipse 90% 60% at 50% -10%,rgba(31,79,216,.05),transparent);color:var(--ink);font-family:'DM Sans',sans-serif;min-height:100vh;transition:background-color 0.4s ease}
a{text-decoration:none}

.nav{position:sticky;top:0;z-index:200;display:flex;align-items:center;justify-content:space-between;padding:0 40px;height:64px;background:rgba(254,255,255,.85);backdrop-filter:blur(20px);border-bottom:1px solid var(--border);transition:all 0.3s ease}
.nav-brand{font-family:'Bebas Neue',sans-serif;font-size:1.65rem;letter-spacing:.06em;background:linear-gradient(130deg,var(--blue),var(--blue2));-webkit-background-clip:text;background-clip:text;color:transparent}
.nav-right{display:flex;align-items:center;gap:12px}
.nav-chip{font-size:.8rem;font-weight:600;color:var(--blue);background:rgba(31,79,216,.06);border:1px solid rgba(31,79,216,.14);padding:5px 16px;border-radius:99px}

.btn{display:inline-flex;align-items:center;justify-content:center;gap:6px;padding:11px 24px;border-radius:99px;border:none;cursor:pointer;font-family:'DM Sans',sans-serif;font-size:.875rem;font-weight:600;transition:all .25s cubic-bezier(0.4, 0, 0.2, 1);text-decoration:none;white-space:nowrap}
.btn-blue{background:var(--blue);color:#fff;box-shadow:0 4px 12px var(--blue-gl)}
.btn-blue:hover{background:var(--blue2);transform:translateY(-2px);box-shadow:0 8px 20px rgba(31,79,216,.25)}
.btn-outline{background:#fff;color:var(--blue);border:1.5px solid var(--border)}
.btn-outline:hover{border-color:var(--blue);background:rgba(31,79,216,.03);transform:translateY(-1px)}
.btn-ghost{background:transparent;color:var(--muted);font-weight:500}
.btn-ghost:hover{color:var(--ink);background:rgba(0,0,0,.04);border-radius:99px}
.btn-danger{background:var(--red-bg);color:var(--red);border:1px solid #fee2e2}
.btn-danger:hover{background:#fecaca;transform:translateY(-1px)}
.btn-sm{padding:8px 20px;font-size:.84rem}
.btn-xs{padding:6px 14px;font-size:.78rem}

.wrap{max-width:1180px;margin:0 auto;padding:40px 24px;animation:fadeIn 0.4s ease}
.wrap-narrow{max-width:460px;margin:0 auto;padding:80px 24px;animation:fadeIn 0.4s ease}

.card{background:var(--card);border:1px solid var(--border);border-radius:var(--r);padding:32px;box-shadow:0 4px 20px rgba(0,0,0,.02);transition:all 0.3s ease}
.card:hover{box-shadow:0 6px 26px rgba(0,0,0,.04)}
.card+.card{margin-top:24px}
.card-title{font-family:'Bebas Neue',sans-serif;font-size:1.35rem;letter-spacing:.05em;color:var(--ink2);margin-bottom:22px}

.field{margin-bottom:18px}
.field label{display:block;font-size:.75rem;font-weight:700;color:var(--muted);margin-bottom:6px;letter-spacing:.05em;text-transform:uppercase}
input[type=text],input[type=password],input[type=date],select{width:100%;padding:12px 16px;background:#fff;border:1.5px solid var(--border);border-radius:12px;color:var(--ink);font-family:'DM Sans',sans-serif;font-size:.9rem;outline:none;transition:all .25s ease}
input:focus,select:focus{border-color:var(--blue);box-shadow:0 0 0 4px rgba(31,79,216,.08)}
input[type=file]{background:#fdfdff;border:1.5px dashed var(--border);border-radius:12px;padding:12px;color:var(--muted);cursor:pointer;width:100%;font-size:.85rem;transition:all 0.25s}
input[type=file]:hover{border-color:var(--blue);background:rgba(31,79,216,.01)}
input[type=checkbox]{width:16px;height:16px;accent-color:var(--blue);cursor:pointer}
.check-row{display:flex;align-items:center;gap:10px;font-size:.85rem;color:var(--muted);cursor:pointer;margin-bottom:20px}

.tbl-wrap{overflow-x:auto;border-radius:14px;border:1px solid var(--border);background:#fff}
table{width:100%;border-collapse:collapse;font-size:.875rem}
th{padding:12px 16px;background:#f8faff;color:var(--muted);font-size:.72rem;font-weight:700;text-transform:uppercase;letter-spacing:.06em;border-bottom:1px solid var(--border);text-align:left}
td{padding:14px 16px;border-bottom:1px solid var(--border);vertical-align:middle;transition:background-color 0.2s ease}
tr:last-child td{border-bottom:none}
tbody tr{transition:all 0.2s}
tbody tr:hover td{background:#f8faff}

.badge{display:inline-flex;align-items:center;gap:4px;padding:4px 12px;border-radius:99px;font-size:.75rem;font-weight:600;transition:all 0.25s}
.b-green{background:var(--green-bg);color:var(--green)}
.b-amber{background:var(--amber-bg);color:var(--amber)}
.b-red{background:var(--red-bg);color:var(--red)}
.b-blue{background:rgba(31,79,216,.06);color:var(--blue)}

.prog-wrap{background:#e8ecf7;border-radius:99px;height:8px;overflow:hidden;position:relative}
.prog-bar{height:100%;border-radius:99px;background:linear-gradient(90deg,var(--blue),var(--blue2));transition:width 0.6s cubic-bezier(0.4, 0, 0.2, 1)}
.prog-bar.done{background:linear-gradient(90deg,var(--green),#22c55e)}

.flash-container{position:fixed;top:24px;left:50%;transform:translateX(-50%);z-index:9999;display:flex;flex-direction:column;gap:10px;width:100%;max-width:400px;pointer-events:none}
.flash{padding:14px 20px;border-radius:12px;font-size:.875rem;display:flex;align-items:center;gap:12px;box-shadow:0 10px 30px rgba(0,0,0,.08);animation:slideDown 0.35s cubic-bezier(0.175, 0.885, 0.32, 1.125) forwards;pointer-events:auto}
.flash.success{background:#fff;color:var(--green);border-left:4px solid var(--green)}
.flash.error{background:#fff;color:var(--red);border-left:4px solid var(--red)}

.scan-field{font-size:1.05rem!important;padding:14px 20px!important;border:2px solid rgba(31,79,216,.2)!important;border-radius:14px!important}
.scan-field:focus{border-color:var(--blue)!important;box-shadow:0 0 0 5px rgba(31,79,216,.08)!important}

.toast-box{position:fixed;bottom:32px;right:32px;z-index:9999;display:flex;flex-direction:column;gap:12px;pointer-events:none}
.toast{background:#fff;border-radius:14px;padding:14px 22px;box-shadow:0 12px 36px rgba(0,0,0,.1);font-size:.875rem;font-weight:500;color:var(--ink2);display:flex;align-items:center;gap:12px;min-width:280px;border-left:4px solid var(--blue);animation:tUp .3s cubic-bezier(0.175, 0.885, 0.32, 1.1) forwards;pointer-events:auto}
.toast.ok{border-left-color:var(--green)}.toast.err{border-left-color:var(--red)}

.home-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:28px;margin-top:40px}
.hcard{background:#fff;border:1px solid var(--border);border-radius:24px;padding:40px 32px;display:flex;flex-direction:column;align-items:flex-start;gap:16px;transition:all .3s cubic-bezier(0.4, 0, 0.2, 1);box-shadow:0 4px 16px rgba(0,0,0,.02)}
.hcard:hover{transform:translateY(-6px);box-shadow:0 20px 40px rgba(31,79,216,.08);border-color:rgba(31,79,216,.2)}
.hcard-icon{width:64px;height:64px;border-radius:18px;display:flex;align-items:center;justify-content:center;font-size:28px;transition:transform 0.3s ease}
.hcard:hover .hcard-icon{transform:scale(1.1)}
.hcard-icon.b{background:linear-gradient(135deg,#eff6ff,#dbeafe)}
.hcard-icon.o{background:linear-gradient(135deg,#fff7ed,#ffedd5)}
.hcard h3{font-family:'Bebas Neue',sans-serif;font-size:1.6rem;letter-spacing:.04em;color:var(--ink2)}
.hcard p{font-size:.9rem;color:var(--muted);line-height:1.6}
.hcard .arr{margin-top:auto;font-size:.85rem;font-weight:600;color:var(--blue);display:flex;align-items:center;gap:4px}

.ws-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(290px,1fr));gap:20px}
.wsc{background:#fff;border:1.5px solid var(--border);border-radius:18px;padding:24px;transition:all .3s ease}
.wsc:hover{border-color:var(--blue);box-shadow:0 10px 25px var(--blue-gl);transform:translateY(-2px)}
.wsc h3{font-weight:700;font-size:1rem;margin-bottom:16px;display:flex;align-items:center;gap:10px}

.stat-row{display:flex;gap:16px;flex-wrap:wrap;margin-bottom:24px}
.stat{background:#fff;border:1px solid var(--border);border-radius:16px;padding:16px 24px;min-width:120px;flex:1;box-shadow:0 2px 10px rgba(0,0,0,.01)}
.stat-v{font-family:'Bebas Neue',sans-serif;font-size:2.1rem;color:var(--ink2)}
.stat-l{font-size:.72rem;color:var(--muted);font-weight:700;text-transform:uppercase;letter-spacing:.06em}

.order-row{cursor:pointer;transition:all 0.25s}
.order-row.active td{background:#f0f4ff!important}
.order-row.active td:first-child{border-left:4px solid var(--blue);border-top-left-radius:4px;border-bottom-left-radius:4px}
body.done{background:#f4fbf7}
body.done .nav{background:rgba(240,253,244,.85);border-color:#bbf7d0}

@keyframes fadeIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
@keyframes tUp{from{opacity:0;transform:translateY(20px)}to{opacity:1;transform:translateY(0)}}
@keyframes slideDown{from{opacity:0;transform:translate(-50%, -20px)}to{opacity:1;transform:translate(-50%, 0)}}

@media(max-width:768px){
  .nav{padding:0 20px}.wrap{padding:24px 16px}.card{padding:24px}
  .home-grid{grid-template-columns:1fr}
}
</style>
"""

FLASH_TEMPLATE = """
{% with messages = get_flashed_messages(with_categories=true) %}
  {% if messages %}
    <div class="flash-container">
      {% for category, message in messages %}
        <div class="flash {{ category }}">
          <span>{% if category == 'success' %}✅{% else %}❌{% endif %}</span>
          <span>{{ message }}</span>
        </div>
      {% endfor %}
    </div>
    <script>
      setTimeout(() => {
        const containers = document.querySelectorAll('.flash-container');
        containers.forEach(c => {
          c.style.opacity = '0';
          c.style.transition = 'opacity 0.4s ease';
          setTimeout(() => c.remove(), 400);
        });
      }, 4000);
    </script>
  {% endif %}
{% endwith %}
"""

TOAST_JS = """
<div class="toast-box" id="tb"></div>
<script>
function toast(msg,type){
  const b=document.getElementById('tb');
  const t=document.createElement('div');
  t.className='toast '+(type==='success'?'ok':type==='error'?'err':'');
  const ico={info:'ℹ️',success:'✅',error:'❌'};
  t.innerHTML='<span>'+(ico[type]||'ℹ️')+'</span><span>'+msg+'</span>';
  b.appendChild(t);
  setTimeout(()=>{
    t.style.opacity='0';
    t.style.transform='translateY(-10px)';
    t.style.transition='all .25s ease';
    setTimeout(()=>t.remove(),250)
  },3500);
}
</script>
"""

WSORDERS_T = '<!DOCTYPE html><html><head><title>{{ wname }} · '+BRAND+'</title>'+CSS+"""</head>
<body id="pgBody">
"""+FLASH_TEMPLATE+"""
<nav class="nav">
  <span class="nav-brand">⬡ """+BRAND+"""</span>
  <div class="nav-right">
    <span class="nav-chip">🏭 {{ wname }}</span>
    <a href="/workshop/logout" class="btn btn-ghost btn-sm">Выйти</a>
  </div>
</nav>
<div class="wrap">
  <div class="card" style="margin-bottom: 28px">
    <div class="card-title">📤 Создать новый заказ</div>
    <form method="POST" action="/workshop/upload" enctype="multipart/form-data">
      <div style="display:grid;grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));gap:20px;margin-bottom:20px">
        <div class="field" style="margin:0">
          <label>Код клиента</label>
          <input type="text" name="client_code" placeholder="KZ-2026-001" required>
        </div>
        <div class="field" style="margin:0">
          <label>Срок готовности</label>
          <input type="date" name="due_date">
        </div>
        <div class="field" style="margin:0">
          <label>Файл спецификации (.xlsx, .xls, .xml)</label>
          <input type="file" name="file" accept=".xlsx,.xls,.xml" required>
        </div>
      </div>
      <button type="submit" class="btn btn-blue">📤 Загрузить и распределить заказ</button>
    </form>
  </div>

  {% if orders %}
  <div class="stat-row">
    <div class="stat"><div class="stat-v">{{ orders|length }}</div><div class="stat-l">Всего заказов</div></div>
    <div class="stat"><div class="stat-v" id="sDone">—</div><div class="stat-l">Выполнено</div></div>
    <div class="stat"><div class="stat-v" id="sPct">—</div><div class="stat-l">Общий прогресс</div></div>
  </div>
  <div style="display:grid;grid-template-columns:repeat(auto-fit, minmax(340px, 1fr));gap:24px">
    <div class="card" style="padding:0;overflow:hidden">
      <div style="padding:20px 24px;border-bottom:1px solid var(--border)">
        <div class="card-title" style="margin:0">📦 Активные заказы</div>
      </div>
      <div style="overflow-y:auto;max-height:560px">
        <table>
          <thead><tr><th>Клиент</th><th>Прогресс</th><th>Статус</th></tr></thead>
          <tbody>
            {% for o in orders %}
            <tr class="order-row" id="orow-{{ o.id }}" onclick="selOrder({{ o.id }},'{{ o.client_code }}')">
              <td>
                <div style="font-weight:600;font-size:.9rem;color:var(--ink2)">{{ o.client_code }}</div>
                <div style="font-size:.75rem;color:var(--muted);margin-top:2px">{{ o.created_at[:10] }}</div>
              </td>
              <td style="min-width:120px">
                <div style="font-size:.74rem;color:var(--muted);margin-bottom:4px;font-family:'DM Mono',monospace">{{ o.scanned }}/{{ o.total }}</div>
                <div class="prog-wrap"><div class="prog-bar{% if o.pct==100 %} done{% endif %}" style="width:{{ o.pct }}%" id="rowbar-{{ o.id }}"></div></div>
              </td>
              <td id="rowbadge-{{ o.id }}">
                {% if o.pct==100 %}<span class="badge b-green">✓ Готово</span>
                {% elif o.pct>0 %}<span class="badge b-amber">⟳ В работе</span>
                {% else %}<span class="badge b-red">○ Ожидает</span>{% endif %}
              </td>
            </tr>
            {% endfor %}
          </tbody>
        </table>
      </div>
    </div>
    <div>
      <div class="card">
        <div class="card-title">🔍 Сканирование штрих-кодов</div>
        <div id="selLabel" style="color:var(--muted);font-size:.9rem;margin-bottom:18px">← Выберите заказ из списка слева для начала работы</div>
        <div id="progSec" style="display:none;margin-bottom:20px">
          <div style="display:flex;justify-content:space-between;margin-bottom:8px">
            <span style="font-size:.74rem;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.05em">Прогресс сборки</span>
            <span style="font-family:'DM Mono',monospace;font-size:.9rem;font-weight:600" id="pctTxt">0%</span>
          </div>
          <div class="prog-wrap" style="height:10px"><div id="mainBar" class="prog-bar" style="width:0%"></div></div>
        </div>
        <div id="scanSec" style="display:none">
          <div style="display:flex;gap:12px">
            <input type="text" id="scanInput" class="scan-field" placeholder="Считайте штрих-код сканером..." autocomplete="off">
            <button class="btn btn-blue" onclick="doScan()">Ввод</button>
          </div>
        </div>
      </div>
      <div class="card" style="margin-top:24px;padding:0;overflow:hidden">
        <div style="padding:16px 24px;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;align-items:center">
          <div class="card-title" style="margin:0">📋 Состав спецификации</div>
          <span id="dcnt" style="font-size:.8rem;color:var(--muted);font-weight:500"></span>
        </div>
        <div style="overflow-y:auto;max-height:380px">
          <table>
            <thead><tr><th>Код</th><th>Обозначение</th><th>Наименование</th><th>План</th><th>Факт</th><th>Статус</th></tr></thead>
            <tbody id="dtbody">
              <tr><td colspan="6" style="text-align:center;color:var(--muted);padding:40px;font-size:.875rem">Спецификация не загружена</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  </div>
  {% else %}
  <div style="text-align:center;padding:100px 24px">
    <div style="font-size:80px;margin-bottom:20px">📭</div>
    <h2 style="color:var(--ink2);font-weight:600;font-size:1.4rem">Производственный план пуст</h2>
    <p style="color:var(--muted);margin-top:8px;font-size:.95rem">Загрузите первый файл через форму распределения выше</p>
  </div>
  {% endif %}
</div>"""+TOAST_JS+"""
<script>
let cid=null;
(function(){
  let d=0;
  const rows=document.querySelectorAll('.order-row');
  rows.forEach(r=>{
    const b=r.querySelector('.badge');
    if(b&&b.classList.contains('b-green'))d++;
  });
  document.getElementById('sDone').textContent=d;
})();

function selOrder(id,label){
  cid=id;
  document.querySelectorAll('.order-row').forEach(r=>r.classList.remove('active'));
  const row=document.getElementById('orow-'+id);
  if(row)row.classList.add('active');
  document.getElementById('selLabel').innerHTML='Текущий заказ: <strong style="color:var(--blue)">'+label+'</strong>';
  document.getElementById('progSec').style.display='block';
  document.getElementById('scanSec').style.display='block';
  refreshStatus();
  setTimeout(()=>document.getElementById('scanInput').focus(),80);
}

function doScan(){
  const inp=document.getElementById('scanInput');
  const code=inp.value.trim();inp.value='';inp.focus();
  if(!code||!cid){toast('Не выбран активный заказ','error');return;}
  fetch('/api/scan/'+cid,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({code})})
    .then(r=>r.json()).then(d=>{
      if(d.ok){toast('Компонент '+code+' успешно учтен ('+d.scanned+'/'+d.planned+')','success');refreshStatus();}
      else toast(d.error||'Ошибка сканирования','error');
    }).catch(()=>toast('Сбой сетевого соединения','error'));
}
document.getElementById('scanInput')?.addEventListener('keydown',e=>{if(e.key==='Enter')doScan();});

function refreshStatus(){
  if(!cid)return;
  fetch('/api/status/'+cid).then(r=>r.json()).then(d=>{
    const pct=d.progress_pct;
    const bar=document.getElementById('mainBar');
    bar.style.width=pct+'%';
    bar.className='prog-bar'+(pct===100?' done':'');
    document.getElementById('pctTxt').textContent=pct+'%';
    document.getElementById('sPct').textContent=pct+'%';
    if(pct===100)document.getElementById('pgBody').classList.add('done');
    else document.getElementById('pgBody').classList.remove('done');

    const rb=document.getElementById('rowbar-'+cid);
    if(rb){rb.style.width=pct+'%';rb.className='prog-bar'+(pct===100?' done':'');}
    const rbd=document.getElementById('rowbadge-'+cid);
    if(rbd){
      if(pct===100)rbd.innerHTML='<span class="badge b-green">✓ Готово</span>';
      else if(pct>0)rbd.innerHTML='<span class="badge b-amber">⟳ В работе</span>';
      else rbd.innerHTML='<span class="badge b-red">○ Ожидает</span>';
    }

    const tbody=document.getElementById('dtbody');
    tbody.innerHTML='';
    document.getElementById('dcnt').textContent=d.items.length+' позиций';
    d.items.forEach(it=>{
      const tr=document.createElement('tr');
      let bdg='<span class="badge b-red">○</span>';
      if(it.scanned_qty>=it.planned_qty)bdg='<span class="badge b-green">✓</span>';
      else if(it.scanned_qty>0)bdg='<span class="badge b-amber">'+it.scanned_qty+'/'+it.planned_qty+'</span>';
      tr.innerHTML='<td><code style="font-family:\\'DM Mono\\',monospace;font-size:.77rem;background:#f0f3f9;padding:3px 8px;border-radius:6px;color:var(--ink3)">'+it.item_code+'</code></td>'
        +'<td style="max-width:110px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:.8rem;color:var(--muted)" title="'+(it.designation||'')+'">'+(it.designation||'—')+'</td>'
        +'<td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-weight:500" title="'+it.item_name+'">'+it.item_name+'</td>'
        +'<td style="font-family:\\'DM Mono\\',monospace;font-weight:600">'+it.planned_qty+'</td>'
        +'<td style="font-family:\\'DM Mono\\',monospace;color:var(--blue);font-weight:600">'+it.scanned_qty+'</td>'
        +'<td>'+bdg+'</td>';
      tbody.appendChild(tr);
    });
  });
}
setInterval(refreshStatus,4000);
</script>
</body></html>"""

HOME_T = '<!DOCTYPE html><html><head><title>'+BRAND+'</title>'+CSS+"""</head><body>
<nav class="nav"><span class="nav-brand">⬡ """+BRAND+"""</span></nav>
<div class="wrap">
  <div style="text-align:center;padding:70px 0 30px">
    <div style="font-family:'Bebas Neue',sans-serif;font-size:3.8rem;letter-spacing:.06em;
         background:linear-gradient(135deg,#0b0e14 20%,var(--blue));-webkit-background-clip:text;
         background-clip:text;color:transparent;line-height:1.1">"""+BRAND+"""</div>
    <p style="color:var(--muted);margin-top:12px;font-size:1.05rem;letter-spacing:0.02em">Цифровая экосистема операционного контроля мебельного производства</p>
  </div>
  <div class="home-grid">
    <a href="/workshop/select" class="hcard">
      <div class="hcard-icon b">🏭</div>
      <div>
        <h3>Производственные цеха</h3>
        <p>Авторизация на рабочих местах сборки, мониторинг текущих планов и сканирование упаковочных листов.</p>
      </div>
      <div class="arr">Перейти в терминал →</div>
    </a>
    <a href="/admin/login" class="hcard">
      <div class="hcard-icon o">🔐</div>
      <div>
        <h3>Управление системой</h3>
        <p>Панель суперпользователя. Конфигурация структуры фабрики, менеджмент прав доступа и аналитика.</p>
      </div>
      <div class="arr">Авторизация менеджера →</div>
    </a>
  </div>
</div>
<div style="text-align:center;padding:32px;color:var(--muted);font-size:.8rem;margin-top:64px;border-top:1px solid var(--border)">
  © 2026 """+BRAND+""" · Enterprise Production Control System
</div></body></html>"""

WSSELECT_T = '<!DOCTYPE html><html><head><title>Выбор подразделения · '+BRAND+'</title>'+CSS+"""</head><body>
"""+FLASH_TEMPLATE+"""
<nav class="nav">
  <span class="nav-brand">⬡ """+BRAND+"""</span>
  <div class="nav-right"><a href="/" class="btn btn-ghost btn-sm">← На главную</a></div>
</nav>
<div class="wrap">
  <div style="margin-bottom:36px">
    <div style="font-family:'Bebas Neue',sans-serif;font-size:2.4rem;letter-spacing:.05em;color:var(--ink2)">🏭 Производственные участки</div>
    <p style="color:var(--muted);margin-top:6px">Выберите терминал назначения и введите ключ доступа</p>
  </div>
  <div class="ws-grid">
    {% for w in workshops %}
    <div class="wsc">
      <h3>🏭 {{ w.name }}</h3>
      <form method="POST" action="/workshop/login">
        <input type="hidden" name="workshop_name" value="{{ w.name }}">
        <div class="field"><label>Пароль участка</label>
          <input type="password" name="password" placeholder="••••••" required autocomplete="off">
        </div>
        <label class="check-row"><input type="checkbox" name="remember" value="1"> Сохранить сессию</label>
        <button type="submit" class="btn btn-blue" style="width:100%">Активировать терминал</button>
      </form>
    </div>
    {% endfor %}
    {% if not workshops %}
    <div style="color:var(--muted);text-align:center;padding:60px 20px;grid-column:1/-1">
      <div style="font-size:56px;margin-bottom:16px">🏗</div>
      Активные производственные линии не зарегистрированы в системе.
    </div>
    {% endif %}
  </div>
</div></body></html>"""

ADMINLOGIN_T = '<!DOCTYPE html><html><head><title>Авторизация · '+BRAND+'</title>'+CSS+"""</head><body>
"""+FLASH_TEMPLATE+"""
<nav class="nav">
  <span class="nav-brand">⬡ """+BRAND+"""</span>
  <div class="nav-right"><a href="/" class="btn btn-ghost btn-sm">← Главная</a></div>
</nav>
<div class="wrap-narrow">
  <div class="card">
    <div style="text-align:center;margin-bottom:32px">
      <div style="font-size:44px;margin-bottom:14px">🔐</div>
      <div style="font-family:'Bebas Neue',sans-serif;font-size:2rem;letter-spacing:.05em;color:var(--ink2)">Контроль доступа</div>
      <p style="color:var(--muted);font-size:.9rem;margin-top:4px">Панель управления фабрикой</p>
    </div>
    <form method="POST">
      <div class="field"><label>Идентификатор</label><input type="text" name="username" autofocus required placeholder="Логин"></div>
      <div class="field"><label>Пароль</label><input type="password" name="password" required placeholder="••••••"></div>
      <label class="check-row"><input type="checkbox" name="remember" value="1"> Запомнить сессию на 30 дней</label>
      <button type="submit" class="btn btn-blue" style="width:100%;padding:14px">Выполнить вход</button>
    </form>
  </div>
</div></body></html>"""

ADMINPANEL_T = '<!DOCTYPE html><html><head><title>Панель администратора · '+BRAND+'</title>'+CSS+"""</head><body>
"""+FLASH_TEMPLATE+"""
<nav class="nav">
  <span class="nav-brand">⬡ """+BRAND+"""</span>
  <div class="nav-right">
    <span class="nav-chip">🔐 Администратор</span>
    <a href="/admin/logout" class="btn btn-ghost btn-sm">Выйти</a>
  </div>
</nav>
<div class="wrap">
  <div class="card">
    <div class="card-title">🏭 Конфигурация производственных участков</div>
    <form method="POST" action="/admin/workshop/create" style="display:flex;gap:16px;margin-bottom:28px;align-items:flex-end;flex-wrap:wrap">
      <div class="field" style="margin:0;flex:2;min-width:180px"><label>Название участка</label>
        <input type="text" name="name" placeholder="Например, Сборочный цех №4" required>
      </div>
      <div class="field" style="margin:0;flex:1;min-width:140px"><label>Кредо-пароль</label>
        <input type="password" name="password" placeholder="••••••" required>
      </div>
      <button type="submit" class="btn btn-blue btn-sm">＋ Добавить участок</button>
    </form>
    {% if workshops %}
    <div class="tbl-wrap">
      <table>
        <thead><tr><th>ID</th><th>Производственная линия</th><th>Заказов в базе</th><th>Управление параметрами</th></tr></thead>
        <tbody>
        {% for w in workshops %}
        <tr>
          <td style="color:var(--muted);font-family:'DM Mono',monospace;font-size:.8rem">{{ w.id }}</td>
          <td><strong style="color:var(--ink2)">{{ w.name }}</strong></td>
          <td><span class="badge b-blue">{{ w.order_count }} шт.</span></td>
          <td>
            <div style="display:flex;gap:10px;flex-wrap:wrap;align-items:center">
              <button class="btn btn-outline btn-xs" onclick="toggleRP({{ w.id }})">🔑 Обновить пароль</button>
              <form method="POST" action="/admin/workshop/{{ w.id }}/delete" style="margin:0" onsubmit="return confirm('Удалить участок «{{ w.name }}» и связанные заказы без возможности восстановления?')">
                <button type="submit" class="btn btn-danger btn-xs">🗑 Удалить линию</button>
              </form>
            </div>
            <div id="rp-{{ w.id }}" style="display:none;margin-top:14px;animation: fadeIn 0.25s ease">
              <form method="POST" action="/admin/workshop/{{ w.id }}/reset" style="display:flex;gap:10px;flex-wrap:wrap">
                <input type="password" name="new_password" placeholder="Новый токен доступа" required style="flex:1;min-width:140px">
                <button type="submit" class="btn btn-blue btn-xs">Сохранить</button>
                <button type="button" class="btn btn-ghost btn-xs" onclick="toggleRP({{ w.id }})">Отмена</button>
              </form>
            </div>
          </td>
        </tr>
        {% endfor %}
        </tbody>
      </table>
    </div>
    {% else %}
    <div style="text-align:center;padding:40px;color:var(--muted)">Зарегистрированные производственные участки отсутствуют</div>
    {% endif %}
  </div>

  <div class="card">
    <div class="card-title">🔑 Безопасность учетной записи администратора</div>
    <form method="POST" action="/admin/change_password" style="display:flex;gap:16px;flex-wrap:wrap;align-items:flex-end">
      <div class="field" style="margin:0;flex:1;min-width:180px"><label>Новый пароль</label>
        <input type="password" name="new_password" placeholder="••••••" required>
      </div>
      <div class="field" style="margin:0;flex:1;min-width:180px"><label>Подтверждение пароля</label>
        <input type="password" name="confirm_password" placeholder="••••••" required>
      </div>
      <button type="submit" class="btn btn-outline btn-sm">Обновить учетные данные</button>
    </form>
  </div>
</div>
<script>
function toggleRP(id){
  const el=document.getElementById('rp-'+id);
  el.style.display=el.style.display==='none'?'block':'none';
}
</script>
</body></html>"""

@app.route('/')
def home():
    return render_template_string(HOME_T)

@app.route('/workshop/select')
def workshop_select():
    ws = Workshop.query.order_by(Workshop.name).all()
    return render_template_string(WSSELECT_T, workshops=ws)

@app.route('/workshop/login', methods=['POST'])
def workshop_login():
    name = request.form.get('workshop_name','').strip()
    pw = request.form.get('password','')
    rem = request.form.get('remember')=='1'
    w = Workshop.query.filter_by(name=name).first()
    if w and hashlib.sha256(pw.encode()).hexdigest() == w.password_hash:
        if rem: 
            session.permanent = True
        session['workshop'] = name
        return redirect('/workshop/orders')
    flash('Неверный ключ доступа для выбранного участка', 'error')
    return redirect('/workshop/select')

@app.route('/workshop/orders')
def workshop_orders():
    if 'workshop' not in session: 
        return redirect('/workshop/select')
    wname = session['workshop']
    orders = orders_for(wname)
    return render_template_string(WSORDERS_T, wname=wname, orders=orders)

@app.route('/workshop/upload', methods=['POST'])
def workshop_upload():
    if 'workshop' not in session:
        flash('Требуется авторизация в терминале цеха', 'error')
        return redirect('/workshop/select')
    
    wname = session['workshop']
    cc = request.form.get('client_code', '').strip()
    dd = request.form.get('due_date', '').strip() or None
    f = request.files.get('file')
    
    if not cc or not f or not f.filename:
        flash('Не заполнен код контрагента или отсутствует файл', 'error')
        return redirect('/workshop/orders')
    
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ('.xlsx', '.xls', '.xml'):
        flash(f'Запрещенный тип файла: {ext}. Допустимы только .xlsx, .xls и .xml', 'error')
        return redirect('/workshop/orders')
    
    if not Workshop.query.filter_by(name=wname).first():
        flash('Ошибка верификации производственного участка', 'error')
        return redirect('/workshop/select')
    
    tmp = os.path.join(tempfile.gettempdir(), f'itvms_{int(datetime.now().timestamp())}{ext}')
    f.save(tmp)
    
    try:
        items = parse_file(tmp)
    except Exception as e:
        db.session.rollback()
        flash(str(e), 'error')
        return redirect('/workshop/orders')
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    
    try:
        order = Order(
            client_code=cc,
            due_date=dd,
            created_at=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            source_file=f.filename,
            workshop_name=wname
        )
        db.session.add(order)
        db.session.flush()
        
        for it in items:
            db.session.add(OrderItem(
                order_id=order.id,
                item_code=it['code'],
                item_name=it['name'],
                planned_qty=it['qty'],
                designation=it.get('designation', '')
            ))
        
        db.session.commit()
        flash(f'Заказ успешно сформирован. Интегрировано {len(items)} компонентов для участка «{wname}»', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Критическая ошибка сохранения спецификации в СУБД: {e}', 'error')
    
    return redirect('/workshop/orders')

@app.route('/workshop/logout')
def workshop_logout():
    session.pop('workshop', None)
    return redirect('/workshop/select')

@app.route('/admin/login', methods=['GET','POST'])
def admin_login():
    global ADMIN_HASH
    if admin_ok(): 
        return redirect('/admin')
    if request.method == 'POST':
        uname = request.form.get('username','').strip()
        pw = request.form.get('password','')
        rem = request.form.get('remember')=='1'
        cur_hash = session.get('admin_hash', ADMIN_HASH)
        if uname == ADMIN_USER and hashlib.sha256(pw.encode()).hexdigest() == cur_hash:
            if rem: 
                session.permanent = True
            session['admin'] = True
            session['admin_hash'] = cur_hash
            return redirect('/admin')
        flash('Аутентификация отклонена: неверный логин или пароль', 'error')
    return render_template_string(ADMINLOGIN_T)

@app.route('/admin/logout')
def admin_logout():
    session.pop('admin', None)
    session.pop('admin_hash', None)
    return redirect('/')

@app.route('/admin')
def admin_panel():
    if not admin_ok(): 
        return redirect('/admin/login')
    ws = Workshop.query.order_by(Workshop.name).all()
    workshop_data = []
    for w in ws:
        cnt = Order.query.filter_by(workshop_name=w.name).count()
        workshop_data.append({'id': w.id, 'name': w.name, 'order_count': cnt})
    return render_template_string(ADMINPANEL_T, workshops=workshop_data)

@app.route('/admin/workshop/create', methods=['POST'])
def admin_create_workshop():
    if not admin_ok(): 
        return redirect('/admin/login')
    name = request.form.get('name', '').strip()
    pw = request.form.get('password', '')
    if not name or not pw:
        flash('Все поля обязательны к заполнению', 'error')
        return redirect('/admin')
    if Workshop.query.filter_by(name=name).first():
        flash(f'Участок «{name}» уже зарегистрирован', 'error')
        return redirect('/admin')
    db.session.add(Workshop(name=name, password_hash=hashlib.sha256(pw.encode()).hexdigest()))
    db.session.commit()
    flash(f'Производственный участок «{name}» добавлен в систему', 'success')
    return redirect('/admin')

@app.route('/admin/workshop/<int:wid>/delete', methods=['POST'])
def admin_delete_workshop(wid):
    if not admin_ok(): 
        return redirect('/admin/login')
    w = Workshop.query.get(wid)
    if w:
        db.session.delete(w)
        db.session.commit()
        flash(f'Участок «{w.name}» деактивирован и удален из реестра', 'success')
    else:
        flash('Объект конфигурации не найден', 'error')
    return redirect('/admin')

@app.route('/admin/workshop/<int:wid>/reset', methods=['POST'])
def admin_reset_workshop_password(wid):
    if not admin_ok(): 
        return redirect('/admin/login')
    pw = request.form.get('new_password', '')
    if not pw:
        flash('Поле пароля не может быть пустым', 'error')
        return redirect('/admin')
    w = Workshop.query.get(wid)
    if w:
        w.password_hash = hashlib.sha256(pw.encode()).hexdigest()
        db.session.commit()
        flash(f'Токен безопасности для участка «{w.name}» изменен', 'success')
    else:
        flash('Участок не обнаружен', 'error')
    return redirect('/admin')

@app.route('/admin/change_password', methods=['POST'])
def admin_change_password():
    global ADMIN_HASH
    if not admin_ok(): 
        return redirect('/admin/login')
    pw = request.form.get('new_password', '')
    c2 = request.form.get('confirm_password', '')
    if not pw:
        flash('Пароль не должен быть пустым', 'error')
        return redirect('/admin')
    if pw != c2:
        flash('Введенные пароли не совпадают', 'error')
        return redirect('/admin')
    ADMIN_HASH = hashlib.sha256(pw.encode()).hexdigest()
    session['admin_hash'] = ADMIN_HASH
    flash('Глобальный мастер-пароль успешно обновлен', 'success')
    return redirect('/admin')

@app.route('/api/scan/<int:oid>', methods=['POST'])
def api_scan(oid):
    data = request.get_json(silent=True) or {}
    code = str(data.get('code', '')).strip()
    if not code:
        return jsonify({'ok': False, 'error': 'Идентификатор кода пуст'})
    item = OrderItem.query.filter_by(order_id=oid, item_code=code).first()
    if not item:
        return jsonify({'ok': False, 'error': f'Компонент «{code}» отсутствует в спецификации'})
    if item.scanned_qty >= item.planned_qty:
        return jsonify({'ok': False, 'error': f'Компонент «{code}» уже полностью собран'})
    item.scanned_qty += 1
    db.session.commit()
    return jsonify({'ok': True, 'scanned': item.scanned_qty, 'planned': item.planned_qty})

@app.route('/api/status/<int:oid>')
def api_status(oid):
    items = OrderItem.query.filter_by(order_id=oid).all()
    total = sum(i.planned_qty for i in items)
    scanned = sum(i.scanned_qty for i in items)
    pct = int(scanned / total * 100) if total else 0
    return jsonify({
        'progress_pct': pct,
        'total_planned': total,
        'total_scanned': scanned,
        'items': [{
            'item_code': i.item_code,
            'item_name': i.item_name,
            'planned_qty': i.planned_qty,
            'scanned_qty': i.scanned_qty,
            'designation': i.designation or ''
        } for i in items]
    })

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
