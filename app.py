# -*- coding: utf-8 -*-
"""
山羊育种管理平台 - Goat Breeding Management Platform
Flask + SQLite + Bootstrap 5
"""

import os
import sys
import re
import io
import json
import secrets
import hashlib
from datetime import datetime, timedelta
from functools import wraps

from flask import (Flask, render_template, request, redirect, url_for,
                   session, jsonify, send_file, flash)
from werkzeug.utils import secure_filename
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill


def resource_base():
    """打包后资源（templates/static/seed）所在目录。
    onefile 模式：PyInstaller 解压到 sys._MEIPASS；
    onedir 模式：可执行文件所在目录；开发模式：脚本所在目录。
    """
    if getattr(sys, 'frozen', False):
        return getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


# ──────────────────────────────────────────────────────────
# 应用初始化
# ──────────────────────────────────────────────────────────
app = Flask(__name__,
            template_folder=os.path.join(resource_base(), 'templates'),
            static_folder=os.path.join(resource_base(), 'static'))
def _load_secret_key():
    """会话签名密钥：优先环境变量 SHEEP_SECRET_KEY；
    否则在用户数据目录持久化随机密钥（不在仓库内）；
    均不可用则使用临时随机密钥（重启后会话失效）。"""
    env_key = os.environ.get('SHEEP_SECRET_KEY')
    if env_key:
        return env_key
    key_file = os.path.join(DATA_DIR, '.secret_key')
    try:
        if os.path.exists(key_file):
            with open(key_file, 'r', encoding='utf-8') as f:
                return f.read().strip()
        key = secrets.token_hex(32)
        with open(key_file, 'w', encoding='utf-8') as f:
            f.write(key)
        return key
    except OSError:
        return secrets.token_hex(32)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# 桌面端：数据目录优先使用环境变量 SHEEP_DATA_DIR（由 launcher 注入，指向 %APPDATA%）；
# 缺省回退脚本目录（开发模式）。保证打包后业务数据落在用户可写目录，可持久化。
DATA_DIR = os.environ.get('SHEEP_DATA_DIR', BASE_DIR)
os.makedirs(DATA_DIR, exist_ok=True)
DB_PATH = os.path.join(DATA_DIR, 'sheep_farm.db')
app.secret_key = _load_secret_key()
UPLOAD_FOLDER = os.path.join(DATA_DIR, 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ──────────────────────────────────────────────────────────
# 数据库
# ──────────────────────────────────────────────────────────
import sqlite3

def get_db():
    conn = sqlite3.connect(DB_PATH, timeout=10)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=5000")
    conn.execute("PRAGMA foreign_keys = ON")
    conn.row_factory = sqlite3.Row
    return conn

def ensure_breed(conn, name):
    """导入或录入品种时，确保品种名已登记进 breeds 表（避免下拉选不到导入进来的品种）。"""
    if name and str(name).strip():
        conn.execute("INSERT OR IGNORE INTO breeds (name) VALUES (?)", (str(name).strip(),))

def init_db():
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS sheep (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ear_tag TEXT UNIQUE NOT NULL,
            electronic_ear_tag TEXT DEFAULT '',
            breed TEXT DEFAULT '努比亚山羊',
            birth_date TEXT DEFAULT '',
            sex TEXT DEFAULT '',  -- 允许空（性别缺失时入库，不计入公/母统计）
            source TEXT DEFAULT '',
            status TEXT DEFAULT '在群' CHECK(status IN ('在群','淘汰','死亡','出售')),
            status_reason TEXT DEFAULT '',
            father_tag TEXT DEFAULT '',
            mother_tag TEXT DEFAULT '',
            weak INTEGER DEFAULT 0,
            card_downloaded_at TIMESTAMP,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS breeding_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ewe_tag TEXT NOT NULL,
            ram_tag TEXT DEFAULT '',
            mating_date TEXT DEFAULT '',
            lambing_date TEXT NOT NULL,
            total_born INTEGER DEFAULT 0,
            live_born INTEGER DEFAULT 0,
            remark TEXT DEFAULT '',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS lambs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            breeding_record_id INTEGER NOT NULL,
            lamb_ear_tag TEXT NOT NULL,
            sex TEXT DEFAULT '',
            birth_weight REAL DEFAULT 0,
            FOREIGN KEY (breeding_record_id) REFERENCES breeding_records(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS measurements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sheep_id INTEGER NOT NULL,
            stage TEXT NOT NULL CHECK(stage IN ('初生','3月龄','6月龄','12月龄')),
            weight REAL DEFAULT 0,
            body_height REAL DEFAULT 0,
            body_length REAL DEFAULT 0,
            chest_girth REAL DEFAULT 0,
            cannon_circumference REAL DEFAULT 0,
            hip_width REAL DEFAULT 0,
            testis_circumference REAL DEFAULT 0,
            testis_diameter REAL DEFAULT 0,
            measure_date TEXT DEFAULT '',
            judgment TEXT DEFAULT '正常' CHECK(judgment IN ('正常','弱羊')),
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (sheep_id) REFERENCES sheep(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS breeds (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE INDEX IF NOT EXISTS idx_sheep_ear_tag ON sheep(ear_tag);
        CREATE INDEX IF NOT EXISTS idx_sheep_sex ON sheep(sex);
        CREATE INDEX IF NOT EXISTS idx_lambs_breeding ON lambs(breeding_record_id);
        CREATE INDEX IF NOT EXISTS idx_measurements_sheep ON measurements(sheep_id);
        CREATE INDEX IF NOT EXISTS idx_measurements_stage ON measurements(stage);
    """)

    # 兼容已有库：新增字段迁移
    def col_exists(table, col):
        cols = [r['name'] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()]
        return col in cols

    if not col_exists('sheep', 'weak'):
        conn.execute("ALTER TABLE sheep ADD COLUMN weak INTEGER DEFAULT 0")
    if not col_exists('measurements', 'judgment'):
        conn.execute("ALTER TABLE measurements ADD COLUMN judgment TEXT DEFAULT '正常'")
    if not col_exists('sheep', 'status_reason'):
        conn.execute("ALTER TABLE sheep ADD COLUMN status_reason TEXT DEFAULT ''")
    if not col_exists('sheep', 'card_downloaded_at'):
        conn.execute("ALTER TABLE sheep ADD COLUMN card_downloaded_at TIMESTAMP")
    # 依据已有测定记录回灌 weak 标签
    conn.execute("""UPDATE sheep SET weak=1 WHERE id IN
        (SELECT DISTINCT sheep_id FROM measurements WHERE judgment='弱羊')""")

    # 默认品种
    default_breeds = ['努比亚山羊', '都安山羊', '波尔山羊', '隆林山羊', '湖羊', '小尾寒羊']
    for b in default_breeds:
        conn.execute("INSERT OR IGNORE INTO breeds (name) VALUES (?)", (b,))

    conn.commit()
    conn.close()

init_db()

# ──────────────────────────────────────────────────────────
# 认证装饰器
# ──────────────────────────────────────────────────────────
def login_required(f):
    return f

def admin_required(f):
    return f

def super_admin_required(f):
    return f

# ──────────────────────────────────────────────────────────
# 页面路由
# ──────────────────────────────────────────────────────────
@app.route('/')
def index():
    return redirect(url_for('dashboard'))

@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html')

@app.route('/sheep')
@login_required
def sheep_list():
    return render_template('sheep_list.html')

@app.route('/sheep/add')
@admin_required
def sheep_add():
    return render_template('sheep_form.html')

@app.route('/sheep/<int:sheep_id>/edit')
@admin_required
def sheep_edit(sheep_id):
    return render_template('sheep_form.html', sheep_id=sheep_id)

@app.route('/sheep/<int:sheep_id>/card')
@login_required
def sheep_card_page(sheep_id):
    return render_template('sheep_card.html', sheep_id=sheep_id)

@app.route('/breeding')
@login_required
def breeding_list():
    return render_template('breeding_list.html')

@app.route('/breeding/add')
@admin_required
def breeding_add():
    return render_template('breeding_form.html')

@app.route('/breeding/<int:rec_id>/edit')
@admin_required
def breeding_edit(rec_id):
    return render_template('breeding_form.html', rec_id=rec_id)

@app.route('/measurement')
@login_required
def measurement_list():
    return render_template('measurement_list.html')

@app.route('/measurement/add')
@admin_required
def measurement_add():
    conn = get_db()
    breeds = conn.execute("SELECT name FROM breeds ORDER BY name").fetchall()
    conn.close()
    return render_template('measurement_form.html', breeds=breeds)

@app.route('/import')
@admin_required
def import_page():
    return render_template('import.html')

# ──────────────────────────────────────────────────────────
# 看板 API
# ──────────────────────────────────────────────────────────
@app.route('/api/dashboard')
@login_required
def api_dashboard():
    conn = get_db()

    # 种公羊：后代羊只的父亲（在配种记录中出现过的公羊，去重）
    ram_count = conn.execute("SELECT COUNT(DISTINCT ram_tag) FROM breeding_records WHERE ram_tag IS NOT NULL AND ram_tag != ''").fetchone()[0]
    # 基础母羊：已经生仔过的母羊（在配种记录中出现过的母羊，去重）
    ewe_count = conn.execute("SELECT COUNT(DISTINCT ewe_tag) FROM breeding_records WHERE ewe_tag IS NOT NULL AND ewe_tag != ''").fetchone()[0]
    # 总羊只
    total_sheep = conn.execute("SELECT COUNT(*) FROM sheep WHERE status='在群'").fetchone()[0]

    # 测定统计
    measurements = conn.execute("""
        SELECT stage, COUNT(*) as cnt, ROUND(AVG(weight), 2) as avg_weight
        FROM measurements m JOIN sheep s ON m.sheep_id = s.id
        WHERE s.status = '在群'
        GROUP BY stage
    """).fetchall()

    # 配种/产羔统计
    breeding_total = conn.execute("SELECT COUNT(*) FROM breeding_records").fetchone()[0]
    lamb_total = conn.execute("SELECT COUNT(*) FROM lambs").fetchone()[0]

    # 性别分布
    sex_dist = conn.execute("""
        SELECT sex, COUNT(*) as cnt FROM sheep WHERE status='在群' GROUP BY sex
    """).fetchall()

    # 品种分布
    breed_dist = conn.execute("""
        SELECT breed, sex, COUNT(*) as cnt FROM sheep
        WHERE status='在群' AND breed IS NOT NULL AND breed != ''
        GROUP BY breed, sex ORDER BY breed, sex
    """).fetchall()
    
    breed_totals = conn.execute("""
        SELECT breed, COUNT(*) as total FROM sheep
        WHERE status='在群' AND breed IS NOT NULL AND breed != ''
        GROUP BY breed ORDER BY total DESC
    """).fetchall()

    # 羊群结构：繁殖公羊(在群且配过种的公) / 繁殖母羊(在群且产过羔的母) / 育成羊(其余在群)
    ram_breeding = conn.execute("""
        SELECT COUNT(DISTINCT s.id) FROM sheep s
        INNER JOIN breeding_records br ON br.ram_tag = s.ear_tag
        WHERE s.status='在群' AND s.sex='公'
    """).fetchone()[0]
    ewe_breeding = conn.execute("""
        SELECT COUNT(DISTINCT s.id) FROM sheep s
        INNER JOIN breeding_records br ON br.ewe_tag = s.ear_tag
        WHERE s.status='在群' AND s.sex='母'
    """).fetchone()[0]

    conn.close()

    stage_stats = {}
    for m in measurements:
        stage_stats[m['stage']] = {'count': m['cnt'], 'avg_weight': m['avg_weight']}

    # 品种分布（按品种统计公/母）
    breed_map = {}
    for r in breed_dist:
        b = r['breed']
        breed_map.setdefault(b, {'breed': b, '公': 0, '母': 0, 'total': 0})
        breed_map[b][r['sex']] = r['cnt']
        breed_map[b]['total'] += r['cnt']
    breed_stats = list(breed_map.values())

    flock_structure = {
        'ram_breeding': ram_breeding,
        'ewe_breeding': ewe_breeding,
        'young': total_sheep - ram_breeding - ewe_breeding
    }

    # 近12月产羔趋势（简化）
    conn2 = get_db()
    lambing_trend = conn2.execute("""
        SELECT substr(lambing_date,1,7) as month, COUNT(*) as cnt
        FROM breeding_records
        WHERE lambing_date >= date('now','-12 months')
        GROUP BY month ORDER BY month
    """).fetchall()
    conn2.close()

    return jsonify({
        'ram_count': ram_count,
        'ewe_count': ewe_count,
        'total_sheep': total_sheep,
        'breeding_total': breeding_total,
        'lamb_total': lamb_total,
        'stage_stats': stage_stats,
        'sex_dist': [dict(r) for r in sex_dist],
        'breed_stats': breed_stats,
        'flock_structure': flock_structure,
        'lambing_trend': [dict(r) for r in lambing_trend]
    })

# ──────────────────────────────────────────────────────────
# 测定预警 API
# ──────────────────────────────────────────────────────────
def _add_months(d, months):
    """日期 d 加 months 个月，day 取合理值"""
    m = d.month - 1 + months
    y = d.year + m // 12
    m = m % 12 + 1
    # 每月天数（处理闰年2月）
    leap = (y % 4 == 0 and y % 100 != 0) or (y % 400 == 0)
    days_in_month = [31, 29 if leap else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1]
    return datetime(y, m, min(d.day, days_in_month)).date()

def _parse_date(s):
    if not s:
        return None
    for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%Y-%m-%d %H:%M:%S'):
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except ValueError:
            continue
    return None

@app.route('/api/measurement_alerts')
@login_required
def api_measurement_alerts():
    """测定提醒：
    - 仅统计在群且非弱羊
    - 提醒窗口：当月达到待测月龄的羊 + 上月需测但未测的羊
      （到期日落在 [上月1日, 本月末] 且未测定该阶段）
    - 到期即提醒：窗口内任一未测阶段都提醒，不因早期漏测而隐藏
    - 对每条提醒标注“缺测的早期阶段”(missing_early)，便于先补前面的阶段
    - 另列 history_missing：早于窗口且仍未测的阶段（历史漏测待补测清单）
    - 仅做提醒，不执行任何淘汰操作
    """
    stages = [('初生', 0), ('3月龄', 3), ('6月龄', 6), ('12月龄', 12)]
    stage_idx = {st: i for i, (st, _) in enumerate(stages)}
    today = datetime.now().date()

    # 当前月与上月边界（窗口 = 上月整月 + 本月整月）
    if today.month == 1:
        prev_y, prev_m = today.year - 1, 12
    else:
        prev_y, prev_m = today.year, today.month - 1
    prev_start = datetime(prev_y, prev_m, 1).date()
    if today.month == 12:
        next_y, next_m = today.year + 1, 1
    else:
        next_y, next_m = today.year, today.month + 1
    cur_end = datetime(next_y, next_m, 1).date() - timedelta(days=1)

    conn = get_db()
    sheep_rows = conn.execute("""
        SELECT id, ear_tag, birth_date, weak FROM sheep
        WHERE status='在群'
    """).fetchall()
    meas_rows = conn.execute("SELECT sheep_id, stage FROM measurements").fetchall()
    conn.close()

    measured = {}
    for r in meas_rows:
        measured.setdefault(r['sheep_id'], set()).add(r['stage'])

    alerts = []           # 窗口内到期的未测阶段（主提醒）
    history_missing = []  # 历史漏测（早于窗口且仍未测）
    for s in sheep_rows:
        if s['weak']:
            continue
        birth = _parse_date(s['birth_date'])
        if not birth:
            continue
        done = measured.get(s['id'], set())
        for (stage, thr) in stages:
            if stage in done:
                continue
            dd = _add_months(birth, thr)
            if dd < prev_start:
                # 早于窗口：历史漏测待补测（如应早已测的初生/3月龄）
                history_missing.append({
                    'sheep_id': s['id'],
                    'ear_tag': s['ear_tag'],
                    'stage': stage,
                    'due_date': dd.isoformat(),
                    'overdue': (today - dd).days
                })
            elif prev_start <= dd <= cur_end:
                # 窗口内：主提醒；标注此前未测的早期阶段
                missing_early = [st for (st, _) in stages[:stage_idx[stage]] if st not in done]
                alerts.append({
                    'sheep_id': s['id'],
                    'ear_tag': s['ear_tag'],
                    'stage': stage,
                    'birth_date': s['birth_date'],
                    'due_date': dd.isoformat(),
                    'overdue': (today - dd).days,
                    'missing_early': missing_early
                })
            # dd > cur_end：未来到期，忽略

    # 按阶段汇总
    summary = {}
    for st, _ in stages:
        summary[st] = sum(1 for a in alerts if a['stage'] == st)
    # 按阶段分组列表
    grouped = {}
    for a in alerts:
        grouped.setdefault(a['stage'], []).append(a)

    return jsonify({
        'alerts': alerts,
        'grouped': grouped,
        'summary': summary,
        'total': len(alerts),
        'history_missing': history_missing,
        'history_count': len(history_missing)
    })

# ──────────────────────────────────────────────────────────
# 羊只 CRUD API
# ──────────────────────────────────────────────────────────
@app.route('/api/sheep')
@login_required
def api_sheep_list():
    conn = get_db()
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)
    search = request.args.get('search', '')
    sex = request.args.get('sex', '')
    status = request.args.get('status', '')
    start_date = request.args.get('start_date', '')
    end_date = request.args.get('end_date', '')

    where = ["1=1"]
    params = []
    if search:
        where.append("(ear_tag LIKE ? OR electronic_ear_tag LIKE ? OR father_tag LIKE ? OR mother_tag LIKE ?)")
        like = f'%{search}%'
        params.extend([like, like, like, like])
    if sex:
        where.append("sex = ?")
        params.append(sex)
    if status:
        where.append("status = ?")
        params.append(status)
    if start_date:
        where.append("birth_date >= ?")
        params.append(start_date)
    if end_date:
        where.append("birth_date <= ?")
        params.append(end_date)

    # 排序
    sort_by = request.args.get('sort_by', 'id')
    sort_order = request.args.get('sort_order', 'desc')
    allowed_sort = {'id','birth_date','ear_tag'}
    if sort_by not in allowed_sort: sort_by = 'id'
    if sort_order not in ('asc','desc'): sort_order = 'desc'
    order_clause = f'ORDER BY {sort_by} {sort_order}, id DESC'

    count = conn.execute(f"SELECT COUNT(*) FROM sheep WHERE {' AND '.join(where)}", params).fetchone()[0]
    offset = (page - 1) * per_page
    rows = conn.execute(
        f"SELECT * FROM sheep WHERE {' AND '.join(where)} {order_clause} LIMIT ? OFFSET ?",
        params + [per_page, offset]
    ).fetchall()
    conn.close()

    return jsonify({
        'total': count,
        'page': page,
        'per_page': per_page,
        'items': [dict(r) for r in rows]
    })

@app.route('/api/sheep/<int:sheep_id>')
@login_required
def api_sheep_get(sheep_id):
    conn = get_db()
    sheep = conn.execute("SELECT * FROM sheep WHERE id=?", (sheep_id,)).fetchone()
    if not sheep:
        conn.close()
        return jsonify({'error': '未找到'}), 404

    # 获取测定记录
    measurements = conn.execute(
        "SELECT * FROM measurements WHERE sheep_id=? ORDER BY CASE stage WHEN '初生' THEN 1 WHEN '3月龄' THEN 2 WHEN '6月龄' THEN 3 WHEN '12月龄' THEN 4 END",
        (sheep_id,)
    ).fetchall()

    # 获取该羊的后代
    offspring = conn.execute(
        "SELECT * FROM sheep WHERE (father_tag=? OR mother_tag=?) AND ear_tag != ?",
        (sheep['ear_tag'], sheep['ear_tag'], sheep['ear_tag'])
    ).fetchall()

    # 如果是母羊，获取繁殖记录
    breeding = []
    if sheep['sex'] == '母':
        breeding = conn.execute(
            "SELECT * FROM breeding_records WHERE ewe_tag=? ORDER BY lambing_date DESC",
            (sheep['ear_tag'],)
        ).fetchall()
        # 每个繁殖记录的羔羊
        for i, br in enumerate(breeding):
            lambs = conn.execute("SELECT * FROM lambs WHERE breeding_record_id=?", (br['id'],)).fetchall()
            breeding[i] = dict(br)
            breeding[i]['lambs'] = [dict(l) for l in lambs]

    conn.close()

    return jsonify({
        'sheep': dict(sheep),
        'measurements': [dict(m) for m in measurements],
        'offspring': [dict(o) for o in offspring],
        'breeding': [dict(b) for b in breeding] if sheep['sex'] == '母' else []
    })

@app.route('/api/sheep', methods=['POST'])
@admin_required
def api_sheep_create():
    data = request.get_json()
    conn = get_db()
    try:
        conn.execute("""
            INSERT INTO sheep (ear_tag, electronic_ear_tag, breed, birth_date, sex, source, status, status_reason, father_tag, mother_tag)
            VALUES (?,?,?,?,?,?,?,?,?,?)
        """, (
            data['ear_tag'], data.get('electronic_ear_tag', ''),
            data.get('breed', '努比亚山羊'), data.get('birth_date', ''),
            data['sex'], data.get('source', ''),
            data.get('status', '在群'), data.get('status_reason', ''),
            data.get('father_tag', ''),
            data.get('mother_tag', '')
        ))
        sheep_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        conn.commit()
        return jsonify({'success': True, 'id': sheep_id})
    except sqlite3.IntegrityError:
        return jsonify({'success': False, 'message': '耳号已存在'})
    finally:
        conn.close()

@app.route('/api/sheep/<int:sheep_id>', methods=['PUT'])
@admin_required
def api_sheep_update(sheep_id):
    data = request.get_json()
    conn = get_db()
    try:
        new_tag = data['ear_tag']
        # 取旧耳号，用于判断是否需要级联
        old = conn.execute("SELECT ear_tag FROM sheep WHERE id=?", (sheep_id,)).fetchone()
        if not old:
            return jsonify({'success': False, 'message': '羊只不存在'}), 404
        old_tag = old['ear_tag']

        # 耳号变更时，先校验新耳号不与其他羊重复
        if new_tag != old_tag:
            dup = conn.execute(
                "SELECT id FROM sheep WHERE ear_tag=? AND id!=?", (new_tag, sheep_id)
            ).fetchone()
            if dup:
                return jsonify({'success': False, 'message': f'耳号 {new_tag} 已存在，不能重复'}), 400

        # 1) 更新本行
        conn.execute("""
            UPDATE sheep SET ear_tag=?, electronic_ear_tag=?, breed=?, birth_date=?, sex=?,
            source=?, status=?, status_reason=?, father_tag=?, mother_tag=?, updated_at=CURRENT_TIMESTAMP
            WHERE id=?
        """, (
            new_tag, data.get('electronic_ear_tag', ''),
            data.get('breed', '努比亚山羊'), data.get('birth_date', ''),
            data['sex'], data.get('source', ''),
            data.get('status', '在群'), data.get('status_reason', ''),
            data.get('father_tag', ''),
            data.get('mother_tag', ''), sheep_id
        ))

        # 2) 耳号级联：同步所有以旧耳号为关联键的记录（系谱全局联动）
        if new_tag != old_tag:
            conn.execute("UPDATE sheep SET father_tag=? WHERE father_tag=?", (new_tag, old_tag))
            conn.execute("UPDATE sheep SET mother_tag=? WHERE mother_tag=?", (new_tag, old_tag))
            conn.execute("UPDATE breeding_records SET ram_tag=? WHERE ram_tag=?", (new_tag, old_tag))
            conn.execute("UPDATE breeding_records SET ewe_tag=? WHERE ewe_tag=?", (new_tag, old_tag))
            conn.execute("UPDATE lambs SET lamb_ear_tag=? WHERE lamb_ear_tag=?", (new_tag, old_tag))

        conn.commit()
        return jsonify({'success': True})
    except Exception as e:
        conn.rollback()
        return jsonify({'success': False, 'message': str(e)}), 500
    finally:
        conn.close()

@app.route('/api/sheep/<int:sheep_id>', methods=['DELETE'])
@admin_required
def api_sheep_delete(sheep_id):
    conn = get_db()
    try:
        conn.execute("DELETE FROM measurements WHERE sheep_id=?", (sheep_id,))
        conn.execute("DELETE FROM sheep WHERE id=?", (sheep_id,))
        conn.commit()
        return jsonify({'success': True})
    except Exception as e:
        conn.rollback()
        return jsonify({'success': False, 'message': str(e)}), 500
    finally:
        conn.close()

# ──────────────────────────────────────────────────────────
# 配种繁殖 CRUD API
# ──────────────────────────────────────────────────────────
@app.route('/api/breeding')
@login_required
def api_breeding_list():
    conn = get_db()
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)
    search = request.args.get('search', '')

    where = ["1=1"]
    params = []
    if search:
        where.append("(ewe_tag LIKE ? OR ram_tag LIKE ?)")
        like = f'%{search}%'
        params.extend([like, like])

    count = conn.execute(f"SELECT COUNT(*) FROM breeding_records br WHERE {' AND '.join(where)}", params).fetchone()[0]
    offset = (page - 1) * per_page

    # 排序
    sort_by = request.args.get('sort_by', 'lambing_date')
    sort_order = request.args.get('sort_order', 'desc')
    allowed_sort = {'lambing_date','mating_date','id','ewe_tag'}
    if sort_by not in allowed_sort: sort_by = 'lambing_date'
    if sort_order not in ('asc','desc'): sort_order = 'desc'

    rows = conn.execute(
        f"""SELECT br.*,
            (SELECT COUNT(*) FROM lambs WHERE breeding_record_id=br.id) as lamb_count
            FROM breeding_records br
            WHERE {' AND '.join(where)}
            ORDER BY br.{sort_by} {sort_order} LIMIT ? OFFSET ?""",
        params + [per_page, offset]
    ).fetchall()
    conn.close()

    return jsonify({
        'total': count, 'page': page, 'per_page': per_page,
        'items': [dict(r) for r in rows]
    })

@app.route('/api/breeding/<int:rec_id>')
@login_required
def api_breeding_get(rec_id):
    conn = get_db()
    rec = conn.execute("SELECT * FROM breeding_records WHERE id=?", (rec_id,)).fetchone()
    if not rec:
        conn.close()
        return jsonify({'error': '未找到'}), 404
    lambs = conn.execute("SELECT * FROM lambs WHERE breeding_record_id=?", (rec_id,)).fetchall()
    conn.close()

    result = dict(rec)
    result['lambs'] = [dict(l) for l in lambs]
    return jsonify(result)

@app.route('/api/breeding', methods=['POST'])
@admin_required
def api_breeding_create():
    data = request.get_json()
    conn = get_db()
    try:
        conn.execute("""
            INSERT INTO breeding_records (ewe_tag, ram_tag, mating_date, lambing_date, total_born, live_born, remark)
            VALUES (?,?,?,?,?,?,?)
        """, (
            data['ewe_tag'], data.get('ram_tag', ''),
            data.get('mating_date', ''), data['lambing_date'],
            data.get('total_born', 0), data.get('live_born', 0),
            data.get('remark', '')
        ))
        rec_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]

        # 羔羊明细
        lambs_data = data.get('lambs', [])
        for l in lambs_data:
            conn.execute("""
                INSERT INTO lambs (breeding_record_id, lamb_ear_tag, sex, birth_weight)
                VALUES (?,?,?,?)
            """, (rec_id, l['lamb_ear_tag'], l.get('sex', ''), l.get('birth_weight', 0)))
            # 同步录入羊只表（如果耳号不重复）
            tag = (l.get('lamb_ear_tag') or '').strip()
            if tag:
                exist = conn.execute("SELECT id FROM sheep WHERE ear_tag=?", (tag,)).fetchone()
                if not exist:
                    conn.execute("""
                        INSERT INTO sheep (ear_tag, sex, birth_date, father_tag, mother_tag, status)
                        VALUES (?,?,?,?,?,'在群')
                    """, (tag, l.get('sex', ''),
                          data.get('lambing_date', ''),
                          data.get('ram_tag', ''),
                          data.get('ewe_tag', '')))

        # 确保父母羊只存在
        for ptag, psex in [(data.get('ram_tag',''), '公'), (data.get('ewe_tag',''), '母')]:
            ptag = (ptag or '').strip()
            if ptag:
                pex = conn.execute("SELECT id FROM sheep WHERE ear_tag=?", (ptag,)).fetchone()
                if not pex:
                    conn.execute("INSERT INTO sheep (ear_tag, sex, status) VALUES (?,?, '在群')", (ptag, psex))

        conn.commit()
        return jsonify({'success': True, 'id': rec_id})
    finally:
        conn.close()

@app.route('/api/breeding/<int:rec_id>', methods=['PUT'])
@admin_required
def api_breeding_update(rec_id):
    data = request.get_json()
    conn = get_db()
    conn.execute("""
        UPDATE breeding_records SET ewe_tag=?, ram_tag=?, mating_date=?,
        lambing_date=?, total_born=?, live_born=?, remark=?
        WHERE id=?
    """, (
        data['ewe_tag'], data.get('ram_tag', ''),
        data.get('mating_date', ''), data['lambing_date'],
        data.get('total_born', 0), data.get('live_born', 0),
        data.get('remark', ''), rec_id
    ))

    # 更新羔羊
    conn.execute("DELETE FROM lambs WHERE breeding_record_id=?", (rec_id,))
    lambs_data = data.get('lambs', [])
    for l in lambs_data:
        conn.execute("""
            INSERT INTO lambs (breeding_record_id, lamb_ear_tag, sex, birth_weight)
            VALUES (?,?,?,?)
        """, (rec_id, l['lamb_ear_tag'], l.get('sex', ''), l.get('birth_weight', 0)))
        # 同步录入羊只表
        tag = (l.get('lamb_ear_tag') or '').strip()
        if tag:
            exist = conn.execute("SELECT id FROM sheep WHERE ear_tag=?", (tag,)).fetchone()
            if not exist:
                conn.execute("INSERT INTO sheep (ear_tag, sex, birth_date, father_tag, mother_tag, status) VALUES (?,?,?,?,?,'在群')",
                    (tag, l.get('sex', ''), data.get('lambing_date', ''), data.get('ram_tag', ''), data.get('ewe_tag', '')))

    # 确保父母羊只存在
    for ptag, psex in [(data.get('ram_tag',''), '公'), (data.get('ewe_tag',''), '母')]:
        ptag = (ptag or '').strip()
        if ptag:
            pex = conn.execute("SELECT id FROM sheep WHERE ear_tag=?", (ptag,)).fetchone()
            if not pex:
                conn.execute("INSERT INTO sheep (ear_tag, sex, status) VALUES (?,?, '在群')", (ptag, psex))

    conn.commit()
    conn.close()
    return jsonify({'success': True})

@app.route('/api/breeding/<int:rec_id>', methods=['DELETE'])
@admin_required
def api_breeding_delete(rec_id):
    conn = get_db()
    conn.execute("DELETE FROM lambs WHERE breeding_record_id=?", (rec_id,))
    conn.execute("DELETE FROM breeding_records WHERE id=?", (rec_id,))
    conn.commit()
    conn.close()
    return jsonify({'success': True})

# ──────────────────────────────────────────────────────────
# 测定 CRUD API
# ──────────────────────────────────────────────────────────
@app.route('/api/measurement')
@login_required
def api_measurement_list():
    conn = get_db()
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)
    stage = request.args.get('stage', '')

    where = ["1=1"]
    params = []
    if stage:
        where.append("m.stage = ?")
        params.append(stage)

    count = conn.execute(
        f"SELECT COUNT(*) FROM measurements m JOIN sheep s ON m.sheep_id = s.id WHERE {' AND '.join(where)}",
        params
    ).fetchone()[0]

    offset = (page - 1) * per_page
    rows = conn.execute(
        f"""SELECT m.*, s.ear_tag, s.sex, s.birth_date
            FROM measurements m JOIN sheep s ON m.sheep_id = s.id
            WHERE {' AND '.join(where)}
            ORDER BY m.id DESC LIMIT ? OFFSET ?""",
        params + [per_page, offset]
    ).fetchall()
    conn.close()

    return jsonify({
        'total': count, 'page': page, 'per_page': per_page,
        'items': [dict(r) for r in rows]
    })


@app.route('/api/measurement/by_sheep')
@login_required
def api_measurement_by_sheep():
    """按羊只聚合测定记录，阶段列为列（初生→3月龄→6月龄→12月龄），适合单行展示"""
    conn = get_db()
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)
    search = request.args.get('search', '')

    where_sheep = ["1=1"]
    params = []
    if search:
        where_sheep.append("s.ear_tag LIKE ?")
        params.append(f"%{search}%")

    count = conn.execute(
        f"""SELECT COUNT(DISTINCT s.id) FROM sheep s
            JOIN measurements m ON m.sheep_id = s.id
            WHERE {' AND '.join(where_sheep)}""",
        params
    ).fetchone()[0]

    # 排序
    sort_by = request.args.get('sort_by', 'ear_tag')
    sort_order = request.args.get('sort_order', 'asc')
    allowed_sort = {'ear_tag','birth_date','sex'}
    if sort_by not in allowed_sort: sort_by = 'ear_tag'
    if sort_order not in ('asc','desc'): sort_order = 'asc'

    offset = (page - 1) * per_page
    sheep_ids = conn.execute(
        f"""SELECT DISTINCT s.id, s.ear_tag, s.sex, s.birth_date
            FROM sheep s JOIN measurements m ON m.sheep_id = s.id
            WHERE {' AND '.join(where_sheep)}
            ORDER BY s.{sort_by} {sort_order} LIMIT ? OFFSET ?""",
        params + [per_page, offset]
    ).fetchall()

    sid_list = [s['id'] for s in sheep_ids]
    if not sid_list:
        conn.close()
        return jsonify({'total': 0, 'page': page, 'per_page': per_page, 'items': []})

    placeholders = ','.join('?' * len(sid_list))
    meas_rows = conn.execute(
        f"SELECT * FROM measurements WHERE sheep_id IN ({placeholders}) ORDER BY sheep_id, stage",
        sid_list
    ).fetchall()
    conn.close()

    stage_order = ['初生', '3月龄', '6月龄', '12月龄']
    from collections import defaultdict
    grouped = defaultdict(dict)
    for m in meas_rows:
        grouped[m['sheep_id']][m['stage']] = dict(m)

    items = []
    for s in sheep_ids:
        stages = {}
        next_stage = None
        for st in stage_order:
            if st in grouped[s['id']]:
                stages[st] = grouped[s['id']][st]
            else:
                stages[st] = None
                if next_stage is None:
                    next_stage = st
        items.append({
            'sheep_id': s['id'],
            'ear_tag': s['ear_tag'],
            'sex': s['sex'],
            'birth_date': s['birth_date'],
            'stages': stages,
            'next_stage': next_stage
        })

    return jsonify({
        'total': count, 'page': page, 'per_page': per_page,
        'items': items
    })

@app.route('/api/measurement/<int:meas_id>')
@login_required
def api_measurement_get(meas_id):
    conn = get_db()
    m = conn.execute("SELECT * FROM measurements WHERE id=?", (meas_id,)).fetchone()
    conn.close()
    return jsonify({'error': '未找到'}) if not m else jsonify(dict(m))

@app.route('/api/measurement', methods=['POST'])
@admin_required
def api_measurement_create():
    data = request.get_json()
    if not data or 'stage' not in data:
        return jsonify({'success': False, 'message': '缺少必要字段 stage'}), 400

    conn = get_db()
    try:
        sheep_id = data.get('sheep_id')
        ear_tag = data.get('ear_tag', '').strip()
        imported_breeding = 0

        # 新羊创建
        if not sheep_id and ear_tag:
            # 检查是否已存在该耳号（仅作二次保护，前端已校验）
            exist = conn.execute("SELECT id FROM sheep WHERE ear_tag=?", (ear_tag,)).fetchone()
            if exist:
                sheep_id = exist[0]
            else:
                new_birth = data.get('new_birth_date', '')
                new_breed = data.get('new_breed', '') or ''
                ensure_breed(conn, new_breed)
                new_sex = data.get('new_sex', '母')
                new_father = (data.get('new_father') or '').strip()
                new_mother = (data.get('new_mother') or '').strip()

                if not new_birth:
                    return jsonify({'success': False, 'message': '新羊需填写出生日期'}), 400

                conn.execute("""
                    INSERT INTO sheep (ear_tag, breed, birth_date, sex, father_tag, mother_tag, status)
                    VALUES (?, ?, ?, ?, ?, ?, '在群')
                """, (ear_tag, new_breed, new_birth, new_sex, new_father, new_mother))
                sheep_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]

                # 确保父母羊只存在
                for ptag, psex in [(new_father, '公'), (new_mother, '母')]:
                    if ptag:
                        pex = conn.execute("SELECT id FROM sheep WHERE ear_tag=?", (ptag,)).fetchone()
                        if not pex:
                            conn.execute("INSERT INTO sheep (ear_tag, sex, status, breed) VALUES (?,?, '在群', ?)", (ptag, psex, new_breed))

                # 自动创建配种记录
                if new_father and new_mother and new_birth:
                    try:
                        bd = datetime.strptime(new_birth, '%Y-%m-%d')
                        mating_date = (bd - timedelta(days=150)).strftime('%Y-%m-%d')
                    except:
                        mating_date = None
                    if mating_date:
                        dup = conn.execute("""
                            SELECT id FROM breeding_records
                            WHERE ewe_tag=? AND ram_tag=? AND lambing_date=?
                        """, (new_mother, new_father, new_birth)).fetchone()
                        if not dup:
                            conn.execute("""
                                INSERT INTO breeding_records (ewe_tag, ram_tag, mating_date, lambing_date, total_born, live_born, remark)
                                VALUES (?,?,?,?,1,1,'手动新增测定自动生成')
                            """, (new_mother, new_father, mating_date, new_birth))
                            new_br_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
                            # 同步创建羔羊记录（这只羔羊就是当前正在录入的羊）
                            conn.execute("""
                                INSERT INTO lambs (breeding_record_id, lamb_ear_tag, sex, birth_weight)
                                VALUES (?, ?, ?, ?)
                            """, (new_br_id, ear_tag, new_sex, data.get('weight', 0)))
                            imported_breeding = 1

        if not sheep_id:
            return jsonify({'success': False, 'message': '请选择或输入羊只耳号'}), 400

        conn.execute("""
            INSERT INTO measurements (sheep_id, stage, weight, body_height, body_length,
            chest_girth, cannon_circumference, hip_width, testis_circumference, testis_diameter, measure_date, judgment)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            sheep_id, data['stage'], data.get('weight', 0),
            data.get('body_height', 0), data.get('body_length', 0),
            data.get('chest_girth', 0), data.get('cannon_circumference', 0),
            data.get('hip_width', 0), data.get('testis_circumference', 0),
            data.get('testis_diameter', 0), data.get('measure_date', ''),
            data.get('judgment', '正常')
        ))
        m_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        _sync_weak(conn, sheep_id)
        conn.commit()
        return jsonify({'success': True, 'id': m_id, 'new_sheep': imported_breeding > 0, 'breeding_created': imported_breeding})
    except Exception as e:
        conn.rollback()
        return jsonify({'success': False, 'message': str(e)}), 500
    finally:
        conn.close()

def _sync_weak(conn, sheep_id):
    """根据测定记录重算该羊 weak 标签：任一阶段为弱羊则打标"""
    cnt = conn.execute("SELECT COUNT(*) FROM measurements WHERE sheep_id=? AND judgment='弱羊'",
                       (sheep_id,)).fetchone()[0]
    conn.execute("UPDATE sheep SET weak=? WHERE id=?", (1 if cnt > 0 else 0, sheep_id))

@app.route('/api/measurement/<int:meas_id>', methods=['PUT'])
@admin_required
def api_measurement_update(meas_id):
    data = request.get_json()
    conn = get_db()
    conn.execute("""
        UPDATE measurements SET sheep_id=?, stage=?, weight=?, body_height=?,
        body_length=?, chest_girth=?, cannon_circumference=?, hip_width=?,
        testis_circumference=?, testis_diameter=?, measure_date=?, judgment=?
        WHERE id=?
    """, (
        data['sheep_id'], data['stage'], data.get('weight', 0),
        data.get('body_height', 0), data.get('body_length', 0),
        data.get('chest_girth', 0), data.get('cannon_circumference', 0),
        data.get('hip_width', 0), data.get('testis_circumference', 0),
        data.get('testis_diameter', 0), data.get('measure_date', ''),
        data.get('judgment', '正常'),
        meas_id
    ))
    _sync_weak(conn, data['sheep_id'])
    conn.commit()
    conn.close()
    return jsonify({'success': True})

@app.route('/api/measurement/<int:meas_id>', methods=['DELETE'])
@admin_required
def api_measurement_delete(meas_id):
    conn = get_db()
    conn.execute("DELETE FROM measurements WHERE id=?", (meas_id,))
    conn.commit()
    conn.close()
    return jsonify({'success': True})

# ──────────────────────────────────────────────────────────
# 导入 API
# ──────────────────────────────────────────────────────────
def parse_number(val):
    """解析数值，处理中文逗号等"""
    if val is None:
        return 0
    s = str(val).strip()
    s = s.replace('，', ',').replace(' ', '').replace('\u3000', '')
    try:
        return float(s) if '.' in s else int(s)
    except:
        return 0

def parse_date(val):
    """日期解析"""
    if val is None:
        return ''
    if isinstance(val, datetime):
        return val.strftime('%Y-%m-%d')
    s = str(val).strip()
    # 尝试多种格式
    patterns = [
        (r'(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})', '%Y-%m-%d'),
        (r'(\d{4})年(\d{1,2})月(\d{1,2})日?', '%Y-%m-%d'),
    ]
    for pattern, fmt in patterns:
        m = re.match(pattern, s)
        if m:
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            return f'{y:04d}-{mo:02d}-{d:02d}'
    return s

def split_values(val):
    """分割多值字段，支持 顿号、逗号(中英文)、分号、空格、斜杠 等分隔符"""
    if val is None:
        return []
    s = str(val).strip()
    if not s:
        return []
    parts = re.split(r"[、,，;；\s/\\]+", s)
    return [p.strip() for p in parts if p and p.strip()]

def normalize_sex(val):
    """标准化性别值"""
    if val is None:
        return '公'
    s = str(val).strip()
    if s in ('公', '雄', 'male', 'Male', 'M'):
        return '公'
    if s in ('母', '雌', 'female', 'Female', 'F'):
        return '母'
    return '公'  # 默认公

def safe_get(row, idx, default=''):
    """安全获取单元格值"""
    if idx >= len(row) or row[idx] is None:
        return default
    return str(row[idx]).strip()

@app.route('/api/admin/clear-data', methods=['POST'])
@super_admin_required
def api_clear_data():
    """清空业务数据（羊只/配种/羔羊/测定），保留用户与品种。仅超级管理员。"""
    conn = get_db()
    tables = ['measurements', 'lambs', 'breeding_records', 'sheep']
    counts = {}
    for t in tables:
        try:
            counts[t] = conn.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
        except Exception:
            counts[t] = 0
    for t in tables:
        conn.execute(f"DELETE FROM {t}")
    try:
        conn.execute("DELETE FROM sqlite_sequence WHERE name IN ('measurements','lambs','breeding_records','sheep')")
    except Exception:
        pass
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '已清空业务数据（保留用户与品种）', 'deleted': counts})


def _resolve_workbook():
    """从上传文件或桌面端传入的本地路径载入 Excel 工作簿。

    桌面端(pywebview)通过原生文件对话框选择后，以 filepath 表单字段提交本地绝对路径；
    浏览器端仍以 multipart 上传文件。
    返回 (workbook, None) 或 (None, 错误信息)。
    """
    # 1. 浏览器/网页上传
    f = request.files.get('file')
    if f is not None and f.filename:
        return openpyxl.load_workbook(f), None
    # 2. 桌面端：表单或 JSON 传入的本地绝对路径
    filepath = request.form.get('filepath')
    if not filepath:
        try:
            filepath = (request.get_json(silent=True) or {}).get('filepath')
        except Exception:
            filepath = None
    if not filepath:
        return None, '请选择文件'
    if not os.path.exists(filepath):
        return None, f'文件不存在: {filepath}'
    if not str(filepath).lower().endswith(('.xlsx', '.xlsm')):
        return None, '仅支持 .xlsx 格式'
    try:
        return openpyxl.load_workbook(filepath), None
    except Exception as e:
        return None, f'读取 Excel 失败: {e}'


@app.route('/api/import/production', methods=['POST'])
@admin_required
def api_import_production():
    """导入产仔记录 Excel"""
    wb, err = _resolve_workbook()
    if err:
        return jsonify({'success': False, 'message': err})
    ws = wb.active

    conn = get_db()
    imported_breeding = 0
    imported_lambs = 0
    imported_sheep = 0
    skipped_lambs = 0
    errors = []

    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        try:
            ewe_tag = safe_get(row, 0)
            ram_tag = safe_get(row, 1)
            breed_from_excel = safe_get(row, 2)   # 品种字段（新增）
            ensure_breed(conn, breed_from_excel)
            lambing_date = parse_date(row[3])
            total_born = parse_number(row[4]) if len(row) > 4 and row[4] is not None else 0
            live_born = parse_number(row[5]) if len(row) > 5 and row[5] is not None else 0
            lamb_sexes = split_values(row[6]) if len(row) > 6 else []
            # 初生重：跳过非数字部分（如录入笔误 '3.6.3.4'），避免整行失败
            if len(row) > 7 and row[7] is not None:
                lamb_weights = []
                for x in split_values(row[7]):
                    try:
                        lamb_weights.append(float(x))
                    except ValueError:
                        pass
            else:
                lamb_weights = []
            lamb_tags = split_values(row[8]) if len(row) > 8 else []
            remark = safe_get(row, 9)

            if not lambing_date:
                continue

            # 确保母羊存在，品种来自Excel或继承
            ewe_breed = breed_from_excel
            if ewe_tag:
                existing_ewe = conn.execute("SELECT id, breed FROM sheep WHERE ear_tag=?", (ewe_tag,)).fetchone()
                if existing_ewe:
                    ewe_breed = existing_ewe['breed'] or breed_from_excel
                    # 如果Excel有品种且与现有不同，更新
                    if breed_from_excel and existing_ewe['breed'] != breed_from_excel:
                        conn.execute("UPDATE sheep SET breed=? WHERE ear_tag=?", (breed_from_excel, ewe_tag))
                else:
                    conn.execute("""
                        INSERT INTO sheep (ear_tag, sex, status, breed)
                        VALUES (?, '母', '在群', ?)
                    """, (ewe_tag, breed_from_excel))
                    imported_sheep += 1

            # 确保公羊存在
            ram_breed = breed_from_excel
            if ram_tag:
                existing_ram = conn.execute("SELECT id, breed FROM sheep WHERE ear_tag=?", (ram_tag,)).fetchone()
                if existing_ram:
                    ram_breed = existing_ram['breed'] or breed_from_excel
                    if breed_from_excel and existing_ram['breed'] != breed_from_excel:
                        conn.execute("UPDATE sheep SET breed=? WHERE ear_tag=?", (breed_from_excel, ram_tag))
                else:
                    conn.execute("""
                        INSERT INTO sheep (ear_tag, sex, status, breed)
                        VALUES (?, '公', '在群', ?)
                    """, (ram_tag, breed_from_excel))
                    imported_sheep += 1

            # 品种继承逻辑：父母品种一致则自动继承，否则留空
            inherited_breed = ''
            if ewe_breed and ram_breed and ewe_breed == ram_breed:
                inherited_breed = ewe_breed

            # 估算配种日期 (产羔日期 - 150天)
            try:
                ld = datetime.strptime(lambing_date, '%Y-%m-%d')
                md = (ld - timedelta(days=150)).strftime('%Y-%m-%d')
            except:
                md = ''

            # 创建繁殖记录
            conn.execute("""
                INSERT INTO breeding_records (ewe_tag, ram_tag, mating_date, lambing_date, total_born, live_born, remark)
                VALUES (?,?,?,?,?,?,?)
            """, (ewe_tag, ram_tag, md, lambing_date, total_born, live_born, remark))
            rec_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            imported_breeding += 1

            # 创建羔羊记录
            lamb_count = len(lamb_tags)
            for i in range(lamb_count):
                sex = normalize_sex(lamb_sexes[i]) if i < len(lamb_sexes) else '公'
                weight = lamb_weights[i] if i < len(lamb_weights) else 0
                tag = lamb_tags[i]

                # 耳号唯一：同一耳号已存在羔羊明细则视为源数据重复，跳过以免虚增
                if conn.execute("SELECT 1 FROM lambs WHERE lamb_ear_tag=?", (tag,)).fetchone():
                    skipped_lambs += 1
                    # 仍确保该羊在羊只表中（更新父母/出生）
                    conn.execute("""
                        UPDATE sheep SET father_tag=?, mother_tag=?, birth_date=? WHERE ear_tag=?
                    """, (ram_tag, ewe_tag, lambing_date, tag))
                    continue

                conn.execute("""
                    INSERT INTO lambs (breeding_record_id, lamb_ear_tag, sex, birth_weight)
                    VALUES (?,?,?,?)
                """, (rec_id, tag, sex, weight))
                imported_lambs += 1

                # 确保羔羊在羊只表中
                existing_lamb = conn.execute("SELECT id FROM sheep WHERE ear_tag=?", (tag,)).fetchone()
                if not existing_lamb:
                    conn.execute("""
                        INSERT INTO sheep (ear_tag, sex, birth_date, father_tag, mother_tag, status, breed)
                        VALUES (?,?,?,?,?,'在群',?)
                    """, (tag, sex, lambing_date, ram_tag, ewe_tag, inherited_breed))
                    imported_sheep += 1
                    new_sheep_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
                    # 创建初生测定记录
                    conn.execute("""
                        INSERT INTO measurements (sheep_id, stage, weight, measure_date)
                        VALUES (?, '初生', ?, ?)
                    """, (new_sheep_id, weight, lambing_date))
                else:
                    # 更新父号母号、出生日期
                    conn.execute("""
                        UPDATE sheep SET father_tag=?, mother_tag=?, birth_date=? WHERE ear_tag=?
                    """, (ram_tag, ewe_tag, lambing_date, tag))

        except Exception as e:
            errors.append(f'第{row_idx}行: {str(e)}')

    conn.commit()
    conn.close()

    return jsonify({
        'success': True,
        'imported_breeding': imported_breeding,
        'imported_lambs': imported_lambs,
        'imported_sheep': imported_sheep,
        'skipped_lambs': skipped_lambs,
        'errors': errors[:10]
    })

@app.route('/api/import/growth', methods=['POST'])
@admin_required
def api_import_growth():
    """导入生长测定记录 Excel"""
    wb, err = _resolve_workbook()
    if err:
        return jsonify({'success': False, 'message': err})
    ws = wb.active

    conn = get_db()
    imported_sheep = 0
    imported_measurements = 0
    imported_breeding = 0
    skipped_lambs = 0
    errors = []

    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        try:
            ear_tag = safe_get(row, 0)
            electronic_tag = safe_get(row, 1)
            breed_from_excel = safe_get(row, 2)   # 品种字段（新增）
            ensure_breed(conn, breed_from_excel)
            birth_date = parse_date(row[3])
            sex = safe_get(row, 4)
            father_tag = safe_get(row, 5)
            mother_tag = safe_get(row, 6)

            if not ear_tag:
                continue
            if not birth_date:
                continue

            # 羊只信息（含品种）
            existing = conn.execute("SELECT id, breed FROM sheep WHERE ear_tag=?", (ear_tag,)).fetchone()
            if existing:
                sheep_id = existing[0]
                # 更新品种（如果Excel有新值且与现有不同）
                if breed_from_excel and existing['breed'] != breed_from_excel:
                    conn.execute("""
                        UPDATE sheep SET electronic_ear_tag=?, breed=?, birth_date=?, sex=?, father_tag=?, mother_tag=?, updated_at=CURRENT_TIMESTAMP
                        WHERE id=?
                    """, (electronic_tag, breed_from_excel, birth_date, sex, father_tag, mother_tag, sheep_id))
                else:
                    conn.execute("""
                        UPDATE sheep SET electronic_ear_tag=?, birth_date=?, sex=?, father_tag=?, mother_tag=?, updated_at=CURRENT_TIMESTAMP
                        WHERE id=?
                    """, (electronic_tag, birth_date, sex, father_tag, mother_tag, sheep_id))
            else:
                conn.execute("""
                    INSERT INTO sheep (ear_tag, electronic_ear_tag, breed, birth_date, sex, father_tag, mother_tag, status)
                    VALUES (?,?,?,?,?,?,?,'在群')
                """, (ear_tag, electronic_tag, breed_from_excel, birth_date, sex, father_tag, mother_tag))
                sheep_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
                imported_sheep += 1

            # 确保父母在羊只表中（品种为空）
            for ptag, psex in [(father_tag, '公'), (mother_tag, '母')]:
                if ptag and ptag.strip():
                    pex = conn.execute("SELECT id FROM sheep WHERE ear_tag=?", (ptag,)).fetchone()
                    if not pex:
                        conn.execute("""
                            INSERT INTO sheep (ear_tag, sex, status, breed) VALUES (?,?, '在群', ?)
                        """, (ptag, psex, breed_from_excel))
                        imported_sheep += 1

            # --- 自动创建配种记录（由生长测定模板反推）---
            if father_tag and father_tag.strip() and mother_tag and mother_tag.strip() and birth_date and sheep_id:
                # 山羊妊娠期约150天 → 配种日期 = 出生日期 - 150天
                try:
                    bd = datetime.strptime(birth_date, '%Y-%m-%d')
                    mating_date = (bd - timedelta(days=150)).strftime('%Y-%m-%d')
                except:
                    mating_date = None
                if mating_date:
                    dup = conn.execute("""
                        SELECT id FROM breeding_records
                        WHERE ewe_tag=? AND ram_tag=? AND lambing_date=?
                    """, (mother_tag.strip(), father_tag.strip(), birth_date)).fetchone()
                    lamb_exist = conn.execute("SELECT 1 FROM lambs WHERE lamb_ear_tag=?", (ear_tag,)).fetchone()
                    if not dup and not lamb_exist:
                        # 该羔羊尚无配种/羔羊记录 → 由测定反推创建
                        conn.execute("""
                            INSERT INTO breeding_records (ewe_tag, ram_tag, mating_date, lambing_date, total_born, live_born, remark)
                            VALUES (?, ?, ?, ?, ?, ?, ?)
                        """, (mother_tag.strip(), father_tag.strip(), mating_date, birth_date, 1, 1, '由测定导入自动生成'))
                        new_br_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
                        conn.execute("""
                            INSERT INTO lambs (breeding_record_id, lamb_ear_tag, sex, birth_weight)
                            VALUES (?, ?, ?, ?)
                        """, (new_br_id, ear_tag, sex, parse_number(row[7]) if len(row) > 7 and row[7] is not None else 0))
                        imported_breeding += 1
                    elif lamb_exist:
                        # 羔羊已由生产记录建立，无需再自动建配种（避免孤儿配种记录）
                        skipped_lambs += 1
                # ---

            # 初生重 (col 7)
            birth_weight = parse_number(row[7]) if len(row) > 7 and row[7] is not None else 0

            # 检查并创建初生测定
            m_exists = conn.execute(
                "SELECT id FROM measurements WHERE sheep_id=? AND stage='初生'", (sheep_id,)
            ).fetchone()
            if not m_exists and birth_weight > 0:
                conn.execute("""
                    INSERT INTO measurements (sheep_id, stage, weight, measure_date)
                    VALUES (?, '初生', ?, ?)
                """, (sheep_id, birth_weight, birth_date))
                imported_measurements += 1
            elif m_exists and birth_weight > 0:
                conn.execute("UPDATE measurements SET weight=? WHERE id=?", (birth_weight, m_exists[0]))

            # 3月龄 (col 8, 9)
            w3 = parse_number(row[8]) if len(row) > 8 and row[8] is not None else 0
            d3_raw = row[9] if len(row) > 9 and row[9] is not None else ''
            if isinstance(d3_raw, str) and d3_raw.startswith('=EDATE'):
                # 公式，用出生日期+3月计算
                try:
                    bd = datetime.strptime(birth_date, '%Y-%m-%d')
                    # Python: add ~91 days (3 months)
                    d3 = (bd.replace(month=bd.month+3) if bd.month <= 9
                         else bd.replace(year=bd.year+1, month=bd.month-9)).strftime('%Y-%m-%d')
                except:
                    d3 = ''
            else:
                d3 = parse_date(d3_raw) if d3_raw else ''

            if w3 > 0:
                m3_exists = conn.execute(
                    "SELECT id FROM measurements WHERE sheep_id=? AND stage='3月龄'", (sheep_id,)
                ).fetchone()
                if not m3_exists:
                    conn.execute("""
                        INSERT INTO measurements (sheep_id, stage, weight, measure_date)
                        VALUES (?, '3月龄', ?, ?)
                    """, (sheep_id, w3, d3))
                    imported_measurements += 1
                else:
                    conn.execute("UPDATE measurements SET weight=? WHERE id=?", (w3, m3_exists[0]))

            # 6月龄 (col 10, 11)
            w6 = parse_number(row[10]) if len(row) > 10 and row[10] is not None else 0
            d6_raw = row[11] if len(row) > 11 and row[11] is not None else ''
            d6 = parse_date(d6_raw) if d6_raw and not (isinstance(d6_raw, str) and d6_raw.startswith('=')) else ''

            if w6 > 0:
                m6_exists = conn.execute(
                    "SELECT id FROM measurements WHERE sheep_id=? AND stage='6月龄'", (sheep_id,)
                ).fetchone()
                if not m6_exists:
                    conn.execute("""
                        INSERT INTO measurements (sheep_id, stage, weight, measure_date)
                        VALUES (?, '6月龄', ?, ?)
                    """, (sheep_id, w6, d6))
                    imported_measurements += 1
                else:
                    conn.execute("UPDATE measurements SET weight=? WHERE id=?", (w6, m6_exists[0]))

            # 12月龄（完整体尺数据）
            w12 = parse_number(row[12]) if len(row) > 12 and row[12] is not None else 0
            if w12 > 0:
                m12_exists = conn.execute(
                    "SELECT id FROM measurements WHERE sheep_id=? AND stage='12月龄'", (sheep_id,)
                ).fetchone()

                m12_data = {
                    'sheep_id': sheep_id, 'stage': '12月龄',
                    'weight': w12,
                    'body_height': parse_number(row[13]) if len(row) > 13 else 0,
                    'body_length': parse_number(row[14]) if len(row) > 14 else 0,
                    'chest_girth': parse_number(row[15]) if len(row) > 15 else 0,
                    'hip_width': parse_number(row[16]) if len(row) > 16 else 0,
                    'cannon_circumference': parse_number(row[17]) if len(row) > 17 else 0,
                    'testis_diameter': parse_number(row[18]) if len(row) > 18 else 0,
                    'testis_circumference': parse_number(row[19]) if len(row) > 19 else 0,
                    'measure_date': parse_date(row[20]) if len(row) > 20 and row[20] is not None else ''
                }

                if not m12_exists:
                    conn.execute("""
                        INSERT INTO measurements (sheep_id, stage, weight, body_height, body_length,
                        chest_girth, cannon_circumference, hip_width, testis_circumference,
                        testis_diameter, measure_date)
                        VALUES (:sheep_id, :stage, :weight, :body_height, :body_length,
                        :chest_girth, :cannon_circumference, :hip_width, :testis_circumference,
                        :testis_diameter, :measure_date)
                    """, m12_data)
                    imported_measurements += 1
                else:
                    conn.execute("""
                        UPDATE measurements SET weight=?, body_height=?, body_length=?,
                        chest_girth=?, cannon_circumference=?, hip_width=?,
                        testis_circumference=?, testis_diameter=?, measure_date=?
                        WHERE id=?
                    """, (
                        w12, m12_data['body_height'], m12_data['body_length'],
                        m12_data['chest_girth'], m12_data['cannon_circumference'],
                        m12_data['hip_width'], m12_data['testis_circumference'],
                        m12_data['testis_diameter'], m12_data['measure_date'],
                        m12_exists[0]
                    ))

        except Exception as e:
            errors.append(f'第{row_idx}行: {str(e)}')

    conn.commit()
    conn.close()

    return jsonify({
        'success': True,
        'imported_sheep': imported_sheep,
        'imported_measurements': imported_measurements,
        'imported_breeding': imported_breeding,
        'skipped_lambs': skipped_lambs,
        'errors': errors[:10]
    })

# ──────────────────────────────────────────────────────────
# 导出 Excel API
# ──────────────────────────────────────────────────────────
@app.route('/api/export/sheep')
@login_required
def export_sheep():
    conn = get_db()
    sheep_list = conn.execute("SELECT * FROM sheep ORDER BY ear_tag").fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '羊只信息'

    headers = ['耳号', '电子耳号', '品种', '出生日期', '性别', '来源', '状态', '父号', '母号']
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    for i, s in enumerate(sheep_list, 2):
        ws.cell(row=i, column=1, value=s['ear_tag'])
        ws.cell(row=i, column=2, value=s['electronic_ear_tag'])
        ws.cell(row=i, column=3, value=s['breed'])
        ws.cell(row=i, column=4, value=s['birth_date'])
        ws.cell(row=i, column=5, value=s['sex'])
        ws.cell(row=i, column=6, value=s['source'])
        ws.cell(row=i, column=7, value=s['status'])
        ws.cell(row=i, column=8, value=s['father_tag'])
        ws.cell(row=i, column=9, value=s['mother_tag'])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name='羊只信息.xlsx')

@app.route('/api/export/breeding')
@login_required
def export_breeding():
    conn = get_db()
    records = conn.execute("""
        SELECT br.*,
            GROUP_CONCAT(l.lamb_ear_tag || '(' || l.sex || ' ' || l.birth_weight || 'kg)') as lamb_details
        FROM breeding_records br
        LEFT JOIN lambs l ON l.breeding_record_id = br.id
        GROUP BY br.id
        ORDER BY br.lambing_date DESC
    """).fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '配种繁殖记录'

    headers = ['母羊耳号', '公羊耳号', '配种日期', '产羔日期', '产羔数', '活羔数', '羔羊明细', '备注']
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    for i, r in enumerate(records, 2):
        ws.cell(row=i, column=1, value=r['ewe_tag'])
        ws.cell(row=i, column=2, value=r['ram_tag'])
        ws.cell(row=i, column=3, value=r['mating_date'])
        ws.cell(row=i, column=4, value=r['lambing_date'])
        ws.cell(row=i, column=5, value=r['total_born'])
        ws.cell(row=i, column=6, value=r['live_born'])
        ws.cell(row=i, column=7, value=r['lamb_details'] or '')
        ws.cell(row=i, column=8, value=r['remark'])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name='配种繁殖记录.xlsx')

@app.route('/api/export/measurement')
@login_required
def export_measurement():
    conn = get_db()
    records = conn.execute("""
        SELECT m.*, s.ear_tag, s.sex, s.birth_date, s.father_tag, s.mother_tag
        FROM measurements m JOIN sheep s ON m.sheep_id = s.id
        ORDER BY s.ear_tag, CASE m.stage WHEN '初生' THEN 1 WHEN '3月龄' THEN 2 WHEN '6月龄' THEN 3 WHEN '12月龄' THEN 4 END
    """).fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '生长测定记录'

    headers = ['羊耳号', '性别', '出生日期', '父号', '母号', '阶段', '测定日期',
               '体重(kg)', '体高(cm)', '体长(cm)', '胸围(cm)', '管围(cm)', '臀宽(cm)',
               '睾丸围(cm)', '睾丸直径(cm)', '判定']
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    for i, r in enumerate(records, 2):
        ws.cell(row=i, column=1, value=r['ear_tag'])
        ws.cell(row=i, column=2, value=r['sex'])
        ws.cell(row=i, column=3, value=r['birth_date'])
        ws.cell(row=i, column=4, value=r['father_tag'])
        ws.cell(row=i, column=5, value=r['mother_tag'])
        ws.cell(row=i, column=6, value=r['stage'])
        ws.cell(row=i, column=7, value=r['measure_date'])
        ws.cell(row=i, column=8, value=r['weight'])
        ws.cell(row=i, column=9, value=r['body_height'])
        ws.cell(row=i, column=10, value=r['body_length'])
        ws.cell(row=i, column=11, value=r['chest_girth'])
        ws.cell(row=i, column=12, value=r['cannon_circumference'])
        ws.cell(row=i, column=13, value=r['hip_width'])
        ws.cell(row=i, column=14, value=r['testis_circumference'])
        ws.cell(row=i, column=15, value=r['testis_diameter'])
        ws.cell(row=i, column=16, value=r['judgment'] if r['judgment'] else '正常')

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name='生长测定记录.xlsx')

# ──────────────────────────────────────────────────────────
# 羊只档案卡 PDF
# ──────────────────────────────────────────────────────────
# 使用纯文本+HTML转PDF的方案，不依赖外部字体
@app.route('/api/sheep/<int:sheep_id>/card/pdf')
@login_required
def sheep_card_pdf(sheep_id):
    """生成羊只档案卡PDF"""
    conn = get_db()
    sheep = conn.execute("SELECT * FROM sheep WHERE id=?", (sheep_id,)).fetchone()
    if not sheep:
        conn.close()
        return jsonify({'error': '未找到'}), 404

    measurements = conn.execute(
        "SELECT * FROM measurements WHERE sheep_id=? ORDER BY CASE stage WHEN '初生' THEN 1 WHEN '3月龄' THEN 2 WHEN '6月龄' THEN 3 WHEN '12月龄' THEN 4 END",
        (sheep_id,)
    ).fetchall()

    breeding = []
    if sheep['sex'] == '母':
        breeding = conn.execute(
            "SELECT * FROM breeding_records WHERE ewe_tag=? ORDER BY lambing_date",
            (sheep['ear_tag'],)
        ).fetchall()
        for i, br in enumerate(breeding):
            lambs = conn.execute("SELECT * FROM lambs WHERE breeding_record_id=?", (br['id'],)).fetchall()
            breeding[i] = dict(br)
            breeding[i]['lambs'] = [dict(l) for l in lambs]

    meas_map = {m['stage']: dict(m) for m in measurements}

    # 使用 reportlab 生成 PDF
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import mm
        from reportlab.pdfgen import canvas
        from reportlab.platypus import Table, TableStyle, SimpleDocTemplate
        from reportlab.lib import colors
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont

        # 注册中文字体
        font_paths = [
            '/usr/share/fonts/cesi/CESI_HT_GB18030.TTF',
            '/usr/share/fonts/kyfonts/STFANGSO.TTF',
            '/usr/share/fonts/gb/国标楷体.ttf',
            '/usr/share/fonts/wps-office/FZFSK.TTF',
            'C:/Windows/Fonts/simhei.ttf',
            '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
        ]
        font_registered = False
        for fp in font_paths:
            if os.path.exists(fp):
                try:
                    pdfmetrics.registerFont(TTFont('Chinese', fp))
                    font_registered = True
                    break
                except:
                    continue

        if not font_registered:
            return jsonify({'error': '系统中未找到中文字体'}), 500

        # 修复 reportlab + OpenSSL md5 兼容问题
        try:
            from reportlab.pdfbase import pdfdoc
            import hashlib
            pdfdoc.md5 = hashlib.md5
        except Exception:
            pass

        buf = io.BytesIO()
        c = canvas.Canvas(buf, pagesize=A4)
        width, height = A4
        font_name = 'Chinese'

        c.setFont(font_name, 18)
        title = '母羊档案卡' if (sheep['sex'] == '母' and breeding) else '羊只档案卡'
        c.drawCentredString(width/2, height - 30*mm, title)

        y = height - 45*mm
        left_margin = 20*mm
        right_margin = width - 20*mm

        # 基本信息
        c.setFont(font_name, 14)
        c.drawString(left_margin, y, '一、基本信息')
        y -= 12*mm

        basic_info = [
            ('场耳号', sheep['ear_tag']),
            ('电子耳号', sheep['electronic_ear_tag'] or ''),
            ('品种', sheep['breed']),
            ('性别', sheep['sex']),
            ('出生日期', sheep['birth_date']),
            ('来源', sheep['source'] or ''),
            ('状态', sheep['status']),
        ]
        c.setFont(font_name, 10)
        for label, value in basic_info:
            c.drawString(left_margin, y, f'{label}：{value}')
            y -= 8*mm

        # 系谱 - 横向四分代树形
        y -= 4*mm
        c.setFont(font_name, 14)
        c.drawString(left_margin, y, '二、系谱（三代）')
        y -= 6*mm
        c.setFont(font_name, 6.5)
        c.setFillColor(colors.HexColor('#adb5bd'))
        pg_note_y = y
        c.drawString(left_margin, y, '绿色=公 | 红色=母')
        c.setFillColor(colors.black)
        y -= 4*mm

        male_clr = colors.HexColor('#2d6a4f')
        female_clr = colors.HexColor('#c0392b')
        line_clr = colors.HexColor('#adb5bd')

        margin_inner = 12*mm
        tree_left = left_margin + margin_inner
        tree_right = right_margin - margin_inner
        tree_width = tree_right - tree_left

        # 4代，等宽列
        col_w = tree_width / 4.0
        col_centers = [tree_left + col_w * (i + 0.5) for i in range(4)]
        col_labels = ['羊只', '亲代', '祖代', '曾祖代']
        node_h = 7*mm
        node_w = col_w * 0.55
        gen_gap_y = 10.5*mm   # 代内节点垂直间距

        c.setFont(font_name, 7.5)
        c.setFillColor(colors.HexColor('#6c757d'))
        for ci, label in enumerate(col_labels):
            c.drawCentredString(col_centers[ci], pg_note_y, label)
        c.setFillColor(colors.black)

        pg_top = pg_note_y - 5*mm  # 节点起始Y
        max_nodes_per_col = 8        # 最多8个（曾祖代）

        def get_pedigree_node(ear_tag):
            if not ear_tag or not ear_tag.strip():
                return None
            row = conn2.execute("SELECT ear_tag, sex, father_tag, mother_tag FROM sheep WHERE ear_tag=?", (ear_tag,)).fetchone()
            if not row:
                # 羊只表中不存在：返回仅含耳号的占位节点（无父/母信息）
                return {'ear_tag': ear_tag, 'sex': '', 'father': None, 'mother': None}
            return {'ear_tag': row['ear_tag'], 'sex': row['sex'],
                    'father': get_pedigree_node(row['father_tag']),
                    'mother': get_pedigree_node(row['mother_tag'])}

        conn2 = get_db()
        root = {
            'ear_tag': sheep['ear_tag'], 'sex': sheep['sex'],
            'father': get_pedigree_node(sheep['father_tag']),
            'mother': get_pedigree_node(sheep['mother_tag']),
        }
        conn2.close()

        # 按列精确收集节点（不递归重复，每代直接展开）
        collected = [[], [], [], []]

        def collect_children(source_nodes, dst_col):
            """把 source_nodes 列表中每个节点的 father/mother 加入 dst_col"""
            for node in source_nodes:
                if not node:
                    collected[dst_col].append(None)
                    collected[dst_col].append(None)
                    continue
                collected[dst_col].append(node.get('father'))
                collected[dst_col].append(node.get('mother'))

        collected[0] = [root]
        collect_children(collected[0], 1)                      # 亲代: [父, 母]
        collect_children(collected[1], 2)                      # 祖代: [爷,奶,姥,姥爷爷]
        collect_children(collected[2], 3)                      # 曾祖代: 8个

        # 计算各列总Y偏移（居中）
        import math
        def get_y_offsets(col_items, y_start, gap):
            """给定一列节点列表，返回各节点的中心Y（自顶向下）"""
            n = len(col_items)
            if n == 0:
                return []
            total_h = (n - 1) * gap
            top_y = y_start  # 最顶端Y
            return [top_y - i * gap for i in range(n)]

        # 计算每列高度
        col_heights = [len(c) * gen_gap_y for c in collected]
        max_height = max(col_heights) if col_heights else gen_gap_y
        # 已最大列高度居中
        tree_mid_y = pg_top - max_height / 2

        col_y_offsets = []
        for ci in range(4):
            items = collected[ci]
            n = len(items)
            if n == 0:
                col_y_offsets.append([])
                continue
            # 从 mid_y 向下分布
            start_y = tree_mid_y + (n - 1) * gen_gap_y / 2
            col_y_offsets.append([start_y - i * gen_gap_y for i in range(n)])

        # 画节点
        def draw_pg_node(cx, cy, node, w):
            if not node:
                return
            tag = node['ear_tag']
            sex = node.get('sex', '')
            c.setFont(font_name, 7)
            tw = c.stringWidth(tag, font_name, 7)
            bw = max(tw + 10, 24)
            bw = min(bw, w)
            bx = cx - bw / 2
            by = cy - node_h / 2
            if sex == '公':
                clr = male_clr
            elif sex == '母':
                clr = female_clr
            else:
                clr = colors.HexColor('#6c757d')  # 未知性别灰色
            # 仅边框，不填充
            c.setStrokeColor(clr)
            c.setLineWidth(1)
            c.roundRect(bx, by, bw, node_h, 3, fill=0)
            c.setFont(font_name, 7)
            c.setFillColor(clr)
            c.drawCentredString(cx, by + 2.6*mm, tag)
            c.setFillColor(colors.black)
            return bw

        # 存储节点位置用于画线
        node_positions = [[], [], [], []]  # [(node, cx, cy, bw), ...]

        for ci in range(4):
            cx = col_centers[ci]
            for ni, node in enumerate(collected[ci]):
                cy = col_y_offsets[ci][ni]
                bw = draw_pg_node(cx, cy, node, node_w)
                if node:
                    node_positions[ci].append((node, cx, cy, bw))

        # 连线: 父子间
        c.setStrokeColor(line_clr)
        c.setLineWidth(0.5)
        for ci in range(3):
            for src_node, sx, sy, sw in node_positions[ci]:
                father = src_node.get('father')
                mother = src_node.get('mother')
                if not father and not mother:
                    continue
                # 在下一列找到 father/mother
                for dst_node, dx, dy, dw in node_positions[ci + 1]:
                    for parent in [father, mother]:
                        if parent and dst_node is parent:
                            # 画: sx + sw/2 → dx - dw/2
                            c.line(sx + sw/2, sy, dx - dw/2, dy)
                            break

        # 羊只节点加粗强调框（仅加粗边框，不重复文字）
        if node_positions[0]:
            _, cx_self, cy_self, bw_self = node_positions[0][0]
            rct_w = bw_self + 6
            rct_h = node_h + 3
            clr = male_clr if root['sex'] == '公' else female_clr
            c.setStrokeColor(clr)
            c.setLineWidth(2.5)
            c.roundRect(cx_self - rct_w/2, cy_self - rct_h/2, rct_w, rct_h, 5, fill=0)

        # 计算最底部Y
        bottom_ys = []
        for ci in range(4):
            if node_positions[ci]:
                bottom_ys.append(node_positions[ci][-1][2] - node_h/2)
        y = min(bottom_ys) - 8*mm if bottom_ys else pg_top - max_height - 8*mm

        # 测定记录
        y -= 6*mm
        c.setFont(font_name, 14)
        c.drawString(left_margin, y, '三、测定记录')
        y -= 14*mm

        stage_order = ['初生', '3月龄', '6月龄', '12月龄']

        if sheep['sex'] == '公':
            headers = ['阶段', '体重(kg)', '体高', '体斜长', '胸围', '管围', '臀宽', '睾丸围', '睾丸直径']
            col_widths = [20*mm, 16*mm, 14*mm, 14*mm, 14*mm, 14*mm, 14*mm, 16*mm, 16*mm]
        else:
            headers = ['阶段', '体重(kg)', '体高', '体斜长', '胸围', '管围', '臀宽']
            col_widths = [22*mm, 18*mm, 16*mm, 16*mm, 16*mm, 16*mm, 16*mm]

        col_starts = [left_margin]
        for cw in col_widths[:-1]:
            col_starts.append(col_starts[-1] + cw)

        c.setFont(font_name, 9)
        header_y = y
        for i, h in enumerate(headers):
            c.drawString(col_starts[i], header_y, h)
        # 横线在表头下方
        line_y = header_y - 4*mm
        c.setStrokeColor(colors.HexColor('#333333'))
        c.setLineWidth(1)
        c.line(left_margin, line_y, left_margin + sum(col_widths), line_y)
        y = line_y - 3*mm

        def fmt(v):
            """格式化测定值：0或空值留空"""
            if v is None or v == 0 or v == '':
                return ''
            return str(v)

        for stage in stage_order:
            m = meas_map.get(stage, {})
            if sheep['sex'] == '公':
                vals = [
                    stage,
                    fmt(m.get('weight')), fmt(m.get('body_height')),
                    fmt(m.get('body_length')), fmt(m.get('chest_girth')),
                    fmt(m.get('cannon_circumference')), fmt(m.get('hip_width')),
                    fmt(m.get('testis_circumference')), fmt(m.get('testis_diameter'))
                ]
            else:
                vals = [
                    stage,
                    fmt(m.get('weight')), fmt(m.get('body_height')),
                    fmt(m.get('body_length')), fmt(m.get('chest_girth')),
                    fmt(m.get('cannon_circumference')), fmt(m.get('hip_width'))
                ]
            for i, v in enumerate(vals):
                c.drawString(col_starts[i], y, v)
            y -= 7*mm

        # 母羊的生产信息
        if sheep['sex'] == '母' and breeding:
            # 检查是否需要新页
            if y < 80*mm:
                c.showPage()
                y = height - 30*mm

            y -= 6*mm
            c.setFont(font_name, 14)
            c.drawString(left_margin, y, '四、生产信息')
            y -= 12*mm

            c.setFont(font_name, 9)
            prod_headers = ['胎次', '与配公羊', '产羔日期', '产羔数', '活羔数', '羔羊耳号']
            pcol_widths = [18*mm, 22*mm, 28*mm, 16*mm, 16*mm, 50*mm]
            pcol_starts = [left_margin]
            for cw in pcol_widths[:-1]:
                pcol_starts.append(pcol_starts[-1] + cw)

            for i, h in enumerate(prod_headers):
                c.drawString(pcol_starts[i], y, h)
            y -= 6*mm
            c.line(left_margin, y, right_margin, y)
            y -= 2*mm

            for idx, br in enumerate(breeding):
                lamb_str = ', '.join([f"{l['lamb_ear_tag']}({l['sex']})" for l in br['lambs']])
                vals = [
                    f'第{idx+1}胎',
                    str(br['ram_tag'] or '—'),
                    str(br['lambing_date']),
                    str(br['total_born']),
                    str(br['live_born']),
                    lamb_str
                ]
                for i, v in enumerate(vals):
                    c.drawString(pcol_starts[i], y, v)
                y -= 7*mm

        c.save()
        buf.seek(0)
        # 记录下载时间戳
        conn.execute("UPDATE sheep SET card_downloaded_at=CURRENT_TIMESTAMP WHERE id=?", (sheep_id,))
        conn.commit()
        conn.close()
        return send_file(buf, mimetype='application/pdf',
                        as_attachment=True,
                        download_name=f"{sheep['ear_tag']}_档案卡.pdf")
    except ImportError:
        return jsonify({'error': 'reportlab 未安装'}), 500

# ──────────────────────────────────────────────────────────
# 获取耳号列表(用于下拉选择)
# ──────────────────────────────────────────────────────────
@app.route('/api/sheep/tags')
@login_required
def api_sheep_tags():
    conn = get_db()
    q = request.args.get('q', '').strip()
    all_req = request.args.get('all', '')
    if q:
        tags = conn.execute("SELECT ear_tag, sex, id FROM sheep WHERE ear_tag LIKE ? ORDER BY ear_tag LIMIT 200",
                            (f'%{q}%',)).fetchall()
    elif all_req:
        tags = conn.execute("SELECT ear_tag, sex, id FROM sheep ORDER BY ear_tag").fetchall()
    else:
        tags = conn.execute("SELECT ear_tag, sex, id FROM sheep ORDER BY ear_tag LIMIT 300").fetchall()
    conn.close()
    return jsonify([{'id': t['id'], 'ear_tag': t['ear_tag'], 'sex': t['sex']} for t in tags])

# ──────────────────────────────────────────────────────────
# 品种管理 API
# ──────────────────────────────────────────────────────────
@app.route('/api/breeds')
@login_required
def api_breeds_list():
    conn = get_db()
    breeds = conn.execute("SELECT * FROM breeds ORDER BY id").fetchall()
    conn.close()
    return jsonify([dict(b) for b in breeds])

@app.route('/api/breeds', methods=['POST'])
@login_required
def api_breeds_create():
    data = request.get_json()
    name = data.get('name', '').strip()
    if not name:
        return jsonify({'success': False, 'message': '品种名不能为空'})
    conn = get_db()
    try:
        conn.execute("INSERT INTO breeds (name) VALUES (?)", (name,))
        conn.commit()
        return jsonify({'success': True, 'id': conn.execute("SELECT last_insert_rowid()").fetchone()[0]})
    except sqlite3.IntegrityError:
        return jsonify({'success': False, 'message': '品种名已存在'})
    finally:
        conn.close()

@app.route('/api/breeds/<int:breed_id>', methods=['PUT'])
@admin_required
def api_breeds_update(breed_id):
    data = request.get_json()
    name = data.get('name', '').strip()
    if not name:
        return jsonify({'success': False, 'message': '品种名不能为空'})
    conn = get_db()
    conn.execute("UPDATE breeds SET name=? WHERE id=?", (name, breed_id))
    conn.commit()
    conn.close()
    return jsonify({'success': True})

@app.route('/api/breeds/<int:breed_id>', methods=['DELETE'])
@admin_required
def api_breeds_delete(breed_id):
    conn = get_db()
    conn.execute("DELETE FROM breeds WHERE id=?", (breed_id,))
    conn.commit()
    conn.close()
    return jsonify({'success': True})

# ──────────────────────────────────────────────────────────
# 系谱树状图 API
# ──────────────────────────────────────────────────────────
@app.route('/api/sheep/<int:sheep_id>/pedigree')
@login_required
def api_sheep_pedigree(sheep_id):
    """获取三代系谱数据"""
    conn = get_db()
    sheep = conn.execute("SELECT * FROM sheep WHERE id=?", (sheep_id,)).fetchone()
    if not sheep:
        conn.close()
        return jsonify({'error': '未找到'}), 404

    def find_parents(ear_tag):
        if not ear_tag or not ear_tag.strip():
            return None
        row = conn.execute("SELECT * FROM sheep WHERE ear_tag=?", (ear_tag,)).fetchone()
        if not row:
            return {'ear_tag': ear_tag, 'sex': '', 'breed': '', 'birth_date': '', 'father_tag': '', 'mother_tag': ''}
        return dict(row)

    def build_generation(ear_tag):
        """递归获取祖先"""
        p = find_parents(ear_tag)
        if not p:
            return None
        node = {
            'ear_tag': p['ear_tag'],
            'sex': p['sex'],
            'breed': p['breed'],
            'birth_date': p['birth_date'],
        }
        node['father'] = build_generation(p['father_tag'])
        node['mother'] = build_generation(p['mother_tag'])
        return node

    pedigree = {
        'ear_tag': sheep['ear_tag'],
        'sex': sheep['sex'],
        'breed': sheep['breed'],
        'birth_date': sheep['birth_date'],
    }
    pedigree['father'] = build_generation(sheep['father_tag'])
    pedigree['mother'] = build_generation(sheep['mother_tag'])

    conn.close()
    return jsonify(pedigree)

# ──────────────────────────────────────────────────────────
# 启动
# ──────────────────────────────────────────────────────────
if __name__ == '__main__':
    # 端口与调试模式支持环境变量覆盖，方便同一份代码在本地(5002)与服务器(8081)共用
    PORT = int(os.environ.get('PORT', '5002'))
    DEBUG = os.environ.get('FLASK_DEBUG', 'True').lower() in ('1', 'true', 'yes')
    app.config['TEMPLATES_AUTO_RELOAD'] = DEBUG
    app.run(host='0.0.0.0', port=PORT, debug=DEBUG)
