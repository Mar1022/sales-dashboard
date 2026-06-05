import os
import sys
import json
import traceback
from datetime import datetime
from flask import Flask, render_template, request, jsonify
import sqlite3
import openpyxl

app = Flask(__name__)

# ── 错误日志 ──
@app.errorhandler(Exception)
def handle_error(e):
    print("=" * 50, flush=True)
    print(f"ERROR: {e}", flush=True)
    traceback.print_exc(file=sys.stdout)
    print("=" * 50, flush=True)
    return jsonify({"error": str(e)}), 500

@app.errorhandler(500)
def handle_500(e):
    print("=" * 50, flush=True)
    print("500 INTERNAL SERVER ERROR", flush=True)
    if hasattr(e, '__traceback__'):
        traceback.print_tb(e.__traceback__, file=sys.stdout)
    else:
        traceback.print_exc(file=sys.stdout)
    print("=" * 50, flush=True)
    return jsonify({"error": "服务器内部错误"}), 500
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data.db')
UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')

os.makedirs(UPLOAD_DIR, exist_ok=True)


# ── 默认新品规格清单 ──
DEFAULT_NEW_PRODS = [
    "AH100-1","BD-D75","BD-AJ112","AH100-2","BD-AJ117","BD-AJ115","ZT1","BD-Q120","BD-F38",
    "FPQ1","BD-D37","BD-Q12","BD-F37-1","BD-D69-2","DPR2","BD-H01-3","BD-D10-1","BD-D35",
    "BD-H01-10","BD-H01-9","BD-D38","DPD2","BD-G811","DPD1","DPC1","BD-H01-4","BD-LG91",
    "BD-LG10","FP16","DPC3","DPC2","DPR1","BD-H01-2","DPJ1","BD-H01-6","BD-D36","BD-H01-5",
    "BD-H01-8","AJ101","BD-H01-7","BD-H01-1","AJ102","AJ103","AJ104","FP1","FP2","FP3","FP4",
    "DPX1T","BD-L63","BD-D85","BD-ZT3T","BD-STR001","XPG1","XPG2","XPG3","XPG4","XPG5",
    "XPG6","XPS","X14","X15","BD-D39","BD-D32","BD-Q110","BD-Q130-1","BD-Q130-2","BD-F37-2",
    "AJ105","AJ106","AJ107","BD-AJ110","BD-AJ111","BD-AJ113","BD-AJ114","BD-AJ116",
    "AJ120","AJ121","AJ122","AJ123","XM1","XM2","XM3","BD-F51","BD-D69-1","BD-J47",
    "BD-J48","BD-A40D","BD-G341","BD-G812","BD-G813","BD-G814","BD-G815","BD-G816",
    "BD-G817","BD-G818","BD-G821","BD-G822","BD-G823","BD-G824","BD-G825","BD-G826",
    "BD-G827","BD-G831","BD-G832","BD-G833","BD-G834","BD-G835","BD-G836","BD-G837",
    "BD-G841","BD-G842","BD-G843","BD-G844","BD-G845","BD-G846","BD-G847","BD-GP3000",
    "BD-GP3010","BD-GP3020","BD-GP3040"
]


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    conn.execute('''
        CREATE TABLE IF NOT EXISTS sales_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            cust TEXT NOT NULL,
            product TEXT NOT NULL,
            spec TEXT NOT NULL,
            unit TEXT DEFAULT '',
            qty REAL DEFAULT 0,
            price REAL DEFAULT 0,
            amount REAL DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(date, cust, product, spec)
        )
    ''')
    conn.execute('''
        CREATE TABLE IF NOT EXISTS targets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            year INTEGER UNIQUE NOT NULL,
            amount REAL NOT NULL,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.execute('''
        CREATE TABLE IF NOT EXISTS new_products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            spec TEXT UNIQUE NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()

    # 首次初始化时写入默认新品规格
    cur = conn.execute("SELECT COUNT(*) FROM new_products")
    if cur.fetchone()[0] == 0:
        conn.executemany(
            "INSERT OR IGNORE INTO new_products (spec) VALUES (?)",
            [(s,) for s in DEFAULT_NEW_PRODS]
        )
        conn.commit()
    conn.close()


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/health')
def health():
    # 检查数据库是否正常
    try:
        conn = get_db()
        conn.execute("SELECT 1")
        conn.close()
        return jsonify({"status": "ok", "db": "connected"})
    except Exception as e:
        return jsonify({"status": "error", "db": str(e)}), 500


# ── 数据接口 ──
@app.route('/api/data')
def get_data():
    start = request.args.get('start', '')
    end = request.args.get('end', '')
    conn = get_db()
    sql = "SELECT date, cust, product, spec, unit, qty, price, amount FROM sales_records"
    params = []
    conditions = []
    if start:
        conditions.append("date >= ?")
        params.append(start)
    if end:
        conditions.append("date <= ?")
        params.append(end)
    if conditions:
        sql += " WHERE " + " AND ".join(conditions)
    sql += " ORDER BY date DESC"
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return jsonify([{
        "date": r["date"],
        "cust": r["cust"],
        "product": r["product"],
        "spec": r["spec"],
        "unit": r["unit"],
        "qty": r["qty"],
        "price": r["price"],
        "amount": r["amount"]
    } for r in rows])


@app.route('/api/summary')
def get_summary():
    start = request.args.get('start', '')
    end = request.args.get('end', '')
    prev_start = request.args.get('prev_start', '')
    prev_end = request.args.get('prev_end', '')
    year = request.args.get('year', '')

    conn = get_db()

    def _agg(s, e):
        if s and e:
            rows = conn.execute(
                "SELECT amount, cust, date FROM sales_records WHERE date >= ? AND date <= ?",
                (s, e)
            ).fetchall()
        elif s:
            rows = conn.execute(
                "SELECT amount, cust, date FROM sales_records WHERE date >= ?",
                (s,)
            ).fetchall()
        elif e:
            rows = conn.execute(
                "SELECT amount, cust, date FROM sales_records WHERE date <= ?",
                (e,)
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT amount, cust, date FROM sales_records"
            ).fetchall()
        total = sum(r["amount"] for r in rows)
        custs = list(set(r["cust"] for r in rows))
        return total, custs, rows

    cur_total, cur_custs, cur_rows = _agg(start, end)

    # 所有历史数据用于判断新老客户
    all_rows = conn.execute(
        "SELECT cust, MIN(date) as first_date FROM sales_records GROUP BY cust"
    ).fetchall()
    first_order = {r["cust"]: r["first_date"] for r in all_rows}

    s_val = start
    new_custs = [c for c in cur_custs if first_order.get(c, '') >= s_val] if s_val else cur_custs
    old_custs = [c for c in cur_custs if c not in new_custs]

    new_amt = sum(r["amount"] for r in cur_rows if r["cust"] in new_custs)
    old_amt = sum(r["amount"] for r in cur_rows if r["cust"] in old_custs)

    # 新品销售额
    np_rows = conn.execute("SELECT spec FROM new_products").fetchall()
    np_specs = set(r["spec"] for r in np_rows)
    np_amt = sum(r["amount"] for r in cur_rows if r["spec"] in np_specs)

    # 环比数据
    prev_total = 0
    if prev_start or prev_end:
        prev_total, _, _ = _agg(prev_start, prev_end)

    # 目标完成率
    target_val = 0
    if year:
        t = conn.execute(
            "SELECT amount FROM targets WHERE year = ?", (int(year),)
        ).fetchone()
        if t:
            target_val = t["amount"]

    conn.close()

    return jsonify({
        "total_amount": cur_total,
        "customer_count": len(cur_custs),
        "new_customer_count": len(new_custs),
        "old_customer_count": len(old_custs),
        "new_customer_amount": new_amt,
        "old_customer_amount": old_amt,
        "new_product_amount": np_amt,
        "prev_total_amount": prev_total,
        "target_amount": target_val
    })


# ── 上传接口 ──
@app.route('/api/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({"error": "未找到文件"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "文件名为空"}), 400

    filepath = os.path.join(UPLOAD_DIR, file.filename)
    file.save(filepath)

    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active
        # 读取表头
        headers = [cell.value for cell in ws[1]]
        idx_map = {}
        col_map = {
            '终审日期': 'date', '客户名称': 'cust', '品名': 'product',
            '货品规格': 'spec', '单位': 'unit', '受订数量': 'qty',
            '单价': 'price', '本位币': 'amount'
        }
        for zh, en in col_map.items():
            for i, h in enumerate(headers):
                if h and str(h).strip() == zh:
                    idx_map[en] = i
                    break

        required = ['date', 'cust', 'product', 'spec']
        for k in required:
            if k not in idx_map:
                os.remove(filepath)
                return jsonify({"error": f"缺少必需列：{k}"}), 400

        conn = get_db()
        total_rows = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            date_val = row[idx_map['date']] if idx_map.get('date') is not None else None
            cust_val = row[idx_map['cust']] if idx_map.get('cust') is not None else None
            product_val = row[idx_map['product']] if idx_map.get('product') is not None else None
            spec_val = row[idx_map['spec']] if idx_map.get('spec') is not None else None

            if not date_val or not cust_val or not product_val or not spec_val:
                continue

            # 日期处理
            if isinstance(date_val, datetime):
                date_str = date_val.strftime('%Y-%m-%d')
            else:
                date_str = str(date_val).strip()

            unit_val = str(row[idx_map.get('unit', -1)]) if idx_map.get('unit') is not None and row[idx_map['unit']] is not None else ''
            qty_val = float(row[idx_map.get('qty', -1)]) if idx_map.get('qty') is not None and row[idx_map.get('qty')] is not None else 0
            price_val = float(row[idx_map.get('price', -1)]) if idx_map.get('price') is not None and row[idx_map.get('price')] is not None else 0
            amount_val = float(row[idx_map.get('amount', -1)]) if idx_map.get('amount') is not None and row[idx_map.get('amount')] is not None else 0

            try:
                conn.execute('''
                    INSERT INTO sales_records (date, cust, product, spec, unit, qty, price, amount)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(date, cust, product, spec) DO UPDATE SET
                        unit=excluded.unit, qty=excluded.qty,
                        price=excluded.price, amount=excluded.amount
                ''', (date_str, str(cust_val).strip(), str(product_val).strip(),
                      str(spec_val).strip(), unit_val, qty_val, price_val, amount_val))
                total_rows += 1
            except Exception:
                continue

        conn.commit()
        conn.close()
    except Exception as e:
        if os.path.exists(filepath):
            os.remove(filepath)
        return jsonify({"error": f"解析失败：{str(e)}"}), 400
    finally:
        if os.path.exists(filepath):
            os.remove(filepath)

    return jsonify({"success": True, "count": total_rows})


# ── 目标接口 ──
@app.route('/api/target', methods=['GET', 'POST', 'DELETE'])
def target():
    conn = get_db()
    if request.method == 'GET':
        year = request.args.get('year', '')
        if year:
            row = conn.execute(
                "SELECT year, amount FROM targets WHERE year = ?", (int(year),)
            ).fetchone()
            conn.close()
            if row:
                return jsonify({"year": row["year"], "amount": row["amount"]})
            return jsonify({"year": int(year), "amount": 0})
        else:
            rows = conn.execute("SELECT year, amount FROM targets ORDER BY year").fetchall()
            conn.close()
            return jsonify({str(r["year"]): r["amount"] for r in rows})

    elif request.method == 'POST':
        data = request.get_json()
        year = int(data.get("year", 0))
        amount = float(data.get("amount", 0))
        if not year or amount <= 0:
            conn.close()
            return jsonify({"error": "无效参数"}), 400
        conn.execute(
            "INSERT INTO targets (year, amount, updated_at) VALUES (?, ?, datetime('now')) "
            "ON CONFLICT(year) DO UPDATE SET amount=excluded.amount, updated_at=datetime('now')",
            (year, amount)
        )
        conn.commit()
        conn.close()
        return jsonify({"success": True, "year": year, "amount": amount})

    elif request.method == 'DELETE':
        year = request.args.get('year', '')
        if year:
            conn.execute("DELETE FROM targets WHERE year = ?", (int(year),))
            conn.commit()
        conn.close()
        return jsonify({"success": True})


# ── 新品规格接口 ──
@app.route('/api/new-products', methods=['GET', 'POST'])
def new_products():
    conn = get_db()
    if request.method == 'GET':
        rows = conn.execute("SELECT spec FROM new_products ORDER BY id").fetchall()
        conn.close()
        return jsonify([r["spec"] for r in rows])

    elif request.method == 'POST':
        data = request.get_json()
        specs = data.get("specs", [])
        if isinstance(specs, list):
            conn.execute("DELETE FROM new_products")
            conn.executemany(
                "INSERT OR IGNORE INTO new_products (spec) VALUES (?)",
                [(s,) for s in specs if s and str(s).strip()]
            )
            conn.commit()
        conn.close()
        return jsonify({"success": True, "count": len(specs)})


@app.route('/api/new-products/<int:idx>', methods=['DELETE'])
def delete_new_product(idx):
    conn = get_db()
    conn.execute("DELETE FROM new_products WHERE id = ?", (idx,))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


# ── 启动 ──
if __name__ == '__main__':
    import os
    print(">>> Starting sales dashboard...", flush=True)
    print(f">>> DB path: {DB_PATH}", flush=True)
    print(f">>> Upload dir: {UPLOAD_DIR}", flush=True)
    try:
        init_db()
        print(">>> DB initialized OK", flush=True)
    except Exception as e:
        print(f">>> DB INIT FAILED: {e}", flush=True)
        traceback.print_exc()
    port = int(os.environ.get('PORT', 5000))
    print(f">>> Listening on 0.0.0.0:{port}", flush=True)
    app.run(host='0.0.0.0', port=port, debug=False)
