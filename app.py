import os
import sys
import json
import csv
import io
import traceback
from datetime import datetime, timedelta
from flask import Flask, render_template, request, jsonify
import psycopg2
import psycopg2.extras
from psycopg2.extras import execute_values
import openpyxl
from urllib.parse import urlparse

app = Flask(__name__)

# ── 数据库配置 ──
DATABASE_URL = os.environ.get('DATABASE_URL', '')
UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
os.makedirs(UPLOAD_DIR, exist_ok=True)



def get_db():
    """获取 PostgreSQL 连接"""
    conn = psycopg2.connect(DATABASE_URL)
    return conn


def get_cursor(conn):
    """获取字典游标"""
    return conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)


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


def init_db():
    """初始化 PostgreSQL 表结构（幂等）"""
    conn = get_db()
    cur = get_cursor(conn)
    try:
        cur.execute('''
            CREATE TABLE IF NOT EXISTS sales_records (
                id BIGSERIAL PRIMARY KEY,
                date TIMESTAMPTZ NOT NULL,
                cust TEXT NOT NULL,
                product TEXT NOT NULL,
                spec TEXT NOT NULL,
                unit TEXT DEFAULT '',
                qty REAL DEFAULT 0,
                price REAL DEFAULT 0,
                amount REAL DEFAULT 0,
                created_at TIMESTAMPTZ DEFAULT NOW(),
                UNIQUE(date, cust, product, spec, unit, qty, price, amount)
            )
        ''')
        cur.execute('''
            CREATE TABLE IF NOT EXISTS targets (
                id BIGSERIAL PRIMARY KEY,
                year INTEGER UNIQUE NOT NULL,
                amount REAL NOT NULL,
                updated_at TIMESTAMPTZ DEFAULT NOW()
            )
        ''')
        cur.execute('''
            CREATE TABLE IF NOT EXISTS new_products (
                id BIGSERIAL PRIMARY KEY,
                spec TEXT UNIQUE NOT NULL,
                created_at TIMESTAMPTZ DEFAULT NOW()
            )
        ''')

        cur.execute('''
            CREATE TABLE IF NOT EXISTS customers (
                id BIGSERIAL PRIMARY KEY,
                name TEXT UNIQUE NOT NULL,
                code TEXT,
                level CHAR(1) DEFAULT 'D',
                last_followup DATE,
                next_followup DATE,
                created_at TIMESTAMPTZ DEFAULT NOW(),
                updated_at TIMESTAMPTZ DEFAULT NOW()
            )
        ''')

        cur.execute('CREATE INDEX IF NOT EXISTS idx_customers_name ON customers(name)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_customers_level ON customers(level)')

        # 首次初始化时写入默认新品规格
        cur.execute("SELECT COUNT(*) AS cnt FROM new_products")
        if cur.fetchone()["cnt"] == 0:
            for s in DEFAULT_NEW_PRODS:
                cur.execute(
                    "INSERT INTO new_products (spec) VALUES (%s) ON CONFLICT (spec) DO NOTHING",
                    (s,)
                )

        conn.commit()
    finally:
        cur.close()
        conn.close()


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/health')
def health():
    try:
        conn = get_db()
        cur = get_cursor(conn)
        cur.execute("SELECT 1 AS ok")
        cur.close()
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
    cur = get_cursor(conn)
    try:
        sql = "SELECT date, cust, product, spec, unit, qty, price, amount FROM sales_records"
        params = []
        conditions = []
        if start:
            conditions.append("date >= %s")
            params.append(start)
        if end:
            conditions.append("date <= %s")
            params.append(end if ' ' in end else end + ' 23:59:59')
        if conditions:
            sql += " WHERE " + " AND ".join(conditions)
        sql += " ORDER BY date DESC"
        cur.execute(sql, params)
        rows = cur.fetchall()
        return jsonify([{
            "date": r["date"].strftime('%Y-%m-%d %H:%M:%S') if hasattr(r["date"], 'strftime') else str(r["date"])[:19],
            "cust": r["cust"],
            "product": r["product"],
            "spec": r["spec"],
            "unit": r["unit"],
            "qty": float(r["qty"]),
            "price": float(r["price"]),
            "amount": float(r["amount"])
        } for r in rows])
    finally:
        cur.close()
        conn.close()


@app.route('/api/summary')
def get_summary():
    start = request.args.get('start', '')
    end = request.args.get('end', '')
    prev_start = request.args.get('prev_start', '')
    prev_end = request.args.get('prev_end', '')
    year = request.args.get('year', '')

    conn = get_db()
    cur = get_cursor(conn)
    try:
        def _agg(s, e):
            # 结束日期若无时间则补 23:59:59，确保当天数据不遗漏
            e_fixed = (e if (' ' in e) else (e + ' 23:59:59')) if e else None
            if s and e_fixed:
                cur.execute(
                    "SELECT amount, cust, date FROM sales_records WHERE date >= %s AND date <= %s",
                    (s, e_fixed)
                )
            elif s:
                cur.execute(
                    "SELECT amount, cust, date FROM sales_records WHERE date >= %s",
                    (s,)
                )
            elif e_fixed:
                cur.execute(
                    "SELECT amount, cust, date FROM sales_records WHERE date <= %s",
                    (e_fixed,)
                )
            else:
                cur.execute("SELECT amount, cust, date FROM sales_records")
            rows = cur.fetchall()
            total = sum(float(r["amount"]) for r in rows)
            custs = list(set(r["cust"] for r in rows))
            return total, custs, rows

        cur_total, cur_custs, cur_rows = _agg(start, end)

        # 所有历史数据用于判断新老客户
        cur.execute("SELECT cust, MIN(date) as first_date FROM sales_records GROUP BY cust")
        all_rows = cur.fetchall()
        first_order = {r["cust"]: str(r["first_date"]) for r in all_rows}

        s_val = start
        new_custs = [c for c in cur_custs if first_order.get(c, '9999') >= s_val] if s_val else cur_custs
        old_custs = [c for c in cur_custs if c not in new_custs]

        new_amt = sum(float(r["amount"]) for r in cur_rows if r["cust"] in new_custs)
        old_amt = sum(float(r["amount"]) for r in cur_rows if r["cust"] in old_custs)

        # 新品销售额
        cur.execute("SELECT spec FROM new_products")
        np_rows = cur.fetchall()
        np_specs = set(r["spec"] for r in np_rows)
        np_amt = sum(float(r["amount"]) for r in cur_rows if r["spec"] in np_specs)

        # 环比数据
        prev_total = 0
        if prev_start or prev_end:
            prev_total, _, _ = _agg(prev_start, prev_end)

        # 目标完成率
        target_val = 0
        if year:
            cur.execute("SELECT amount FROM targets WHERE year = %s", (int(year),))
            t = cur.fetchone()
            if t:
                target_val = float(t["amount"])

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
    finally:
        cur.close()
        conn.close()


# ── 安全转换 ──
def safe_float(val, default=0.0):
    """安全转浮点数，空值/空字符串返回默认值"""
    if val is None:
        return default
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if s == '':
        return default
    try:
        return float(s)
    except ValueError:
        return default


def normalize_date(val):
    """规范化日期为 YYYY-MM-DD HH:MM:SS 格式（兼容 PostgreSQL TIMESTAMPTZ）"""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.strftime('%Y-%m-%d %H:%M:%S')

    s = str(val).strip()
    if not s:
        return None

    # 处理 Excel 日期序列号（如 46175 = 2026-06-02）
    try:
        if s.isdigit() and 1 <= int(s) <= 100000:
            serial = int(s)
            excel_epoch = datetime(1899, 12, 30)
            return (excel_epoch + timedelta(days=serial)).strftime('%Y-%m-%d 00:00:00')
    except (ValueError, OverflowError):
        pass

    # 尝试带时间的格式
    for fmt in ['%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S',
                '%Y-%m-%d %H:%M', '%Y/%m/%d %H:%M',
                '%Y-%m-%dT%H:%M:%S', '%Y-%m-%dT%H:%M']:
        try:
            return datetime.strptime(s, fmt).strftime('%Y-%m-%d %H:%M:%S')
        except ValueError:
            continue

    # 尝试纯日期格式（无时间则补 00:00:00）
    for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%m/%d/%Y', '%d/%m/%Y']:
        try:
            return datetime.strptime(s, fmt).strftime('%Y-%m-%d 00:00:00')
        except ValueError:
            continue

    # 兜底：正则补零后再试（处理 2023/3/5 这种非补零日期）
    import re
    padded = re.sub(r'(?<=[-/.])(\d)(?=[-/.]|$)', r'0\1', s)
    for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d']:
        try:
            return datetime.strptime(padded, fmt).strftime('%Y-%m-%d 00:00:00')
        except ValueError:
            continue

    # 最终回退：返回原始值
    return s


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

    conn = None
    cur = None
    total_rows = 0
    skip_count = 0
    sample_empty = []
    first_headers = []
    try:
        # 流式读取，不加载全部到内存
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        ws = wb.active

        # 读表头（第一行）
        first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if first_row is None:
            os.remove(filepath)
            return jsonify({"error": "文件为空"}), 400
        headers = [cell for cell in first_row]
        first_headers = [str(h) for h in headers[:10]]

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

        required = ['date', 'cust', 'product']
        for k in required:
            if k not in idx_map:
                wb.close()
                os.remove(filepath)
                return jsonify({"error": f"缺少必需列：{k}"}), 400

        conn = get_db()
        cur = get_cursor(conn)

        # 批量插入缓冲区
        batch = []
        BATCH_SIZE = 2000
        row_idx = 1  # 从第2行开始

        INSERT_SQL = '''
            INSERT INTO sales_records AS t (date, cust, product, spec, unit, qty, price, amount)
            VALUES %s
            ON CONFLICT(date, cust, product, spec, unit, qty, price, amount) DO NOTHING
        '''

        insert_errors = []
        batch_dup_count = 0
        def flush_batch():
            nonlocal total_rows
            if not batch:
                return
            # 批次内去重：同 (date, cust, product, spec) 保留最后一条
            seen = {}
            for item in reversed(batch):
                key = (item[0], item[1], item[2], item[3], item[4], item[5], item[6], item[7])
                if key not in seen:
                    seen[key] = item
            deduped = list(reversed(list(seen.values())))
            nonlocal batch_dup_count
            batch_dup_count += len(batch) - len(deduped)
            try:
                execute_values(cur, INSERT_SQL, deduped, page_size=BATCH_SIZE)
                total_rows += len(deduped)
                conn.commit()
            except Exception as e:
                conn.rollback()
                insert_errors.append(str(e))
                print(f"[Excel Upload] 批量插入失败: {e}", flush=True)
                traceback.print_exc(file=sys.stdout)
            batch.clear()

        for row in ws.iter_rows(min_row=2, values_only=True):
            row_idx += 1
            date_val = row[idx_map['date']] if idx_map.get('date') is not None else None
            cust_val = row[idx_map['cust']] if idx_map.get('cust') is not None else None
            product_val = row[idx_map['product']] if idx_map.get('product') is not None else None
            spec_val = row[idx_map['spec']] if idx_map.get('spec') is not None else None

            if not date_val or not cust_val or not product_val:
                skip_count += 1
                if len(sample_empty) < 3:
                    sample_empty.append({
                        "row_idx": row_idx,
                        "date": repr(date_val),
                        "cust": repr(cust_val),
                        "product": repr(product_val),
                        "spec": repr(spec_val)
                    })
                continue

            # 日期处理
            date_str = normalize_date(date_val)

            unit_val = row[idx_map.get('unit', -1)] if idx_map.get('unit') is not None and idx_map.get('unit') < len(row) and row[idx_map['unit']] is not None else ''
            if unit_val is None:
                unit_val = ''
            else:
                unit_val = str(unit_val)
            qty_val = safe_float(row[idx_map.get('qty', -1)]) if idx_map.get('qty') is not None else 0
            price_val = safe_float(row[idx_map.get('price', -1)]) if idx_map.get('price') is not None else 0
            amount_val = safe_float(row[idx_map.get('amount', -1)]) if idx_map.get('amount') is not None else 0

            batch.append((date_str, str(cust_val).strip(), str(product_val).strip(),
                         str(spec_val).strip(), unit_val, qty_val, price_val, amount_val))

            if len(batch) >= BATCH_SIZE:
                flush_batch()

        flush_batch()

        wb.close()
    except Exception as e:
        if os.path.exists(filepath):
            os.remove(filepath)
        return jsonify({"error": f"解析失败：{str(e)}"}), 400
    finally:
        if cur is not None:
            cur.close()
        if conn is not None:
            conn.close()
        if os.path.exists(filepath):
            os.remove(filepath)

    sync_customers_from_sales()
    refresh_all_customer_levels()
    refresh_all_customer_followup()

    return jsonify({
        "success": True,
        "count": total_rows,
        "total_rows_in_file": total_rows + skip_count,
        "skipped": skip_count,
        "sample_skipped": sample_empty,
        "idx_map": {k: v for k, v in idx_map.items()},
        "header_count": len(headers),
        "first_3_headers": [str(h) for h in headers[:10]],
        "insert_errors": insert_errors[:3] if insert_errors else [],
        "batch_dup_removed": batch_dup_count
    })


# ── CSV 上传接口（流式读取，适合大数据文件）──
@app.route('/api/upload-csv', methods=['POST'])
def upload_csv():
    if 'file' not in request.files:
        return jsonify({"error": "未找到文件"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "文件名为空"}), 400

    # 读取文件原始内容，尝试多种编码
    raw_bytes = file.read()
    text = None
    encoding_used = 'utf-8-sig'
    for enc in ['utf-8-sig', 'utf-8', 'gbk', 'gb2312']:
        try:
            text = raw_bytes.decode(enc)
            encoding_used = enc
            break
        except (UnicodeDecodeError, UnicodeError):
            continue

    if text is None:
        return jsonify({"error": "无法识别文件编码，请保存为 UTF-8 格式的 CSV"}), 400

    # 释放原始字节，只保留解码后的文本
    del raw_bytes

    reader = csv.reader(io.StringIO(text))
    headers_row = next(reader, None)
    if headers_row is None:
        return jsonify({"error": "文件为空"}), 400

    headers = [str(h).strip() for h in headers_row]

    # 列名映射（与 Excel 上传保持一致）
    col_map = {
        '终审日期': 'date', '客户名称': 'cust', '品名': 'product',
        '货品规格': 'spec', '单位': 'unit', '受订数量': 'qty',
        '单价': 'price', '本位币': 'amount'
    }
    idx_map = {}
    for zh, en in col_map.items():
        for i, h in enumerate(headers):
            if h == zh:
                idx_map[en] = i
                break

    required = ['date', 'cust', 'product']
    for k in required:
        if k not in idx_map:
            return jsonify({"error": f"缺少必需列：{k}（当前表头：{headers[:10]}）"}), 400

    conn = None
    cur = None
    total_rows = 0
    skip_count = 0
    sample_empty = []
    batch = []
    BATCH_SIZE = 2000

    INSERT_SQL = '''
        INSERT INTO sales_records AS t (date, cust, product, spec, unit, qty, price, amount)
        VALUES %s
        ON CONFLICT(date, cust, product, spec, unit, qty, price, amount) DO NOTHING
    '''

    insert_errors = []
    batch_dup_count = 0  # 统计批次内去重数量
    def flush_batch():
        nonlocal total_rows
        if not batch:
            return
        # 批次内去重：同 (date, cust, product, spec) 保留最后一条
        seen = {}
        for item in reversed(batch):
            key = (item[0], item[1], item[2], item[3], item[4], item[5], item[6], item[7])
            if key not in seen:
                seen[key] = item
        deduped = list(reversed(list(seen.values())))
        nonlocal batch_dup_count
        batch_dup_count += len(batch) - len(deduped)
        try:
            execute_values(cur, INSERT_SQL, deduped, page_size=BATCH_SIZE)
            total_rows += len(deduped)
            conn.commit()
        except Exception as e:
            conn.rollback()
            insert_errors.append(str(e))
            print(f"[CSV Upload] 批量插入失败: {e}", flush=True)
            traceback.print_exc(file=sys.stdout)
        batch.clear()

    try:
        conn = get_db()
        cur = get_cursor(conn)

        for row_idx, row in enumerate(reader, start=2):
            # 跳过空行
            if not row or all(not str(c).strip() for c in row):
                skip_count += 1
                continue

            # 安全检查：确保行有足够的列
            max_idx = max(idx_map.values())
            if len(row) <= max_idx:
                skip_count += 1
                if len(sample_empty) < 3:
                    sample_empty.append({
                        "row_idx": row_idx,
                        "reason": f"列数不足（需要{max_idx+1}列，实际{len(row)}列）"
                    })
                continue

            date_val = row[idx_map['date']].strip() if idx_map.get('date') is not None else ''
            cust_val = row[idx_map['cust']].strip() if idx_map.get('cust') is not None else ''
            product_val = row[idx_map['product']].strip() if idx_map.get('product') is not None else ''
            spec_val = row[idx_map['spec']].strip() if idx_map.get('spec') is not None else ''

            if not date_val or not cust_val or not product_val:
                skip_count += 1
                if len(sample_empty) < 3:
                    sample_empty.append({
                        "row_idx": row_idx,
                        "date": repr(date_val),
                        "cust": repr(cust_val),
                        "product": repr(product_val),
                        "spec": repr(spec_val)
                    })
                continue

            # 日期格式处理
            date_str = normalize_date(date_val)

            unit_val = row[idx_map['unit']].strip() if idx_map.get('unit') is not None and idx_map['unit'] < len(row) else ''
            qty_val = safe_float(row[idx_map['qty']]) if idx_map.get('qty') is not None and idx_map['qty'] < len(row) else 0
            price_val = safe_float(row[idx_map['price']]) if idx_map.get('price') is not None and idx_map['price'] < len(row) else 0
            amount_val = safe_float(row[idx_map['amount']]) if idx_map.get('amount') is not None and idx_map['amount'] < len(row) else 0

            batch.append((date_str, cust_val, product_val, spec_val,
                         unit_val, qty_val, price_val, amount_val))

            if len(batch) >= BATCH_SIZE:
                flush_batch()

        flush_batch()

    except Exception as e:
        return jsonify({"error": f"解析失败：{str(e)}"}), 400
    finally:
        if cur is not None:
            cur.close()
        if conn is not None:
            conn.close()

    sync_customers_from_sales()
    refresh_all_customer_levels()
    refresh_all_customer_followup()

    return jsonify({
        "success": True,
        "count": total_rows,
        "total_rows_in_file": total_rows + skip_count,
        "skipped": skip_count,
        "sample_skipped": sample_empty,
        "encoding": encoding_used,
        "header_count": len(headers),
        "headers": headers[:10],
        "insert_errors": insert_errors[:3] if insert_errors else [],
        "batch_dup_removed": batch_dup_count
    })


# ── 目标接口 ──
@app.route('/api/target', methods=['GET', 'POST', 'DELETE'])
def target():
    conn = get_db()
    cur = get_cursor(conn)
    try:
        if request.method == 'GET':
            year = request.args.get('year', '')
            if year:
                cur.execute("SELECT year, amount FROM targets WHERE year = %s", (int(year),))
                row = cur.fetchone()
                if row:
                    return jsonify({"year": row["year"], "amount": float(row["amount"])})
                return jsonify({"year": int(year), "amount": 0})
            else:
                cur.execute("SELECT year, amount FROM targets ORDER BY year")
                rows = cur.fetchall()
                return jsonify({str(r["year"]): float(r["amount"]) for r in rows})

        elif request.method == 'POST':
            data = request.get_json()
            year = int(data.get("year", 0))
            amount = float(data.get("amount", 0))
            if not year or amount <= 0:
                return jsonify({"error": "无效参数"}), 400
            cur.execute(
                "INSERT INTO targets (year, amount, updated_at) VALUES (%s, %s, NOW()) "
                "ON CONFLICT(year) DO UPDATE SET amount = EXCLUDED.amount, updated_at = NOW()",
                (year, amount)
            )
            conn.commit()
            return jsonify({"success": True, "year": year, "amount": amount})

        elif request.method == 'DELETE':
            year = request.args.get('year', '')
            if year:
                cur.execute("DELETE FROM targets WHERE year = %s", (int(year),))
                conn.commit()
            return jsonify({"success": True})
    finally:
        cur.close()
        conn.close()


# ── 新品规格接口 ──
@app.route('/api/new-products', methods=['GET', 'POST'])
def new_products():
    conn = get_db()
    cur = get_cursor(conn)
    try:
        if request.method == 'GET':
            cur.execute("SELECT id, spec FROM new_products ORDER BY id")
            rows = cur.fetchall()
            return jsonify([{"id": r["id"], "spec": r["spec"]} for r in rows])

        elif request.method == 'POST':
            data = request.get_json()
            specs = data.get("specs", [])
            if isinstance(specs, list):
                cur.execute("DELETE FROM new_products")
                for s in specs:
                    if s and str(s).strip():
                        cur.execute(
                            "INSERT INTO new_products (spec) VALUES (%s) ON CONFLICT DO NOTHING",
                            (str(s).strip(),)
                        )
                conn.commit()
            return jsonify({"success": True, "count": len(specs)})
    finally:
        cur.close()
        conn.close()


@app.route('/api/new-products/<int:idx>', methods=['DELETE'])
def delete_new_product(idx):
    conn = get_db()
    cur = get_cursor(conn)
    try:
        cur.execute("DELETE FROM new_products WHERE id = %s", (idx,))
        conn.commit()
        return jsonify({"success": True})
    finally:
        cur.close()
        conn.close()


# ==================== 客户管理与跟进任务模块 ====================

FOLLOWUP_DAYS = {'A': 7, 'B': 14, 'C': 30, 'D': 30}
RECOVERY_DAYS = 60


def get_customer_year_sales(cust_name, year):
    """获取客户指定年份的销售额"""
    conn = get_db()
    cur = get_cursor(conn)
    try:
        start = f"{year}-01-01"
        end = f"{year}-12-31 23:59:59"
        cur.execute(
            "SELECT COALESCE(SUM(amount), 0) as total FROM sales_records WHERE cust = %s AND date >= %s AND date <= %s",
            (cust_name, start, end)
        )
        result = cur.fetchone()
        return float(result['total']) if result else 0
    finally:
        cur.close()
        conn.close()


def get_customer_level_by_sales(sales):
    """根据销售额返回等级"""
    if sales >= 500000:
        return 'A'
    elif sales >= 100000:
        return 'B'
    elif sales > 0:
        return 'C'
    else:
        return 'D'


def get_customer_peak_level(cust_name):
    """获取客户近3年内的最高等级"""
    current_year = datetime.now().year
    years = [current_year - 2, current_year - 1, current_year]
    peak = 'D'
    level_order = {'A': 4, 'B': 3, 'C': 2, 'D': 1}
    for y in years:
        sales = get_customer_year_sales(cust_name, y)
        level = get_customer_level_by_sales(sales)
        if level_order.get(level, 0) > level_order.get(peak, 0):
            peak = level
    return peak


def refresh_all_customer_levels():
    """刷新所有客户的等级（基于上年销售额）"""
    conn = get_db()
    cur = get_cursor(conn)
    try:
        last_year = datetime.now().year - 1
        cur.execute("SELECT DISTINCT cust FROM sales_records")
        customers = cur.fetchall()
        for row in customers:
            cust_name = row['cust']
            sales = get_customer_year_sales(cust_name, last_year)
            level = get_customer_level_by_sales(sales)
            cur.execute(
                "UPDATE customers SET level = %s, updated_at = NOW() WHERE name = %s",
                (level, cust_name)
            )
        conn.commit()
    finally:
        cur.close()
        conn.close()


def calculate_next_followup(level, peak_level, last_followup):
    """计算下次跟进日期"""
    from datetime import date, timedelta
    today = date.today()
    if last_followup is None:
        return today
    base_days = FOLLOWUP_DAYS.get(level, 30)
    next_date = last_followup + timedelta(days=base_days)
    if peak_level in ['A', 'B'] and level in ['C', 'D']:
        recovery_date = last_followup + timedelta(days=RECOVERY_DAYS)
        if recovery_date < next_date:
            next_date = recovery_date
    if next_date < today:
        next_date = today
    return next_date


def refresh_all_customer_followup():
    """刷新所有客户的下次跟进日期"""
    conn = get_db()
    cur = get_cursor(conn)
    try:
        cur.execute("SELECT name, level, last_followup FROM customers")
        customers = cur.fetchall()
        for row in customers:
            cust_name = row['name']
            level = row['level']
            last_followup = row['last_followup']
            peak_level = get_customer_peak_level(cust_name)
            next_followup = calculate_next_followup(level, peak_level, last_followup)
            cur.execute(
                "UPDATE customers SET next_followup = %s, updated_at = NOW() WHERE name = %s",
                (next_followup, cust_name)
            )
        conn.commit()
    finally:
        cur.close()
        conn.close()


def sync_customers_from_sales():
    """从销售记录同步客户到customers表"""
    conn = get_db()
    cur = get_cursor(conn)
    try:
        cur.execute("SELECT DISTINCT cust FROM sales_records")
        rows = cur.fetchall()
        for row in rows:
            cust_name = row['cust']
            cur.execute(
                "INSERT INTO customers (name, code, level, last_followup, next_followup) VALUES (%s, NULL, 'D', NULL, NULL) ON CONFLICT (name) DO NOTHING",
                (cust_name,)
            )
        conn.commit()
    finally:
        cur.close()
        conn.close()


# ==================== 新增API接口 ====================

@app.route('/api/customers/stats')
def get_customer_stats():
    """获取各等级客户数量"""
    year = request.args.get('year', '')
    if not year:
        year = str(datetime.now().year - 1)
    else:
        year = str(int(year))
    conn = get_db()
    cur = get_cursor(conn)
    try:
        cur.execute("SELECT name FROM customers")
        customers = cur.fetchall()
        stats = {'A': 0, 'B': 0, 'C': 0, 'D': 0}
        for row in customers:
            sales = get_customer_year_sales(row['name'], year)
            level = get_customer_level_by_sales(sales)
            stats[level] = stats.get(level, 0) + 1
        return jsonify({'success': True, 'data': stats, 'year': year})
    finally:
        cur.close()
        conn.close()


@app.route('/api/customers/by-level')
def get_customers_by_level():
    """获取某等级客户列表（分页）"""
    level = request.args.get('level', '')
    year = request.args.get('year', '')
    page = int(request.args.get('page', 1))
    limit = int(request.args.get('limit', 20))
    if not year:
        year = str(datetime.now().year - 1)
    if level not in ['A', 'B', 'C', 'D']:
        return jsonify({'error': '无效等级'}), 400
    conn = get_db()
    cur = get_cursor(conn)
    try:
        cur.execute("SELECT name, code FROM customers")
        customers = cur.fetchall()
        result = []
        for row in customers:
            sales = get_customer_year_sales(row['name'], year)
            cust_level = get_customer_level_by_sales(sales)
            if cust_level == level:
                result.append({
                    'name': row['name'],
                    'code': row['code'],
                    'sales': sales
                })
        result.sort(key=lambda x: x['sales'], reverse=True)
        total = len(result)
        offset = (page - 1) * limit
        paginated = result[offset:offset + limit]
        return jsonify({
            'success': True,
            'data': paginated,
            'total': total,
            'page': page,
            'limit': limit,
            'year': year
        })
    finally:
        cur.close()
        conn.close()


@app.route('/api/customers/followup/today')
def get_today_tasks():
    """获取今日需跟进的客户"""
    from datetime import date
    today = date.today()
    conn = get_db()
    cur = get_cursor(conn)
    try:
        cur.execute(
            "SELECT name, level, last_followup, next_followup FROM customers WHERE next_followup <= %s ORDER BY next_followup ASC, FIELD(level, 'A', 'B', 'C', 'D')",
            (today,)
        )
        tasks = cur.fetchall()
        result = []
        for row in tasks:
            peak_level = get_customer_peak_level(row['name'])
            result.append({
                'name': row['name'],
                'level': row['level'],
                'peak_level': peak_level,
                'last_followup': row['last_followup'].isoformat() if row['last_followup'] else None,
                'next_followup': row['next_followup'].isoformat() if row['next_followup'] else None,
                'is_overdue': row['next_followup'] < today if row['next_followup'] else False
            })
        return jsonify({'success': True, 'data': result})
    finally:
        cur.close()
        conn.close()


@app.route('/api/customers/followup/week')
def get_week_tasks():
    """获取本周需跟进的客户"""
    from datetime import date, timedelta
    today = date.today()
    week_end = today + timedelta(days=7)
    conn = get_db()
    cur = get_cursor(conn)
    try:
        cur.execute(
            "SELECT name, level, last_followup, next_followup FROM customers WHERE next_followup <= %s AND next_followup > %s ORDER BY next_followup ASC",
            (week_end, today)
        )
        tasks = cur.fetchall()
        result = []
        for row in tasks:
            result.append({
                'name': row['name'],
                'level': row['level'],
                'last_followup': row['last_followup'].isoformat() if row['last_followup'] else None,
                'next_followup': row['next_followup'].isoformat() if row['next_followup'] else None
            })
        return jsonify({'success': True, 'data': result})
    finally:
        cur.close()
        conn.close()


@app.route('/api/customers/followup/mark', methods=['POST'])
def mark_followup():
    """标记客户已跟进"""
    from datetime import date
    data = request.get_json()
    cust_name = data.get('customer_name', '')
    if not cust_name:
        return jsonify({'error': '缺少客户名称'}), 400
    conn = get_db()
    cur = get_cursor(conn)
    try:
        cur.execute("SELECT level FROM customers WHERE name = %s", (cust_name,))
        row = cur.fetchone()
        if not row:
            return jsonify({'error': '客户不存在'}), 404
        today = date.today()
        peak_level = get_customer_peak_level(cust_name)
        next_followup = calculate_next_followup(row['level'], peak_level, today)
        cur.execute(
            "UPDATE customers SET last_followup = %s, next_followup = %s, updated_at = NOW() WHERE name = %s",
            (today, next_followup, cust_name)
        )
        conn.commit()
        return jsonify({'success': True, 'message': '已标记跟进', 'next_followup': next_followup.isoformat()})
    finally:
        cur.close()
        conn.close()


@app.route('/api/customers/followup/batch', methods=['POST'])
def batch_mark_followup():
    """批量标记跟进"""
    from datetime import date
    data = request.get_json()
    level = data.get('level', '')
    conn = get_db()
    cur = get_cursor(conn)
    try:
        today = date.today()
        if level and level in ['A', 'B', 'C', 'D']:
            cur.execute("SELECT name, level FROM customers WHERE level = %s", (level,))
        else:
            cur.execute("SELECT name, level FROM customers WHERE next_followup <= %s", (today,))
        customers = cur.fetchall()
        count = 0
        for row in customers:
            peak_level = get_customer_peak_level(row['name'])
            next_followup = calculate_next_followup(row['level'], peak_level, today)
            cur.execute(
                "UPDATE customers SET last_followup = %s, next_followup = %s, updated_at = NOW() WHERE name = %s",
                (today, next_followup, row['name'])
            )
            count += 1
        conn.commit()
        return jsonify({'success': True, 'message': f'已批量标记{count}个客户'})
    finally:
        cur.close()
        conn.close()


@app.route('/api/customers/import-codes', methods=['POST'])
def import_codes():
    """导入编码对照表（Excel）"""
    if 'file' not in request.files:
        return jsonify({'error': '未找到文件'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '文件名为空'}), 400
    import openpyxl
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb.active
    conn = get_db()
    cur = get_cursor(conn)
    try:
        updated = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 2:
                continue
            name = str(row[0]).strip() if row[0] else ''
            code = str(row[1]).strip() if row[1] else ''
            if name and code:
                cur.execute(
                    "UPDATE customers SET code = %s, updated_at = NOW() WHERE name = %s",
                    (code, name)
                )
                if cur.rowcount > 0:
                    updated += 1
        conn.commit()
        return jsonify({'success': True, 'message': f'已更新{updated}个客户的编码'})
    finally:
        cur.close()
        conn.close()


@app.route('/api/customers/export-pending-codes')
def export_pending_codes():
    """导出待编码客户"""
    import csv
    import io
    conn = get_db()
    cur = get_cursor(conn)
    try:
        cur.execute("SELECT name, code FROM customers WHERE code IS NULL OR code = '' ORDER BY name")
        rows = cur.fetchall()
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow(['客户名称', '编码'])
        for row in rows:
            writer.writerow([row['name'], ''])
        output.seek(0)
        from flask import Response
        return Response(
            output.getvalue(),
            mimetype='text/csv',
            headers={'Content-Disposition': 'attachment; filename=pending_codes.csv'}
        )
    finally:
        cur.close()
        conn.close()


# ── 启动：gunicorn 不走 __main__，需在模块加载时初始化数据库 ──
if DATABASE_URL:
    try:
        init_db()
        print(">>> DB initialized OK (module load)", flush=True)
    except Exception as e:
        print(f">>> DB INIT FAILED: {e}", flush=True)
        traceback.print_exc()

if __name__ == '__main__':
    print(">>> Starting sales dashboard...", flush=True)
    print(f">>> DATABASE_URL: {'***configured***' if DATABASE_URL else 'MISSING!'}", flush=True)
    try:
        init_db()
        print(">>> DB initialized OK", flush=True)
    except Exception as e:
        print(f">>> DB INIT FAILED: {e}", flush=True)
        traceback.print_exc()
    port = int(os.environ.get('PORT', 5000))
    print(f">>> Listening on 0.0.0.0:{port}", flush=True)
    app.run(host='0.0.0.0', port=port, debug=False)