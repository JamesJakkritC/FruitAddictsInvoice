# -*- coding: utf-8 -*-
"""
Fruit Addicts - Google Sheets Sync Engine
ทำหน้าที่ sync ข้อมูลระหว่าง SQLite local กับ Google Sheets

Architecture:
  - Local SQLite (truth ของเครื่องนั้น) ↔ Google Sheets (truth กลาง)
  - INV reservation pool: จองเลขล่วงหน้าเป็น batch (50 เลข) ใช้ตอนออฟไลน์ได้
  - Sync = push pending rows ขึ้น Sheets + pull rows ที่อัปเดตจาก Sheets
"""
import os
import json
import uuid
import time
import socket
import sqlite3
import threading
from datetime import datetime

# gspread/google-auth - import แบบ lazy (อาจไม่ติดตั้ง)
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_OK = True
except ImportError:
    GSPREAD_OK = False


# Sheet/header definitions ----------------------------------------------------

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

SHEET_INVOICES = 'Invoices'
SHEET_CUSTOMERS = 'Customers'
SHEET_COUNTER = 'Counter'
SHEET_SELLERS = 'Sellers'

INVOICE_COLS = [
    'invoice_no', 'invoice_date', 'seller_name',
    'cust_name', 'cust_address', 'cust_taxid',
    'subtotal', 'grand_total', 'amount_text',
    'pay_method', 'pay_bank', 'pay_account_no', 'pay_date', 'pay_amount',
    'items_json',
    'created_at', 'updated_at', 'device_id', 'deleted',
]

CUSTOMER_COLS = ['cust_name', 'cust_address', 'cust_taxid', 'last_used', 'updated_at', 'device_id']
COUNTER_COLS = ['month_id', 'device_id', 'count', 'timestamp', 'request_uuid']
SELLER_COLS = ['name', 'updated_at', 'device_id', 'deleted']


# =============================================================================
# Helper: online check
# =============================================================================

def is_online_ping(timeout=2.0):
    """ตรวจง่ายๆ ว่ามี internet ไหม (ping Google DNS)"""
    try:
        socket.setdefaulttimeout(timeout)
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.connect(('8.8.8.8', 53))
        s.close()
        return True
    except Exception:
        return False


# =============================================================================
# SyncEngine
# =============================================================================

class SyncEngine:
    def __init__(self, local_db_path, credentials_path, sheet_id, device_id):
        self.local_db_path = local_db_path
        self.credentials_path = credentials_path
        self.sheet_id = sheet_id
        self.device_id = device_id
        self._client = None
        self._sheet = None
        self._lock = threading.Lock()
        self._last_sync_at = None
        self._last_sync_error = None
        self._is_syncing = False
        self._stop_event = threading.Event()
        self._thread = None

    # -- configuration check ---------------------------------------------------

    def is_configured(self):
        return (
            GSPREAD_OK
            and self.credentials_path
            and os.path.exists(self.credentials_path)
            and bool(self.sheet_id)
        )

    def is_online(self):
        return is_online_ping()

    # -- gspread connection ----------------------------------------------------

    def _connect(self, force=False):
        if not self.is_configured():
            raise RuntimeError('ยังไม่ได้ตั้งค่า credentials.json หรือ Sheet ID')
        if self._sheet is not None and not force:
            return self._sheet
        creds = Credentials.from_service_account_file(self.credentials_path, scopes=SCOPES)
        self._client = gspread.authorize(creds)
        self._sheet = self._client.open_by_key(self.sheet_id)
        return self._sheet

    def test_connection(self):
        """Return (ok, message). Used by Settings UI."""
        if not GSPREAD_OK:
            return False, 'ไลบรารี gspread ยังไม่ติดตั้ง — รัน install.bat ใหม่'
        if not self.credentials_path or not os.path.exists(self.credentials_path):
            return False, 'ไม่พบไฟล์ credentials.json ในโฟลเดอร์โปรแกรม'
        if not self.sheet_id:
            return False, 'กรุณาใส่ Sheet ID ก่อน'
        if not self.is_online():
            return False, 'ไม่มี internet — ลองเช็คสัญญาณ'
        try:
            sheet = self._connect(force=True)
            self.setup_sheets()
            return True, f'เชื่อมต่อสำเร็จ — Sheet: "{sheet.title}"'
        except Exception as e:
            return False, f'เชื่อมต่อไม่ได้: {e}'

    # -- ensure sheets/headers exist -------------------------------------------

    def setup_sheets(self):
        """สร้าง tab/headers ใน Google Sheet ถ้ายังไม่มี"""
        sheet = self._connect()
        existing = {ws.title: ws for ws in sheet.worksheets()}

        plan = [
            (SHEET_INVOICES, INVOICE_COLS),
            (SHEET_CUSTOMERS, CUSTOMER_COLS),
            (SHEET_COUNTER, COUNTER_COLS),
            (SHEET_SELLERS, SELLER_COLS),
        ]

        for name, cols in plan:
            if name in existing:
                ws = existing[name]
                # Check headers
                first_row = ws.row_values(1)
                if first_row != cols:
                    # rewrite header
                    ws.update('A1', [cols])
            else:
                ws = sheet.add_worksheet(title=name, rows=1000, cols=len(cols))
                ws.update('A1', [cols])
                ws.freeze(rows=1)

        # Remove default Sheet1 if it exists and is empty
        if 'Sheet1' in existing:
            try:
                sheet.del_worksheet(existing['Sheet1'])
            except Exception:
                pass

    # -- INV reservation -------------------------------------------------------

    def reserve_inv_batch(self, count, prefix='INV'):
        """
        จองเลข INV เป็น batch (atomic) → คืนรายการเลขใบเสร็จเต็มๆ

        ใช้วิธี append-then-read:
        1. append แถวคำขอจองพร้อม UUID เฉพาะของเรา
        2. อ่านทั้งหมด → นับ count ที่มาก่อนเรา = offset
        3. คืนเลข [prefix + month + (offset+1) ... (offset+count)]

        ถ้า 2 เครื่อง append พร้อมกัน Google Sheets จะ serialize ให้
        — ทุกเครื่องเห็น row order เดียวกัน → ไม่มีทาง overlap
        """
        sheet = self._connect()
        try:
            ws = sheet.worksheet(SHEET_COUNTER)
        except gspread.WorksheetNotFound:
            self.setup_sheets()
            ws = sheet.worksheet(SHEET_COUNTER)

        month_id = datetime.now().strftime('%Y%m')
        request_uuid = str(uuid.uuid4())
        now_iso = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # Step 1: append my request
        ws.append_row(
            [month_id, self.device_id, count, now_iso, request_uuid],
            value_input_option='RAW',
        )

        # Step 2: read all rows
        all_values = ws.get_all_values()
        if not all_values:
            raise RuntimeError('Counter sheet ว่าง')
        data_rows = all_values[1:]  # skip header

        # Find my row, accumulate offset
        offset = 0
        found = False
        for row in data_rows:
            if len(row) < 5:
                continue
            row_month, row_device, row_count, row_ts, row_uuid = row[:5]
            if row_month != month_id:
                continue
            if row_uuid == request_uuid:
                found = True
                break
            try:
                offset += int(row_count)
            except (ValueError, TypeError):
                continue

        if not found:
            raise RuntimeError('ไม่พบแถวคำขอจองหลังจาก append — โปรดลองใหม่')

        # Step 3: build INV strings
        invs = []
        for i in range(count):
            seq = offset + i + 1
            invs.append(f"{prefix}{month_id}{seq:04d}")
        return invs

    # -- push: upload pending rows to Sheets -----------------------------------

    def _open_local_db(self):
        conn = sqlite3.connect(self.local_db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def push_pending(self):
        """อัปโหลด rows ที่ยัง sync_status='pending' ขึ้น Sheets"""
        sheet = self._connect()
        ws_inv = sheet.worksheet(SHEET_INVOICES)
        ws_cust = sheet.worksheet(SHEET_CUSTOMERS)
        ws_seller = sheet.worksheet(SHEET_SELLERS)

        # Pre-fetch existing invoice_nos in cloud (for upsert)
        cloud_invoice_rows = ws_inv.get_all_values()
        cloud_invoice_index = {}  # invoice_no -> row index (1-based)
        for idx, row in enumerate(cloud_invoice_rows[1:], start=2):
            if row and len(row) > 0 and row[0]:
                cloud_invoice_index[row[0]] = idx

        cloud_cust_rows = ws_cust.get_all_values()
        cloud_cust_index = {}
        for idx, row in enumerate(cloud_cust_rows[1:], start=2):
            if row and len(row) > 0 and row[0]:
                cloud_cust_index[row[0]] = idx

        cloud_seller_rows = ws_seller.get_all_values()
        cloud_seller_index = {}
        for idx, row in enumerate(cloud_seller_rows[1:], start=2):
            if row and len(row) > 0 and row[0]:
                cloud_seller_index[row[0]] = idx

        conn = self._open_local_db()
        cur = conn.cursor()
        pushed_invoices = 0
        pushed_customers = 0
        pushed_sellers = 0

        # Push invoices
        cur.execute("SELECT * FROM invoices WHERE sync_status = 'pending'")
        rows = cur.fetchall()
        for r in rows:
            # Get items
            cur2 = conn.cursor()
            cur2.execute("SELECT item_no, description, quantity, unit_price, amount FROM invoice_items WHERE invoice_no = ? ORDER BY item_no", (r['invoice_no'],))
            items = [dict(it) for it in cur2.fetchall()]
            items_json = json.dumps(items, ensure_ascii=False)

            row_values = [
                r['invoice_no'], r['invoice_date'] or '', r['seller_name'] or '',
                r['cust_name'] or '', r['cust_address'] or '', r['cust_taxid'] or '',
                r['subtotal'] or 0, r['grand_total'] or 0, r['amount_text'] or '',
                r['pay_method'] or '', r['pay_bank'] or '', r['pay_account_no'] or '',
                r['pay_date'] or '', r['pay_amount'] or 0,
                items_json,
                r['created_at'] or '', r['updated_at'] or '',
                self.device_id, str(r['deleted'] or 0),
            ]
            # Convert all to strings for stable Sheets upload
            row_values = [str(v) if v is not None else '' for v in row_values]

            invno = r['invoice_no']
            if invno in cloud_invoice_index:
                # Update existing row
                row_idx = cloud_invoice_index[invno]
                cell_range = f'A{row_idx}:S{row_idx}'
                ws_inv.update(cell_range, [row_values], value_input_option='USER_ENTERED')
            else:
                ws_inv.append_row(row_values, value_input_option='USER_ENTERED')
                cloud_invoice_index[invno] = len(cloud_invoice_rows) + 1
                cloud_invoice_rows.append(row_values)

            cur2.execute("UPDATE invoices SET sync_status = 'synced' WHERE invoice_no = ?", (invno,))
            conn.commit()
            pushed_invoices += 1

        # Push customers
        cur.execute("SELECT * FROM customers WHERE sync_status = 'pending'")
        for r in cur.fetchall():
            row_values = [
                r['cust_name'] or '', r['cust_address'] or '', r['cust_taxid'] or '',
                r['last_used'] or '', r['updated_at'] or '', self.device_id,
            ]
            row_values = [str(v) if v is not None else '' for v in row_values]
            name = r['cust_name']
            if name in cloud_cust_index:
                row_idx = cloud_cust_index[name]
                cell_range = f'A{row_idx}:F{row_idx}'
                ws_cust.update(cell_range, [row_values], value_input_option='USER_ENTERED')
            else:
                ws_cust.append_row(row_values, value_input_option='USER_ENTERED')
                cloud_cust_index[name] = len(cloud_cust_rows) + 1
                cloud_cust_rows.append(row_values)
            cur2 = conn.cursor()
            cur2.execute("UPDATE customers SET sync_status = 'synced' WHERE cust_name = ?", (name,))
            conn.commit()
            pushed_customers += 1

        # Push sellers
        cur.execute("SELECT * FROM sellers WHERE sync_status = 'pending'")
        for r in cur.fetchall():
            row_values = [
                r['name'] or '', r['updated_at'] or '', self.device_id, str(r['deleted'] or 0),
            ]
            row_values = [str(v) if v is not None else '' for v in row_values]
            name = r['name']
            if name in cloud_seller_index:
                row_idx = cloud_seller_index[name]
                cell_range = f'A{row_idx}:D{row_idx}'
                ws_seller.update(cell_range, [row_values], value_input_option='USER_ENTERED')
            else:
                ws_seller.append_row(row_values, value_input_option='USER_ENTERED')
                cloud_seller_index[name] = len(cloud_seller_rows) + 1
                cloud_seller_rows.append(row_values)
            cur2 = conn.cursor()
            cur2.execute("UPDATE sellers SET sync_status = 'synced' WHERE name = ?", (name,))
            conn.commit()
            pushed_sellers += 1

        conn.close()
        return {
            'invoices': pushed_invoices,
            'customers': pushed_customers,
            'sellers': pushed_sellers,
        }

    # -- pull: download cloud rows into local SQLite ---------------------------

    def pull_changes(self):
        """ดึง rows จาก Sheets → merge เข้า SQLite (last-write-wins)"""
        sheet = self._connect()
        ws_inv = sheet.worksheet(SHEET_INVOICES)
        ws_cust = sheet.worksheet(SHEET_CUSTOMERS)
        ws_seller = sheet.worksheet(SHEET_SELLERS)

        conn = self._open_local_db()
        cur = conn.cursor()
        pulled_invoices = 0
        pulled_customers = 0
        pulled_sellers = 0

        # Pull invoices
        all_rows = ws_inv.get_all_values()
        if all_rows:
            for row in all_rows[1:]:
                if len(row) < len(INVOICE_COLS):
                    row = row + [''] * (len(INVOICE_COLS) - len(row))
                d = dict(zip(INVOICE_COLS, row))
                if not d.get('invoice_no'):
                    continue

                # Check local version
                cur.execute("SELECT updated_at, device_id, sync_status FROM invoices WHERE invoice_no = ?",
                            (d['invoice_no'],))
                existing = cur.fetchone()

                cloud_updated = d.get('updated_at', '') or ''
                if existing:
                    # Skip pending local rows (don't overwrite user's unsynced changes)
                    if existing['sync_status'] == 'pending':
                        continue
                    local_updated = existing['updated_at'] or ''
                    if cloud_updated <= local_updated:
                        continue
                    # cloud is newer → update local

                # Parse items_json
                items = []
                try:
                    items = json.loads(d.get('items_json', '') or '[]')
                except Exception:
                    items = []

                deleted_flag = 1 if str(d.get('deleted', '0')) in ('1', 'True', 'true') else 0

                if existing:
                    cur.execute("""
                        UPDATE invoices SET
                          invoice_date=?, seller_name=?, cust_name=?, cust_address=?, cust_taxid=?,
                          subtotal=?, grand_total=?, amount_text=?, pay_method=?, pay_bank=?,
                          pay_account_no=?, pay_date=?, pay_amount=?, created_at=?, updated_at=?,
                          sync_status='synced', deleted=?
                        WHERE invoice_no=?
                    """, (
                        d.get('invoice_date'), d.get('seller_name'),
                        d.get('cust_name'), d.get('cust_address'), d.get('cust_taxid'),
                        _to_float(d.get('subtotal')), _to_float(d.get('grand_total')),
                        d.get('amount_text'), d.get('pay_method'), d.get('pay_bank'),
                        d.get('pay_account_no'), d.get('pay_date'), _to_float(d.get('pay_amount')),
                        d.get('created_at'), d.get('updated_at'), deleted_flag,
                        d['invoice_no'],
                    ))
                else:
                    cur.execute("""
                        INSERT INTO invoices (
                          invoice_no, invoice_date, seller_name, cust_name, cust_address, cust_taxid,
                          subtotal, grand_total, amount_text, pay_method, pay_bank, pay_account_no,
                          pay_date, pay_amount, created_at, updated_at, sync_status, deleted
                        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?, 'synced', ?)
                    """, (
                        d['invoice_no'], d.get('invoice_date'), d.get('seller_name'),
                        d.get('cust_name'), d.get('cust_address'), d.get('cust_taxid'),
                        _to_float(d.get('subtotal')), _to_float(d.get('grand_total')),
                        d.get('amount_text'), d.get('pay_method'), d.get('pay_bank'),
                        d.get('pay_account_no'), d.get('pay_date'), _to_float(d.get('pay_amount')),
                        d.get('created_at'), d.get('updated_at'), deleted_flag,
                    ))

                # Replace items
                cur.execute("DELETE FROM invoice_items WHERE invoice_no = ?", (d['invoice_no'],))
                for it in items:
                    cur.execute("""
                        INSERT INTO invoice_items (invoice_no, item_no, description, quantity, unit_price, amount)
                        VALUES (?,?,?,?,?,?)
                    """, (
                        d['invoice_no'], int(it.get('item_no') or 0),
                        it.get('description', ''),
                        _to_float(it.get('quantity')),
                        _to_float(it.get('unit_price')),
                        _to_float(it.get('amount')),
                    ))
                pulled_invoices += 1

        # Pull customers
        all_rows = ws_cust.get_all_values()
        if all_rows:
            for row in all_rows[1:]:
                if len(row) < len(CUSTOMER_COLS):
                    row = row + [''] * (len(CUSTOMER_COLS) - len(row))
                d = dict(zip(CUSTOMER_COLS, row))
                if not d.get('cust_name'):
                    continue
                cur.execute("SELECT updated_at, sync_status FROM customers WHERE cust_name = ?", (d['cust_name'],))
                existing = cur.fetchone()
                if existing:
                    if existing['sync_status'] == 'pending':
                        continue
                    if (d.get('updated_at', '') or '') <= (existing['updated_at'] or ''):
                        continue
                    cur.execute("""
                        UPDATE customers SET cust_address=?, cust_taxid=?, last_used=?, updated_at=?, sync_status='synced'
                        WHERE cust_name=?
                    """, (d.get('cust_address'), d.get('cust_taxid'), d.get('last_used'),
                          d.get('updated_at'), d['cust_name']))
                else:
                    cur.execute("""
                        INSERT INTO customers (cust_name, cust_address, cust_taxid, last_used, updated_at, sync_status)
                        VALUES (?,?,?,?,?, 'synced')
                    """, (d['cust_name'], d.get('cust_address'), d.get('cust_taxid'),
                          d.get('last_used'), d.get('updated_at')))
                pulled_customers += 1

        # Pull sellers
        all_rows = ws_seller.get_all_values()
        if all_rows:
            for row in all_rows[1:]:
                if len(row) < len(SELLER_COLS):
                    row = row + [''] * (len(SELLER_COLS) - len(row))
                d = dict(zip(SELLER_COLS, row))
                if not d.get('name'):
                    continue
                cur.execute("SELECT updated_at, sync_status FROM sellers WHERE name = ?", (d['name'],))
                existing = cur.fetchone()
                deleted_flag = 1 if str(d.get('deleted', '0')) in ('1', 'True', 'true') else 0
                if existing:
                    if existing['sync_status'] == 'pending':
                        continue
                    if (d.get('updated_at', '') or '') <= (existing['updated_at'] or ''):
                        continue
                    cur.execute("UPDATE sellers SET updated_at=?, sync_status='synced', deleted=? WHERE name=?",
                                (d.get('updated_at'), deleted_flag, d['name']))
                else:
                    cur.execute("INSERT INTO sellers (name, updated_at, sync_status, deleted) VALUES (?,?, 'synced', ?)",
                                (d['name'], d.get('updated_at'), deleted_flag))
                pulled_sellers += 1

        conn.commit()
        conn.close()
        return {
            'invoices': pulled_invoices,
            'customers': pulled_customers,
            'sellers': pulled_sellers,
        }

    # -- combined sync ---------------------------------------------------------

    def sync_now(self):
        """Push pending → Pull changes. Returns {'ok': bool, 'pushed': {...}, 'pulled': {...}, 'error': ...}"""
        if self._is_syncing:
            return {'ok': False, 'error': 'กำลัง sync อยู่ — รอสักครู่'}
        if not self.is_configured():
            return {'ok': False, 'error': 'ยังไม่ได้ตั้งค่า cloud sync'}
        if not self.is_online():
            return {'ok': False, 'error': 'ออฟไลน์ — รอจน internet กลับมาก่อน'}

        with self._lock:
            self._is_syncing = True
            self._last_sync_error = None
            try:
                self._connect()
                pushed = self.push_pending()
                pulled = self.pull_changes()
                self._last_sync_at = datetime.now()
                return {'ok': True, 'pushed': pushed, 'pulled': pulled}
            except Exception as e:
                self._last_sync_error = str(e)
                return {'ok': False, 'error': str(e)}
            finally:
                self._is_syncing = False

    # -- background sync loop --------------------------------------------------

    def start_background(self, interval_sec=30):
        if self._thread and self._thread.is_alive():
            return
        self._stop_event.clear()
        self._thread = threading.Thread(
            target=self._background_loop, args=(interval_sec,), daemon=True
        )
        self._thread.start()

    def stop_background(self):
        self._stop_event.set()

    def _background_loop(self, interval_sec):
        while not self._stop_event.is_set():
            try:
                if self.is_configured() and self.is_online():
                    self.sync_now()
            except Exception as e:
                self._last_sync_error = str(e)
            # sleep
            self._stop_event.wait(interval_sec)

    # -- status ---------------------------------------------------------------

    def status(self):
        configured = self.is_configured()
        online = self.is_online() if configured else False
        last = self._last_sync_at.strftime('%Y-%m-%d %H:%M:%S') if self._last_sync_at else None
        # Count pending
        pending_count = 0
        try:
            conn = self._open_local_db()
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*) FROM invoices WHERE sync_status='pending'")
            pending_count += cur.fetchone()[0]
            cur.execute("SELECT COUNT(*) FROM customers WHERE sync_status='pending'")
            pending_count += cur.fetchone()[0]
            cur.execute("SELECT COUNT(*) FROM sellers WHERE sync_status='pending'")
            pending_count += cur.fetchone()[0]
            conn.close()
        except Exception:
            pass
        return {
            'configured': configured,
            'online': online,
            'is_syncing': self._is_syncing,
            'last_sync_at': last,
            'last_sync_error': self._last_sync_error,
            'pending_count': pending_count,
            'device_id': self.device_id,
        }


# -----------------------------------------------------------------------------

def _to_float(v):
    try:
        return float(v) if v not in (None, '') else 0.0
    except (ValueError, TypeError):
        return 0.0
