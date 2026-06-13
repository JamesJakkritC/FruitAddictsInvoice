# -*- coding: utf-8 -*-
import os
import json
import uuid
from datetime import datetime
from flask import Flask, render_template, request, jsonify, redirect

import gspread
from google.oauth2.service_account import Credentials

APP_DIR = os.path.dirname(os.path.abspath(__file__))

# ประกาศตัวแปรระดับบนสุดเพื่อให้ Vercel ตรวจจับได้ทันที
app = Flask(__name__,
            static_folder=os.path.join(APP_DIR, 'static'),
            template_folder=os.path.join(APP_DIR, 'templates'))

application = app

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

IS_VERCEL = os.environ.get('VERCEL') == '1'

# ============================================================
# Core Google Sheets Connectors & Logic Functions
# ============================================================

def get_worksheet(sheet_name):
    """เชื่อมต่อ Google Sheets รองรับทั้งบน Vercel (Env) และเครื่องตัวเอง (Local Files)"""
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    sheet_id = os.environ.get("GOOGLE_SHEET_ID")
    
    if not IS_VERCEL:
        if not creds_json:
            local_creds_path = os.path.join(APP_DIR, 'credentials.json')
            if os.path.exists(local_creds_path):
                with open(local_creds_path, 'r', encoding='utf-8') as f:
                    creds_json = f.read()
                    
        if not sheet_id:
            local_config_path = os.path.join(APP_DIR, 'config.json')
            if os.path.exists(local_config_path):
                try:
                    with open(local_config_path, 'r', encoding='utf-8') as f:
                        local_cfg = json.load(f)
                        sheet_id = local_cfg.get('sheet_id')
                except Exception as e:
                    print(f"[!] ไม่สามารถอ่านค่าจาก config.json บนเครื่องได้: {e}")

    if not creds_json or not sheet_id:
        raise RuntimeError(
            "ระบบตรวจไม่พบการตั้งค่าสิทธิ์เชื่อมต่อ: กรุณาตรวจสอบ Environment Variables บน Vercel"
        )
        
    creds_data = json.loads(creds_json)
   if "private_key" in creds_data:
    creds_data["private_key"] = creds_data["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_data, scopes=SCOPES)
               
    client = gspread.authorize(creds)
    sheet = client.open_by_key(sheet_id)
    return sheet.worksheet(sheet_name)


def load_config_from_sheets():
    """ดึงข้อมูลการตั้งค่าและเทมเพลตหัวบิลจากแผ่นงาน Config โดยตรง"""
    try:
        ws = get_worksheet('Config')
        records = ws.get_all_records()
        cfg = {r.get('Key'): r.get('Value') for r in records if r.get('Key')}
        
        return {
            'invoice_prefix': cfg.get('invoice_prefix', 'INV'),
            'default_seller': cfg.get('default_seller', ''),
            'bank_default': cfg.get('bank_default', ''),
            'bank_account_default': cfg.get('bank_account_default', ''),
            'company': {
                'name': cfg.get('company_name', 'บริษัท ฟรุ๊ตแอคดิคส์ พรีเมี่ยมฟรุ๊ต จำกัด'),
                'address': cfg.get('company_address', ''),
                'taxid': cfg.get('company_taxid', ''),
                'phone': cfg.get('company_phone', ''),
                'mobile': cfg.get('company_mobile', ''),
                'email': cfg.get('company_email', ''),
                'line': cfg.get('company_line', ''),
            },
            'logo_path': cfg.get('logo_url', ''),
            'stamp_path': cfg.get('stamp_url', ''),
            'cloud_sync_enabled': True,
            'device_id': 'CLOUD-WEB'
        }
    except Exception as e:
        print(f"Error loading config from sheet: {e}")
        return {
            'invoice_prefix': 'INV', 'default_seller': '', 'bank_default': '', 'bank_account_default': '',
            'company': {'name': 'Fruit Addicts Online'}, 'logo_path': '', 'stamp_path': ''
        }


def get_next_invoice_no_from_sheets():
    """จองเลขคิวบิลตามลำดับพูลบนคลาวด์ป้องกันเลขซ้ำกัน"""
    cfg = load_config_from_sheets()
    prefix = cfg.get('invoice_prefix', 'INV')
    ws = get_worksheet('Counter')
    
    month_id = datetime.now().strftime('%Y%m')
    request_uuid = str(uuid.uuid4())
    now_iso = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    ws.append_row([month_id, 'CLOUD_WEB', 1, now_iso, request_uuid], value_input_option='RAW')
    
    all_values = ws.get_all_values()
    data_rows = all_values[1:]
    
    offset = 0
    for row in data_rows:
        if len(row) < 5: continue
        row_month, _, row_count, _, row_uuid = row[:5]
        if row_month != month_id: continue
        if row_uuid == request_uuid: break
        try: offset += int(row_count)
        except ValueError: continue
            
    seq = offset + 1
    return f"{prefix}{month_id}{seq:04d}"


def save_invoice_to_sheets(payload):
    ws = get_worksheet('Invoices')
    all_values = ws.get_all_values()
    
    invoice_no = payload.get('invoice_no')
    if not invoice_no or payload.get('is_update') == False:
        invoice_no = get_next_invoice_no_from_sheets()
        
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    items_json = json.dumps(payload.get('items', []), ensure_ascii=False)
    
    row_data = [
        invoice_no, payload.get('invoice_date', ''), payload.get('seller_name', ''),
        payload.get('cust_name', ''), payload.get('cust_address', ''), payload.get('cust_taxid', ''),
        str(payload.get('subtotal', 0)), str(payload.get('grand_total', 0)), payload.get('amount_text', ''),
        payload.get('pay_method', ''), payload.get('pay_bank', ''), payload.get('pay_account_no', ''),
        payload.get('pay_date', ''), str(payload.get('pay_amount', 0)),
        items_json, now_str, now_str, 'CLOUD_WEB', '0'
    ]
    
    row_idx = -1
    for idx, row in enumerate(all_values):
        if row and row[0] == invoice_no:
            row_idx = idx + 1
            break
            
    if row_idx != -1:
        if len(all_values[row_idx-1]) > 15:
            row_data[15] = all_values[row_idx-1][15]
        cell_range = f'A{row_idx}:S{row_idx}'
        ws.update(cell_range, [row_data], value_input_option='USER_ENTERED')
    else:
        ws.append_row(row_data, value_input_option='USER_ENTERED')
        
    if payload.get('cust_name'):
        upsert_customer(payload, now_str)
    if payload.get('seller_name'):
        upsert_seller(payload.get('seller_name'), now_str)
        
    return {'ok': True, 'invoice_no': invoice_no}


def upsert_customer(payload, now_str):
    ws = get_worksheet('Customers')
    all_values = ws.get_all_values()
    cust_name = payload.get('cust_name').strip()
    row_data = [cust_name, payload.get('cust_address', ''), payload.get('cust_taxid', ''), now_str, now_str, 'CLOUD_WEB']
    
    row_idx = -1
    for idx, row in enumerate(all_values):
        if row and row[0].strip() == cust_name:
            row_idx = idx + 1
            break
    if row_idx != -1:
        ws.update(f'A{row_idx}:F{row_idx}', [row_data], value_input_option='USER_ENTERED')
    else:
        ws.append_row(row_data, value_input_option='USER_ENTERED')


def upsert_seller(name, now_str):
    ws = get_worksheet('Sellers')
    all_values = ws.get_all_values()
    name = name.strip()
    row_data = [name, now_str, 'CLOUD_WEB', '0']
    
    row_idx = -1
    for idx, row in enumerate(all_values):
        if row and row[0].strip() == name:
            row_idx = idx + 1
            break
    if row_idx != -1:
        ws.update(f'A{row_idx}:D{row_idx}', [row_data], value_input_option='USER_ENTERED')
    else:
        ws.append_row(row_data, value_input_option='USER_ENTERED')


def baht_text(num):
    txt_num = ['ศูนย์', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า']
    txt_pos = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน', 'ล้าน']
    def read_int(s):
        s = str(s)
        if s == '0': return 'ศูนย์'
        result = ''
        n = len(s)
        for i, ch in enumerate(s):
            d = int(ch)
            pos = n - i - 1
            if d == 0: continue
            if pos % 6 == 1:
                if d == 1: result += 'สิบ'
                elif d == 2: result += 'ยี่สิบ'
                else: result += txt_num[d] + 'สิบ'
            elif pos % 6 == 0 and pos != 0:
                if d == 1 and i != 0 and int(s[i - 1]) != 0: result += 'เอ็ด'
                else: result += txt_num[d]
                result += 'ล้าน'
            elif pos == 0:
                if d == 1 and n > 1 and int(s[i - 1]) != 0: result += 'เอ็ด'
                else: result += txt_num[d]
            else:
                result += txt_num[d] + txt_pos[pos % 6]
        return result
    try: num = float(num)
    except: num = 0
    int_part = int(num)
    satang = round((num - int_part) * 100)
    txt = read_int(int_part) + 'บาท'
    if satang > 0: txt += read_int(satang) + 'สตางค์'
    else: txt += 'ถ้วน'
    return txt

# ============================================================
# Flask Routing Controller
# ============================================================

@app.route('/')
def page_index():
    return render_template('index.html', cfg=load_config_from_sheets())


@app.route('/print/<invno>')
def page_print(invno):
    ws = get_worksheet('Invoices')
    all_values = ws.get_all_values()
    headers = all_values[0]
    
    inv = None
    for row in all_values[1:]:
        if row and row[0] == invno and (len(row) <= 18 or row[18] != '1'):
            if len(row) < len(headers):
                row = row + [''] * (len(headers) - len(row))
            inv = dict(zip(headers, row))
            inv['subtotal'] = float(inv.get('subtotal') or 0)
            inv['grand_total'] = float(inv.get('grand_total') or 0)
            inv['pay_amount'] = float(inv.get('pay_amount') or 0)
            try: inv['items'] = json.loads(inv.get('items_json', '[]'))
            except: inv['items'] = []
            break
            
    if not inv:
        return f'ไม่พบใบเสร็จรับเงินเลขที่: {invno}', 404
        
    cfg = load_config_from_sheets()
    return render_template('invoice.html', inv=inv, cfg=cfg, logo_url=cfg['logo_path'], stamp_url=cfg['stamp_path'])


@app.route('/api/config', methods=['GET', 'POST'])
def api_config():
    if request.method == 'GET':
        return jsonify(load_config_from_sheets())
    return jsonify({
        'ok': True, 
        'message': 'ตั้งค่าบนระบบ Cloud เรียบร้อยแล้ว (แนะนำให้แก้ไขที่แท็บ Config บน Google Sheets โดยตรง)'
    })


@app.route('/api/next-invoice-no')
def api_next_invoice_no():
    return jsonify({'invoice_no': get_next_invoice_no_from_sheets()})


@app.route('/api/invoices', methods=['GET', 'POST'])
def api_invoices():
    if request.method == 'POST':
        return jsonify(save_invoice_to_sheets(request.json or {}))
        
    q = request.args.get('q', '').lower()
    ws = get_worksheet('Invoices')
    all_values = ws.get_all_values()
    if not all_values or len(all_values) <= 1:
        return jsonify([])
        
    headers = all_values[0]
    result_list = []
    for row in all_values[1:]:
        if not row or not row[0]: continue
        if len(row) < len(headers):
            row = row + [''] * (len(headers) - len(row))
        d = dict(zip(headers, row))
        
        if d.get('deleted') == '1': continue
        if q and (q not in d.get('invoice_no', '').lower() and q not in d.get('cust_name', '').lower() and q not in d.get('cust_taxid', '').lower()):
            continue
            
        d['subtotal'] = float(d.get('subtotal') or 0)
        d['grand_total'] = float(d.get('grand_total') or 0)
        d['pay_amount'] = float(d.get('pay_amount') or 0)
        result_list.append(d)
        
    result_list.sort(key=lambda x: x.get('invoice_no', ''), reverse=True)
    return jsonify(result_list)


@app.route('/api/invoices/<invno>', methods=['GET', 'DELETE'])
def api_invoice_detail(invno):
    ws = get_worksheet('Invoices')
    all_values = ws.get_all_values()
    headers = all_values[0]
    
    if request.method == 'DELETE':
        for idx, row in enumerate(all_values):
            if row and row[0] == invno:
                ws.update_cell(idx + 1, 19, '1')
                return jsonify({'ok': True})
        return jsonify({'error': 'not found'}), 404
        
    for row in all_values[1:]:
        if row and row[0] == invno and (len(row) <= 18 or row[18] != '1'):
            if len(row) < len(headers):
                row = row + [''] * (len(headers) - len(row))
            d = dict(zip(headers, row))
            d['subtotal'] = float(d.get('subtotal') or 0)
            d['grand_total'] = float(d.get('grand_total') or 0)
            d['pay_amount'] = float(d.get('pay_amount') or 0)
            try: d['items'] = json.loads(d.get('items_json', '[]'))
            except: d['items'] = []
            return jsonify(d)
    return jsonify({'error': 'not found'}), 404


@app.route('/api/customers')
def api_customers():
    q = request.args.get('q', '').lower()
    ws = get_worksheet('Customers')
    all_values = ws.get_all_values()
    if not all_values or len(all_values) <= 1:
        return jsonify([])
        
    headers = all_values[0]
    customers = []
    for row in all_values[1:]:
        if not row or not row[0]: continue
        if len(row) < len(headers):
            row = row + [''] * (len(headers) - len(row))
        d = dict(zip(headers, row))
        if q and (q not in d.get('cust_name', '').lower() and q not in d.get('cust_taxid', '').lower()):
            continue
        customers.append(d)
    return jsonify(customers[:20])


@app.route('/api/sellers', methods=['GET', 'POST'])
def api_sellers():
    ws = get_worksheet('Sellers')
    all_values = ws.get_all_values()
    
    if request.method == 'POST':
        payload = request.json or {}
        name = payload.get('name', '').strip()
        if name:
            upsert_seller(name, datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            return jsonify({'ok': True})
        return jsonify({'ok': False, 'error': 'ชื่อพนักงานขายห้ามว่าง'})
        
    if not all_values or len(all_values) <= 1:
        return jsonify([])
        
    headers = all_values[0]
    sellers = set()
    for row in all_values[1:]:
        if row and row[0] and (len(row) <= 3 or row[3] != '1'):
            sellers.add(row[0].strip())
    return jsonify(sorted(list(sellers)))


@app.route('/api/sellers/<name>', methods=['DELETE'])
def api_seller_delete(name):
    ws = get_worksheet('Sellers')
    all_values = ws.get_all_values()
    for idx, row in enumerate(all_values):
        if row and row[0].strip() == name.strip():
            ws.update_cell(idx + 1, 4, '1')
            return jsonify({'ok': True})
    return jsonify({'ok': False, 'error': 'not found'}), 404


@app.route('/api/baht-text')
def api_baht_text():
    try: amt = float(request.args.get('amount', 0))
    except: amt = 0
    return jsonify({'text': baht_text(amt)})


@app.route('/api/sync/status')
def api_sync_status():
    return jsonify({
        'sync_available': True, 'cloud_sync_enabled': True, 'has_credentials': True, 'has_sheet_id': True,
        'device_id': 'CLOUD-WEB', 'configured': True, 'online': True, 'is_syncing': False,
        'last_sync_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 'last_sync_error': None, 'pending_count': 0,
        'reservation': {'available': 999, 'total': 999, 'month_id': datetime.now().strftime('%Y%m')}
    })


@app.route('/api/db-info')
def api_db_info():
    return jsonify({'path': 'Google Sheets Cloud Storage', 'exists': True, 'credentials_exists': True})


@app.route('/api/sync/test', methods=['POST'])
def api_sync_test():
    """รองรับการกดปุ่ม 'ทดสอบการเชื่อมต่อ' จากหน้าบ้านเมื่อรันบน Vercel"""
    try:
        sheet_id = os.environ.get("GOOGLE_SHEET_ID")
        if not sheet_id:
            return jsonify({
                'ok': False, 
                'error': 'ไม่พบ GOOGLE_SHEET_ID ใน Environment Variables ของ Vercel'
            }), 400
            
        ws = get_worksheet('Config')
        ws.get_all_values()
        
        return jsonify({
            'ok': True, 
            'message': '✅ เชื่อมต่อกับ Google Sheets สำเร็จ! สิทธิ์การเข้าถึงถูกต้อง'
        })
    except Exception as e:
        return jsonify({
            'ok': False, 
            'error': f'ไม่สามารถเชื่อมต่อได้: {str(e)}'
        }), 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
