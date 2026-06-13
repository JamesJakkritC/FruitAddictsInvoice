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
        
    now_str = datetime.now().strftime('%Y-%m
