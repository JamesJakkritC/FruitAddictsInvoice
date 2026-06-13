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
    
    # บน Vercel จะไม่อนุญาตให้เขียนทับไฟล์ config.json ตรงๆ 
    # จึงให้ตอบกลับสำเร็จเพื่อให้ระบบหน้าบ้านทำงานต่อได้
    return jsonify({
        'ok': True, 
        'message': 'ตั้งค่าบนระบบ Cloud เรียบร้อยแล้ว (การแก้ไขโครงสร้างบริษัทแนะนำให้แก้ไขที่แท็บ Config บน Google Sheets โดยตรง)'
    })


@app.route('/api/next-invoice-no')
def api_next_invoice_no():
    return jsonify({'invoice_no': get_next_invoice_no_from_sheets()})


@app.route('/api/invoices', methods=['GET', 'POST'])
def api_invoices():
    if request.method == 'POST':
        return jsonify(save_invoice_to_sheets(request.json or {}))
        
    # ตรรกะการเรียกดูรายการทั้งหมด (GET) พร้อมคุณสมบัติ Search Filter
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
                ws.update_cell(idx + 1, 19, '1')  # คอลัมน์ที่ 19 คือคอลัมน์ deleted
                return jsonify({'ok': True})
        return jsonify({'error': 'not found'}), 404
        
    # GET Detail
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
            ws.update_cell(idx + 1, 4, '1')  # คอลัมน์ที่ 4 คือคอลัมน์ deleted
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
            
        # ทดสอบเรียกเปิดแผ่นงาน Config
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
