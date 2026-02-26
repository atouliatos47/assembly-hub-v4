import re
import os
import json
import uuid
import threading
import socket
from datetime import datetime
from flask import Flask, request, jsonify, send_from_directory
from flask_sock import Sock
from openpyxl import load_workbook
from docx import Document

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__, static_folder=os.path.join(BASE_DIR, 'public'), static_url_path='/static')
sock = Sock(app)

UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
os.makedirs(UPLOAD_DIR, exist_ok=True)

# ─── In-memory state ───────────────────────────────────────────────────────
centers   = {}    # { id: { id, name, color, assignedDoc, currentPage, connected } }
documents = {}    # { id: { id, name, type, pages, uploadedAt, filePath } }
clients   = {}    # { id: { ws, role, centerId } }

# ─── Helpers ───────────────────────────────────────────────────────────────
def broadcast(msg):
    dead = []
    for cid, c in clients.items():
        try:
            c['ws'].send(json.dumps(msg))
        except Exception:
            dead.append(cid)
    for cid in dead:
        clients.pop(cid, None)

def notify_center(center_id, msg):
    dead = []
    for cid, c in clients.items():
        if c.get('centerId') == center_id:
            try:
                c['ws'].send(json.dumps(msg))
            except Exception:
                dead.append(cid)
    for cid in dead:
        clients.pop(cid, None)

def doc_summary(doc):
    return {
        'id': doc['id'],
        'name': doc['name'],
        'type': doc['type'],
        'pageCount': len(doc.get('pages', [])),
        'uploadedAt': doc['uploadedAt'],
        'status': doc.get('status', 'ready')
    }

def make_slug(name):
    """Convert 'Assembly 1' → 'assembly-1', ensure uniqueness"""
    slug = name.lower().strip()
    slug = re.sub(r'[^a-z0-9]+', '-', slug).strip('-')
    base = slug
    counter = 2
    while any(c['slug'] == slug for c in centers.values()):
        slug = f"{base}-{counter}"
        counter += 1
    return slug

# ─── File Parsing ──────────────────────────────────────────────────────────
def kill_office_processes():
    """Kill any hanging Excel/Word processes"""
    import subprocess
    for proc in ['EXCEL.EXE', 'WINWORD.EXE']:
        subprocess.run(['taskkill', '/f', '/im', proc], capture_output=True)

def convert_to_pdf(filepath):
    """Convert Excel or Word to PDF using the installed Office application"""
    import subprocess, shutil
    ext = os.path.splitext(filepath)[1].lower()
    pdf_path = os.path.splitext(filepath)[0] + '.pdf'

    # Remove any existing PDF first
    if os.path.exists(pdf_path):
        os.remove(pdf_path)

    try:
        import win32com.client
        import pythoncom
        abs_path = os.path.abspath(filepath)
        abs_pdf  = os.path.abspath(pdf_path)

        # Initialize COM for this thread
        pythoncom.CoInitialize()
        app = None
        try:
            if ext in ['.xlsx', '.xls']:
                app = win32com.client.Dispatch('Excel.Application')
                app.Visible = False
                app.DisplayAlerts = False
                app.AskToUpdateLinks = False
                wb = app.Workbooks.Open(abs_path, UpdateLinks=0, ReadOnly=True)
                wb.ExportAsFixedFormat(0, abs_pdf)
                wb.Close(False)
            else:
                app = win32com.client.Dispatch('Word.Application')
                app.Visible = False
                app.DisplayAlerts = 0
                doc = app.Documents.Open(abs_path, ReadOnly=True)
                doc.SaveAs(abs_pdf, FileFormat=17)
                doc.Close()
        finally:
            if app:
                try: app.Quit()
                except: pass
            pythoncom.CoUninitialize()

        if os.path.exists(pdf_path):
            return pdf_path
        raise Exception('PDF not created')

    except ImportError:
        pass  # win32com not available

    # Fallback: LibreOffice
    lo = shutil.which('libreoffice') or shutil.which('soffice')
    if lo:
        out_dir = os.path.dirname(filepath)
        subprocess.run([lo, '--headless', '--convert-to', 'pdf', '--outdir', out_dir, filepath],
                      capture_output=True, timeout=60)
        if os.path.exists(pdf_path):
            return pdf_path

    raise Exception('No PDF converter available.')

def parse_file_to_pdf(filepath, original_name, doc_id):
    """Convert to PDF and return iframe pointing to served PDF file"""
    pdf_path = convert_to_pdf(filepath)
    pdf_filename = os.path.basename(pdf_path)

    html = f'''<iframe 
        src="/api/pdf/{doc_id}#toolbar=0&navpanes=0&scrollbar=0&view=FitH&pagemode=none"
        style="width:100%;height:100%;border:none;display:block;position:absolute;inset:0;"
        allowfullscreen
    ></iframe>'''
    return [{'title': original_name, 'html': html, 'isPdf': True, 'pdfId': doc_id}]

def convert_in_background(doc_id, filepath, filename):
    """Run PDF conversion in background thread, broadcast when done"""
    try:
        pages = parse_file_to_pdf(filepath, filename, doc_id)
        documents[doc_id]['pages'] = pages
        documents[doc_id]['status'] = 'ready'
        broadcast({'type': 'DOCUMENT_READY', 'document': doc_summary(documents[doc_id])})
    except Exception as e:
        documents[doc_id]['status'] = 'error'
        documents[doc_id]['error'] = str(e)
        broadcast({'type': 'DOCUMENT_ERROR', 'id': doc_id, 'error': str(e)})

# ─── Routes ────────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return send_from_directory(os.path.join(BASE_DIR, 'public/dashboard'), 'index.html')

@app.route('/dashboard')
def dashboard():
    return send_from_directory(os.path.join(BASE_DIR, 'public/dashboard'), 'index.html')

@app.route('/display/<center_ref>')
def display(center_ref):
    return send_from_directory(os.path.join(BASE_DIR, 'public/display'), 'index.html')

@app.route('/public/<path:filename>')
def static_files(filename):
    return send_from_directory(os.path.join(BASE_DIR, 'public'), filename)

@app.route('/logo.png')
def serve_logo():
    return send_from_directory(BASE_DIR + '/public', 'logo.png', mimetype='image/png')

@app.route('/dashboard/<path:filename>')
def dashboard_static(filename):
    return send_from_directory(os.path.join(BASE_DIR, 'public/dashboard'), filename)

@app.route('/display/manifest.json')
def display_manifest():
    return send_from_directory(os.path.join(BASE_DIR, 'public/display'), 'manifest.json')

# ─── Documents API ─────────────────────────────────────────────────────────
@app.route('/api/pdf/<doc_id>')
def serve_pdf(doc_id):
    doc = documents.get(doc_id)
    if not doc:
        return jsonify({'error': 'Not found'}), 404
    pdf_path = os.path.splitext(doc['filePath'])[0] + '.pdf'
    if not os.path.exists(pdf_path):
        return jsonify({'error': 'PDF not found'}), 404
    return send_from_directory(os.path.dirname(pdf_path), os.path.basename(pdf_path), mimetype='application/pdf')

@app.route('/api/documents', methods=['GET'])
def get_documents():
    return jsonify([doc_summary(d) for d in documents.values()])

@app.route('/api/documents/upload', methods=['POST'])
def upload_document():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    filename = file.filename
    ext = os.path.splitext(filename)[1].lower()

    if ext not in ['.xlsx', '.xls', '.docx', '.doc']:
        return jsonify({'error': 'Only Excel and Word files allowed'}), 400

    saved_name = f"{int(datetime.now().timestamp())}_{filename}"
    filepath = os.path.join(UPLOAD_DIR, saved_name)
    file.save(filepath)

    doc_id = str(uuid.uuid4())

    # Create doc entry immediately with 'converting' status
    doc_name = request.form.get('name', filename)
    doc = {
        'id': doc_id,
        'name': doc_name,
        'originalName': filename,
        'type': 'pdf',
        'pages': [],
        'status': 'converting',
        'uploadedAt': datetime.now().isoformat(),
        'filePath': filepath
    }
    documents[doc['id']] = doc
    broadcast({'type': 'DOCUMENT_ADDED', 'document': doc_summary(doc)})

    # Convert in background
    t = threading.Thread(target=convert_in_background, args=(doc_id, filepath, filename), daemon=True)
    t.start()

    return jsonify({'success': True, 'document': doc_summary(doc)})

@app.route('/api/documents/<doc_id>/full', methods=['GET'])
def get_document_full(doc_id):
    doc = documents.get(doc_id)
    if not doc:
        return jsonify({'error': 'Not found'}), 404
    return jsonify(doc)

@app.route('/api/documents/<doc_id>', methods=['DELETE'])
def delete_document(doc_id):
    doc = documents.get(doc_id)
    if not doc:
        return jsonify({'error': 'Not found'}), 404

    for center in centers.values():
        if center.get('assignedDoc') and center['assignedDoc']['id'] == doc_id:
            center['assignedDoc'] = None
            center['currentPage'] = 0
            notify_center(center['id'], {'type': 'DOCUMENT_REMOVED'})
            broadcast({'type': 'CENTER_UPDATED', 'center': center})

    try:
        os.remove(doc['filePath'])
    except Exception:
        pass

    del documents[doc_id]
    broadcast({'type': 'DOCUMENT_DELETED', 'id': doc_id})
    return jsonify({'success': True})

# ─── Centers API ───────────────────────────────────────────────────────────
@app.route('/api/centers/resolve/<center_ref>', methods=['GET'])
def resolve_center(center_ref):
    # Try by slug first, then by ID
    center = next((c for c in centers.values() if c.get('slug') == center_ref), None)
    if not center:
        center = centers.get(center_ref)
    if not center:
        return jsonify({'error': 'Center not found'}), 404
    return jsonify({'id': center['id']})

@app.route('/api/centers', methods=['GET'])
def get_centers():
    return jsonify(list(centers.values()))

@app.route('/api/centers', methods=['POST'])
def create_center():
    data = request.json or {}
    name = data.get('name', '').strip()
    if not name:
        return jsonify({'error': 'Name required'}), 400

    center = {
        'id': str(uuid.uuid4()),
        'name': name,
        'slug': make_slug(name),
        'color': data.get('color', '#2563eb'),
        'assignedDoc': None,
        'currentPage': 0,
        'connected': False,
        'createdAt': datetime.now().isoformat()
    }
    centers[center['id']] = center
    broadcast({'type': 'CENTER_ADDED', 'center': center})
    return jsonify({'success': True, 'center': center})

@app.route('/api/centers/<center_id>', methods=['PUT'])
def update_center(center_id):
    center = centers.get(center_id)
    if not center:
        return jsonify({'error': 'Not found'}), 404
    data = request.json or {}
    if 'name' in data:
        center['name'] = data['name']
    if 'color' in data:
        center['color'] = data['color']
    broadcast({'type': 'CENTER_UPDATED', 'center': center})
    return jsonify({'success': True, 'center': center})

@app.route('/api/centers/<center_id>', methods=['DELETE'])
def delete_center(center_id):
    if center_id not in centers:
        return jsonify({'error': 'Not found'}), 404
    del centers[center_id]
    broadcast({'type': 'CENTER_DELETED', 'id': center_id})
    return jsonify({'success': True})

@app.route('/api/centers/<center_id>/assign', methods=['POST'])
def assign_document(center_id):
    center = centers.get(center_id)
    if not center:
        return jsonify({'error': 'Center not found'}), 404

    data = request.json or {}
    doc_id = data.get('documentId')

    if not doc_id:
        center['assignedDoc'] = None
        center['currentPage'] = 0
        notify_center(center_id, {'type': 'DOCUMENT_REMOVED'})
    else:
        doc = documents.get(doc_id)
        if not doc:
            return jsonify({'error': 'Document not found'}), 404
        center['assignedDoc'] = {'id': doc['id'], 'name': doc['name'], 'type': doc['type']}
        center['currentPage'] = 0
        notify_center(center_id, {'type': 'DOCUMENT_ASSIGNED', 'document': doc, 'currentPage': 0})

    broadcast({'type': 'CENTER_UPDATED', 'center': center})
    return jsonify({'success': True, 'center': center})

@app.route('/api/centers/<center_id>/page', methods=['POST'])
def set_page(center_id):
    center = centers.get(center_id)
    if not center:
        return jsonify({'error': 'Center not found'}), 404
    if not center.get('assignedDoc'):
        return jsonify({'error': 'No document assigned'}), 400

    doc = documents.get(center['assignedDoc']['id'])
    if not doc:
        return jsonify({'error': 'Document not found'}), 404

    page = int((request.json or {}).get('page', 0))
    max_page = len(doc['pages']) - 1
    center['currentPage'] = max(0, min(page, max_page))

    notify_center(center_id, {
        'type': 'PAGE_CHANGE',
        'page': center['currentPage'],
        'totalPages': len(doc['pages'])
    })
    broadcast({'type': 'CENTER_UPDATED', 'center': center})
    return jsonify({'success': True, 'currentPage': center['currentPage']})

# ─── WebSocket ─────────────────────────────────────────────────────────────
@sock.route('/ws')
def websocket(ws):
    client_id = str(uuid.uuid4())
    clients[client_id] = {'ws': ws, 'role': None, 'centerId': None}

    # Send initial state
    try:
        ws.send(json.dumps({
            'type': 'INIT',
            'centers': list(centers.values()),
            'documents': [doc_summary(d) for d in documents.values()]
        }))
    except Exception:
        return

    try:
        while True:
            raw = ws.receive()
            if raw is None:
                break
            try:
                msg = json.loads(raw)
                handle_message(client_id, msg)
            except Exception as e:
                print(f'WS error: {e}')
    finally:
        client = clients.pop(client_id, None)
        if client and client.get('centerId'):
            cid = client['centerId']
            if cid in centers:
                centers[cid]['connected'] = False
                broadcast({'type': 'CENTER_UPDATED', 'center': centers[cid]})

def handle_message(client_id, msg):
    client = clients.get(client_id)
    if not client:
        return

    if msg.get('type') == 'REGISTER_DASHBOARD':
        client['role'] = 'dashboard'

    elif msg.get('type') == 'REGISTER_DISPLAY':
        client['role'] = 'display'
        center_id = msg.get('centerId')
        client['centerId'] = center_id

        center = centers.get(center_id)
        if center:
            center['connected'] = True
            broadcast({'type': 'CENTER_UPDATED', 'center': center})

            if center.get('assignedDoc'):
                doc = documents.get(center['assignedDoc']['id'])
                if doc:
                    try:
                        client['ws'].send(json.dumps({
                            'type': 'DOCUMENT_ASSIGNED',
                            'document': doc,
                            'currentPage': center['currentPage']
                        }))
                    except Exception:
                        pass

# ─── Start ─────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    PORT = 8443
    hostname = socket.gethostname()
    try:
        local_ip = socket.gethostbyname(hostname)
    except:
        local_ip = 'localhost'

    # Use HTTPS if cert files exist
    cert_file = os.path.join(BASE_DIR, 'cert.pem')
    key_file  = os.path.join(BASE_DIR, 'key.pem')
    ssl_ctx = None
    protocol = 'http'

    if os.path.exists(cert_file) and os.path.exists(key_file):
        ssl_ctx = (cert_file, key_file)
        protocol = 'https'

    print('\n🏭 Assembly Hub Server Running')
    print(f'   Base Dir  : {BASE_DIR}')
    logo_path = os.path.join(BASE_DIR, 'public', 'logo.png')
    print(f'   Logo      : {logo_path} ({"EXISTS" if os.path.exists(logo_path) else "MISSING!"})')
    print(f'   Dashboard : {protocol}://{local_ip}:{PORT}/dashboard')
    print(f'   Local     : {protocol}://localhost:{PORT}/dashboard')
    if protocol == 'https':
        print('   🔒 HTTPS enabled — PWA install available')
    print(f'   Network IP: {local_ip}\n')

    app.run(host='0.0.0.0', port=PORT, debug=False, ssl_context=ssl_ctx)
