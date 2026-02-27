import re
import os
import json
import uuid
import threading
import socket
import sqlite3
from datetime import datetime
from flask import Flask, request, jsonify, send_from_directory
from flask_sock import Sock

app = Flask(__name__, static_folder='public')
sock = Sock(app)

# --- Configuration ---
UPLOAD_DIR = os.path.join(os.path.dirname(__file__), 'uploads')
DB_PATH = os.path.join(os.path.dirname(__file__), 'hub.db')
os.makedirs(UPLOAD_DIR, exist_ok=True)

# --- Database Integration ---
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db() as conn:
        conn.execute('''
            CREATE TABLE IF NOT EXISTS documents (
                id TEXT PRIMARY KEY,
                name TEXT,
                originalName TEXT,
                type TEXT,
                status TEXT,
                uploadedAt TEXT,
                filePath TEXT
            )
        ''')
        conn.execute('''
            CREATE TABLE IF NOT EXISTS centers (
                id TEXT PRIMARY KEY,
                name TEXT,
                slug TEXT,
                color TEXT,
                assignedDocId TEXT,
                currentPage INTEGER DEFAULT 0,
                createdAt TEXT,
                FOREIGN KEY(assignedDocId) REFERENCES documents(id)
            )
        ''')
    print("✅ Database initialized")

# --- In-memory transient state ---
clients = {}  # { id: { ws, role, centerId } }
connected_centers = set() # Track online status in-memory

# --- Helpers ---
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

def get_doc_by_id(doc_id):
    with get_db() as conn:
        row = conn.execute('SELECT * FROM documents WHERE id = ?', (doc_id,)).fetchone()
        return dict(row) if row else None

def get_center_by_id(center_id):
    with get_db() as conn:
        row = conn.execute('SELECT * FROM centers WHERE id = ?', (center_id,)).fetchone()
        if not row: return None
        center = dict(row)
        center['connected'] = center['id'] in connected_centers
        if center['assignedDocId']:
            doc = get_doc_by_id(center['assignedDocId'])
            center['assignedDoc'] = {'id': doc['id'], 'name': doc['name'], 'type': doc['type']} if doc else None
        else:
            center['assignedDoc'] = None
        return center

def doc_summary(doc):
    # Pages are calculated based on the existence of the PDF file
    pdf_path = os.path.splitext(doc['filePath'])[0] + '.pdf'
    page_count = 1 if os.path.exists(pdf_path) else 0
    return {
        'id': doc['id'],
        'name': doc['name'],
        'type': doc['type'],
        'pageCount': page_count,
        'uploadedAt': doc['uploadedAt'],
        'status': doc['status']
    }

def make_slug(name):
    slug = name.lower().strip()
    slug = re.sub(r'[^a-z0-9]+', '-', slug).strip('-')
    base = slug
    counter = 2
    with get_db() as conn:
        while conn.execute('SELECT id FROM centers WHERE slug = ?', (slug,)).fetchone():
            slug = f"{base}-{counter}"
            counter += 1
    return slug

# --- File Parsing ---
def convert_to_pdf(filepath):
    import subprocess, shutil
    ext = os.path.splitext(filepath)[1].lower()
    pdf_path = os.path.splitext(filepath)[0] + '.pdf'

    if os.path.exists(pdf_path):
        os.remove(pdf_path)

    try:
        import win32com.client
        import pythoncom
        abs_path = os.path.abspath(filepath)
        abs_pdf  = os.path.abspath(pdf_path)

        pythoncom.CoInitialize()
        app = None
        try:
            if ext in ['.xlsx', '.xls']:
                app = win32com.client.Dispatch('Excel.Application')
                app.Visible = False
                app.DisplayAlerts = False
                wb = app.Workbooks.Open(abs_path, UpdateLinks=0, ReadOnly=True)
                wb.ExportAsFixedFormat(0, abs_pdf)
                wb.Close(False)
            else:
                app = win32com.client.Dispatch('Word.Application')
                app.Visible = False
                doc = app.Documents.Open(abs_path, ReadOnly=True)
                doc.SaveAs(abs_pdf, FileFormat=17)
                doc.Close()
        finally:
            if app:
                try: app.Quit()
                except: pass
            pythoncom.CoUninitialize()

        if os.path.exists(pdf_path): return pdf_path
    except Exception: pass

    lo = shutil.which('libreoffice') or shutil.which('soffice')
    if lo:
        out_dir = os.path.dirname(filepath)
        subprocess.run([lo, '--headless', '--convert-to', 'pdf', '--outdir', out_dir, filepath], capture_output=True, timeout=60)
        if os.path.exists(pdf_path): return pdf_path

    raise Exception('No PDF converter available.')

def convert_in_background(doc_id, filepath, original_name):
    try:
        convert_to_pdf(filepath)
        with get_db() as conn:
            conn.execute('UPDATE documents SET status = ? WHERE id = ?', ('ready', doc_id))
        
        doc = get_doc_by_id(doc_id)
        broadcast({'type': 'DOCUMENT_READY', 'document': doc_summary(doc)})
    except Exception as e:
        with get_db() as conn:
            conn.execute('UPDATE documents SET status = ? WHERE id = ?', ('error', doc_id))
        broadcast({'type': 'DOCUMENT_ERROR', 'id': doc_id, 'error': str(e)})

# --- Routes ---
@app.route('/dashboard')
@app.route('/')
def dashboard():
    return send_from_directory('public/dashboard', 'index.html')

@app.route('/display/<center_ref>')
def display(center_ref):
    return send_from_directory('public/display', 'index.html')

@app.route('/api/pdf/<doc_id>')
def serve_pdf(doc_id):
    doc = get_doc_by_id(doc_id)
    if not doc: return jsonify({'error': 'Not found'}), 404
    pdf_path = os.path.splitext(doc['filePath'])[0] + '.pdf'
    if not os.path.exists(pdf_path): return jsonify({'error': 'PDF not found'}), 404
    return send_from_directory(os.path.dirname(pdf_path), os.path.basename(pdf_path), mimetype='application/pdf')

@app.route('/api/documents', methods=['GET'])
def get_documents():
    with get_db() as conn:
        rows = conn.execute('SELECT * FROM documents').fetchall()
        return jsonify([doc_summary(dict(r)) for r in rows])

@app.route('/api/documents/upload', methods=['POST'])
def upload_document():
    file = request.files.get('file')
    if not file: return jsonify({'error': 'No file'}), 400
    
    filename = file.filename
    ext = os.path.splitext(filename)[1].lower()
    if ext not in ['.xlsx', '.xls', '.docx', '.doc']:
        return jsonify({'error': 'Unsupported file type'}), 400

    doc_id = str(uuid.uuid4())
    saved_name = f"{int(datetime.now().timestamp())}_{filename}"
    filepath = os.path.join(UPLOAD_DIR, saved_name)
    file.save(filepath)

    doc = {
        'id': doc_id,
        'name': request.form.get('name', filename),
        'originalName': filename,
        'type': 'pdf',
        'status': 'converting',
        'uploadedAt': datetime.now().isoformat(),
        'filePath': filepath
    }

    with get_db() as conn:
        conn.execute('INSERT INTO documents VALUES (?,?,?,?,?,?,?)', 
                    (doc['id'], doc['name'], doc['originalName'], doc['type'], doc['status'], doc['uploadedAt'], doc['filePath']))

    broadcast({'type': 'DOCUMENT_ADDED', 'document': doc_summary(doc)})
    threading.Thread(target=convert_in_background, args=(doc_id, filepath, filename), daemon=True).start()
    return jsonify({'success': True, 'document': doc_summary(doc)})

@app.route('/api/documents/<doc_id>/full', methods=['GET'])
def get_document_full(doc_id):
    doc = get_doc_by_id(doc_id)
    if not doc: return jsonify({'error': 'Not found'}), 404
    
    # Structure match for frontend expects 'pages'
    pdf_url = f"/api/pdf/{doc_id}"
    html = f'<iframe src="{pdf_url}#toolbar=0&navpanes=0&scrollbar=0&view=FitH&pagemode=none" style="width:100%;height:100%;border:none;display:block;position:absolute;inset:0;" allowfullscreen></iframe>'
    doc['pages'] = [{'title': doc['name'], 'html': html, 'isPdf': True, 'pdfId': doc_id}]
    return jsonify(doc)

@app.route('/api/documents/<doc_id>', methods=['DELETE'])
def delete_document(doc_id):
    doc = get_doc_by_id(doc_id)
    if not doc: return jsonify({'error': 'Not found'}), 404

    with get_db() as conn:
        # Clear assignments
        conn.execute('UPDATE centers SET assignedDocId = NULL, currentPage = 0 WHERE assignedDocId = ?', (doc_id,))
        conn.execute('DELETE FROM documents WHERE id = ?', (doc_id,))

    try: os.remove(doc['filePath'])
    except: pass
    
    broadcast({'type': 'DOCUMENT_DELETED', 'id': doc_id})
    return jsonify({'success': True})

# --- Centers API ---
@app.route('/api/centers', methods=['GET'])
def get_centers():
    with get_db() as conn:
        rows = conn.execute('SELECT id FROM centers').fetchall()
        return jsonify([get_center_by_id(r['id']) for r in rows])

@app.route('/api/centers', methods=['POST'])
def create_center():
    data = request.json or {}
    name = data.get('name', '').strip()
    if not name: return jsonify({'error': 'Name required'}), 400

    center_id = str(uuid.uuid4())
    slug = make_slug(name)
    color = data.get('color', '#2563eb')
    now = datetime.now().isoformat()

    with get_db() as conn:
        conn.execute('INSERT INTO centers (id, name, slug, color, createdAt) VALUES (?,?,?,?,?)',
                    (center_id, name, slug, color, now))

    center = get_center_by_id(center_id)
    broadcast({'type': 'CENTER_ADDED', 'center': center})
    return jsonify({'success': True, 'center': center})

@app.route('/api/centers/resolve/<center_ref>', methods=['GET'])
def resolve_center(center_ref):
    with get_db() as conn:
        row = conn.execute('SELECT id FROM centers WHERE slug = ? OR id = ?', (center_ref, center_ref)).fetchone()
        if not row: return jsonify({'error': 'Not found'}), 404
        return jsonify({'id': row['id']})

@app.route('/api/centers/<center_id>/assign', methods=['POST'])
def assign_document(center_id):
    doc_id = (request.json or {}).get('documentId')
    with get_db() as conn:
        conn.execute('UPDATE centers SET assignedDocId = ?, currentPage = 0 WHERE id = ?', (doc_id, center_id))

    center = get_center_by_id(center_id)
    if not doc_id:
        notify_center(center_id, {'type': 'DOCUMENT_REMOVED'})
    else:
        doc_full = get_document_full(doc_id).json
        notify_center(center_id, {'type': 'DOCUMENT_ASSIGNED', 'document': doc_full, 'currentPage': 0})

    broadcast({'type': 'CENTER_UPDATED', 'center': center})
    return jsonify({'success': True, 'center': center})

@app.route('/api/centers/<center_id>/page', methods=['POST'])
def set_page(center_id):
    page = int((request.json or {}).get('page', 0))
    with get_db() as conn:
        conn.execute('UPDATE centers SET currentPage = ? WHERE id = ?', (page, center_id))
    
    center = get_center_by_id(center_id)
    notify_center(center_id, {'type': 'PAGE_CHANGE', 'page': page})
    broadcast({'type': 'CENTER_UPDATED', 'center': center})
    return jsonify({'success': True})

# --- WebSocket ---
@sock.route('/ws')
def websocket(ws):
    client_id = str(uuid.uuid4())
    clients[client_id] = {'ws': ws, 'role': None, 'centerId': None}

    # Send initial state
    with get_db() as conn:
        docs = [doc_summary(dict(r)) for r in conn.execute('SELECT * FROM documents').fetchall()]
        center_ids = [r['id'] for r in conn.execute('SELECT id FROM centers').fetchall()]
        all_centers = [get_center_by_id(cid) for cid in center_ids]
        
    try:
        ws.send(json.dumps({'type': 'INIT', 'centers': all_centers, 'documents': docs}))
        while True:
            raw = ws.receive()
            if raw is None: break
            msg = json.loads(raw)
            
            if msg.get('type') == 'REGISTER_DASHBOARD':
                clients[client_id]['role'] = 'dashboard'
            
            elif msg.get('type') == 'REGISTER_DISPLAY':
                cid = msg.get('centerId')
                clients[client_id].update({'role': 'display', 'centerId': cid})
                connected_centers.add(cid)
                center = get_center_by_id(cid)
                broadcast({'type': 'CENTER_UPDATED', 'center': center})
                if center.get('assignedDoc'):
                    doc_full = get_document_full(center['assignedDoc']['id']).json
                    ws.send(json.dumps({'type': 'DOCUMENT_ASSIGNED', 'document': doc_full, 'currentPage': center['currentPage']}))
    finally:
        client = clients.pop(client_id, None)
        if client and client.get('centerId'):
            cid = client['centerId']
            # Only remove if no other tabs are open for this center
            if not any(c.get('centerId') == cid for c in clients.values()):
                connected_centers.discard(cid)
                center = get_center_by_id(cid)
                if center: broadcast({'type': 'CENTER_UPDATED', 'center': center})

if __name__ == '__main__':
    init_db()
    PORT = 8443
    cert_file = os.path.join(os.path.dirname(__file__), 'cert.pem')
    key_file  = os.path.join(os.path.dirname(__file__), 'key.pem')
    ssl_ctx = (cert_file, key_file) if os.path.exists(cert_file) else None
    
    print(f"\n🏭 Assembly Hub Server Running on port {PORT}")
    app.run(host='0.0.0.0', port=PORT, debug=False, ssl_context=ssl_ctx)