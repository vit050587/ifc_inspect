from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import os
import sys
import uuid
import json

# Add parent directory to path for scripts imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from scripts.page_splitter import extract_pages_from_pdf
from scripts.ifc_parser import parse_ifc_file
from scripts.materials_summary import main as create_materials_summary
from scripts.ifc_viewer import prepare_ifc_for_viewer

app = Flask(__name__, static_folder='../frontend', static_url_path='')
CORS(app)

# Use absolute path for uploads folder
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def serve_frontend():
    return send_from_directory(app.static_folder, 'index.html')

@app.route('/api/upload', methods=['POST'])
def upload_files():
    """Handle file uploads (PDF and IFC files)"""
    if 'files' not in request.files:
        return jsonify({'error': 'No files uploaded'}), 400
    
    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': 'No files selected'}), 400
    
    session_id = str(uuid.uuid4())
    session_folder = os.path.join(UPLOAD_FOLDER, session_id)
    os.makedirs(session_folder, exist_ok=True)
    
    pdf_files = []
    ifc_files = []
    other_files = []
    
    # Save files and categorize them
    for file in files:
        if file.filename:
            filepath = os.path.join(session_folder, file.filename)
            file.save(filepath)
            
            if filepath.lower().endswith('.pdf'):
                pdf_files.append(filepath)
            elif filepath.lower().endswith('.ifc'):
                ifc_files.append(filepath)
            else:
                other_files.append(filepath)
    
    # Initialize session info
    session_info = {
        'session_id': session_id,
        'status': 'files_uploaded',
        'pdf_files': [os.path.basename(f) for f in pdf_files],
        'ifc_files': [os.path.basename(f) for f in ifc_files],
        'other_files': [os.path.basename(f) for f in other_files]
    }
    
    # Save session info
    session_info_path = os.path.join(session_folder, 'session_info.json')
    with open(session_info_path, 'w', encoding='utf-8') as f:
        json.dump(session_info, f, ensure_ascii=False, indent=2)
    
    return jsonify({
        'session_id': session_id,
        'pdf_count': len(pdf_files),
        'ifc_count': len(ifc_files),
        'other_count': len(other_files),
        'message': f'Uploaded {len(pdf_files)} PDF(s) and {len(ifc_files)} IFC file(s)'
    })


@app.route('/api/process', methods=['POST'])
def process_files():
    """Process uploaded files: split PDF pages and parse IFC"""
    data = request.json
    session_id = data.get('session_id')
    
    if not session_id:
        return jsonify({'error': 'Missing session_id'}), 400
    
    session_folder = os.path.join(UPLOAD_FOLDER, session_id)
    if not os.path.exists(session_folder):
        return jsonify({'error': 'Session not found'}), 404
    
    # Load session info
    session_info_path = os.path.join(session_folder, 'session_info.json')
    with open(session_info_path, 'r', encoding='utf-8') as f:
        session_info = json.load(f)
    
    # Update status
    session_info['status'] = 'processing'
    
    results = {
        'text_pages': [],
        'drawing_pages': [],
        'ifc_results': None
    }
    
    # Process PDF files - split into text and drawing pages
    pdf_files = [os.path.join(session_folder, f) for f in session_info.get('pdf_files', [])]
    for pdf_path in pdf_files:
        if os.path.exists(pdf_path):
            print(f"Processing PDF: {pdf_path}")
            text_pages, drawing_pages = extract_pages_from_pdf(pdf_path, session_folder)
            results['text_pages'].extend(text_pages)
            results['drawing_pages'].extend(drawing_pages)
    
    # Process IFC files - FIRST prepare for viewer, THEN parse
    ifc_files = [os.path.join(session_folder, f) for f in session_info.get('ifc_files', [])]
    ifc_viewer_info = []
    
    for ifc_path in ifc_files:
        if os.path.exists(ifc_path):
            # Step 1: Prepare IFC for viewer (runs first)
            print(f"Preparing IFC for viewer: {ifc_path}")
            try:
                viewer_result = prepare_ifc_for_viewer(ifc_path, session_folder)
                ifc_viewer_info.append(viewer_result)
                session_info['ifc_viewer_info'] = ifc_viewer_info
                print(f"✓ IFC ready for viewing: {viewer_result.get('filename', 'unknown')}")
            except Exception as e:
                print(f"Error preparing IFC viewer for {ifc_path}: {e}")
                session_info['ifc_viewer_error'] = str(e)
            
            # Step 2: Parse IFC file (runs after viewer preparation)
            print(f"Processing IFC: {ifc_path}")
            try:
                ifc_result = parse_ifc_file(ifc_path, session_folder)
                results['ifc_results'] = ifc_result
                session_info['ifc_excel_file'] = ifc_result.get('excel_filename')
                
                # Создаем сводную таблицу по материалам после создания IFC отчета
                if ifc_result.get('excel_filename'):
                    excel_path = os.path.join(session_folder, ifc_result['excel_filename'])
                    if os.path.exists(excel_path):
                        print(f"Создание сводной таблицы по материалам: {excel_path}")
                        materials_excel_path = create_materials_summary(excel_path)
                        session_info['materials_excel_file'] = os.path.basename(materials_excel_path)
                        print(f"Сводная таблица сохранена: {session_info['materials_excel_file']}")
            except Exception as e:
                print(f"Error parsing IFC file {ifc_path}: {e}")
                session_info['ifc_error'] = str(e)
    
    # Save results
    results_path = os.path.join(session_folder, 'results.json')
    with open(results_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2, default=str)
    
    # Update session status
    session_info['status'] = 'completed'
    session_info['results_summary'] = {
        'text_pages_count': len(results['text_pages']),
        'drawing_pages_count': len(results['drawing_pages']),
        'ifc_processed': results['ifc_results'] is not None
    }
    
    with open(session_info_path, 'w', encoding='utf-8') as f:
        json.dump(session_info, f, ensure_ascii=False, indent=2)
    
    return jsonify({
        'success': True,
        'session_id': session_id,
        'text_pages_count': len(results['text_pages']),
        'drawing_pages_count': len(results['drawing_pages']),
        'ifc_processed': results['ifc_results'] is not None,
        'ifc_excel_file': session_info.get('ifc_excel_file'),
        'materials_excel_file': session_info.get('materials_excel_file'),
        'summary': session_info['results_summary']
    })


@app.route('/api/viewer/<session_id>/<filename>')
def serve_ifc_viewer_file(session_id, filename):
    """Serve IFC file for the web viewer"""
    session_folder = os.path.join(UPLOAD_FOLDER, session_id)
    viewer_folder = os.path.join(session_folder, 'viewer')
    
    if not os.path.exists(viewer_folder):
        return jsonify({'error': 'Viewer folder not found'}), 404
    
    file_path = os.path.join(viewer_folder, filename)
    if not os.path.exists(file_path):
        return jsonify({'error': 'IFC file not found'}), 404
    
    # Serve with proper MIME type for IFC files
    response = send_from_directory(viewer_folder, filename, mimetype='application/x-step')
    response.headers['Content-Type'] = 'application/x-step'
    return response


@app.route('/api/download/<session_id>/<filename>')
def download_file(session_id, filename):
    """Download processed files (Excel reports, etc.)"""
    session_folder = os.path.join(UPLOAD_FOLDER, session_id)
    if not os.path.exists(session_folder):
        return jsonify({'error': 'Session not found'}), 404
    
    file_path = os.path.join(session_folder, filename)
    if not os.path.exists(file_path):
        return jsonify({'error': 'File not found'}), 404
    
    return send_from_directory(session_folder, filename, as_attachment=True)


@app.route('/api/results/<session_id>')
def get_results(session_id):
    """Get processing results for a session"""
    session_folder = os.path.join(UPLOAD_FOLDER, session_id)
    if not os.path.exists(session_folder):
        return jsonify({'error': 'Session not found'}), 404
    
    results_path = os.path.join(session_folder, 'results.json')
    if not os.path.exists(results_path):
        return jsonify({'error': 'Results not found'}), 404
    
    with open(results_path, 'r', encoding='utf-8') as f:
        results = json.load(f)
    
    # Load session info
    session_info_path = os.path.join(session_folder, 'session_info.json')
    with open(session_info_path, 'r', encoding='utf-8') as f:
        session_info = json.load(f)
    
    return jsonify({
        'session_info': session_info,
        'results': results
    })


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=6003)
