from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import os
import sys
import uuid
import json
import shutil

# Add parent directory to path for scripts imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from scripts.page_splitter import extract_pages_from_pdf
from scripts.ifc_parser import parse_ifc_file
from scripts.materials_summary import main as create_materials_summary
from scripts.ifc_viewer import prepare_ifc_for_viewer
from scripts.pdf_classifier import classify_pdf_files

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
    """Handle file uploads (PDF, IFC, and Excel files)"""
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
    excel_files = []
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
            elif filepath.lower().endswith(('.xlsx', '.xls')):
                excel_files.append(filepath)
            else:
                other_files.append(filepath)
    
    # Initialize session info
    session_info = {
        'session_id': session_id,
        'status': 'files_uploaded',
        'pdf_files': [os.path.basename(f) for f in pdf_files],
        'ifc_files': [os.path.basename(f) for f in ifc_files],
        'excel_files': [os.path.basename(f) for f in excel_files],
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
        'excel_count': len(excel_files),
        'other_count': len(other_files),
        'message': f'Uploaded {len(pdf_files)} PDF(s), {len(ifc_files)} IFC file(s), {len(excel_files)} Excel file(s)'
    })


@app.route('/api/process', methods=['POST'])
def process_files():
    """Process uploaded files: organize by type, classify PDFs, parse IFC"""
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
        'ifc_results': None,
        'pdf_classification': None,
        'materials_summary': None
    }
    
    # Step 1: Organize files into folders (IFC, Excel, PDF)
    print("\n" + "="*60)
    print("📁 ШАГ 1: ОРГАНИЗАЦИЯ ФАЙЛОВ ПО ПАПКАМ")
    print("="*60)
    
    # Create organized folders
    ifc_folder = os.path.join(session_folder, 'ifc_models')
    excel_folder = os.path.join(session_folder, 'спецификации')
    pdf_folder = os.path.join(session_folder, 'pdf_документы')
    
    os.makedirs(ifc_folder, exist_ok=True)
    os.makedirs(excel_folder, exist_ok=True)
    os.makedirs(pdf_folder, exist_ok=True)
    
    # Move IFC files
    ifc_files = [os.path.join(session_folder, f) for f in session_info.get('ifc_files', [])]
    for ifc_path in ifc_files:
        if os.path.exists(ifc_path):
            dest = os.path.join(ifc_folder, os.path.basename(ifc_path))
            if ifc_path != dest:
                shutil.move(ifc_path, dest)
                print(f"   📦 IFC: {os.path.basename(ifc_path)} -> ifc_models/")
    
    # Move Excel files
    excel_files = [os.path.join(session_folder, f) for f in session_info.get('excel_files', [])]
    for excel_path in excel_files:
        if os.path.exists(excel_path):
            dest = os.path.join(excel_folder, os.path.basename(excel_path))
            if excel_path != dest:
                shutil.move(excel_path, dest)
                print(f"   📊 Excel: {os.path.basename(excel_path)} -> спецификации/")
    
    # Move PDF files to pdf_folder for classification
    pdf_files = [os.path.join(session_folder, f) for f in session_info.get('pdf_files', [])]
    for pdf_path in pdf_files:
        if os.path.exists(pdf_path):
            dest = os.path.join(pdf_folder, os.path.basename(pdf_path))
            if pdf_path != dest:
                shutil.move(pdf_path, dest)
                print(f"   📄 PDF: {os.path.basename(pdf_path)} -> pdf_документы/")
    
    # Step 2: Classify PDF files using LLM
    print("\n" + "="*60)
    print("🤖 ШАГ 2: КЛАССИФИКАЦИЯ PDF ДОКУМЕНТОВ (LLM)")
    print("="*60)
    
    try:
        pdf_classification = classify_pdf_files(session_folder)
        results['pdf_classification'] = {
            'categories': {cat: len(files) for cat, files in pdf_classification.items()},
            'files': pdf_classification
        }
        session_info['pdf_classification'] = results['pdf_classification']
        print("✅ Классификация PDF завершена")
    except Exception as e:
        print(f"❌ Ошибка классификации PDF: {e}")
        session_info['pdf_classification_error'] = str(e)
    
    # Step 3: Extract drawings from "пояснительная записка" PDFs
    print("\n" + "="*60)
    print("🏗️ ШАГ 3: ИЗВЛЕЧЕНИЕ ЧЕРТЕЖЕЙ ИЗ ПОЯСНИТЕЛЬНОЙ ЗАПИСКИ")
    print("="*60)
    
    explanatory_folder = os.path.join(session_folder, 'пояснительная_записка')
    drawings_folder = os.path.join(session_folder, 'чертежи')
    os.makedirs(drawings_folder, exist_ok=True)
    
    if os.path.exists(explanatory_folder):
        # Find PDF files in explanatory folder
        explanatory_pdfs = [f for f in os.listdir(explanatory_folder) if f.lower().endswith('.pdf')]
        
        for pdf_filename in explanatory_pdfs:
            pdf_path = os.path.join(explanatory_folder, pdf_filename)
            print(f"Обработка PDF: {pdf_filename}")
            
            try:
                text_pages, drawing_pages = extract_pages_from_pdf(pdf_path, session_folder)
                
                # Move drawing pages to drawings folder
                for page_info in drawing_pages:
                    src_path = page_info['path']
                    dest_path = os.path.join(drawings_folder, os.path.basename(src_path))
                    if os.path.exists(src_path) and src_path != dest_path:
                        shutil.copy2(src_path, dest_path)
                        print(f"   🏗️ Чертеж: {os.path.basename(src_path)} -> чертежи/")
                
                # Move text pages to explanatory folder if not already there
                for page_info in text_pages:
                    src_path = page_info['path']
                    # Text pages stay in text_pages folder or move to explanatory
                
                results['text_pages'].extend(text_pages)
                results['drawing_pages'].extend(drawing_pages)
                
            except Exception as e:
                print(f"❌ Ошибка извлечения страниц: {e}")
    
    # Also process any remaining PDFs in pdf_folder that weren't classified
    remaining_pdfs = [f for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')]
    for pdf_filename in remaining_pdfs:
        pdf_path = os.path.join(pdf_folder, pdf_filename)
        print(f"Обработка неклассифицированного PDF: {pdf_filename}")
        
        try:
            text_pages, drawing_pages = extract_pages_from_pdf(pdf_path, session_folder)
            
            for page_info in drawing_pages:
                src_path = page_info['path']
                dest_path = os.path.join(drawings_folder, os.path.basename(src_path))
                if os.path.exists(src_path) and src_path != dest_path:
                    shutil.copy2(src_path, dest_path)
                    print(f"   🏗️ Чертеж: {os.path.basename(src_path)} -> чертежи/")
            
            results['text_pages'].extend(text_pages)
            results['drawing_pages'].extend(drawing_pages)
            
        except Exception as e:
            print(f"❌ Ошибка извлечения страниц: {e}")
    
    # Count total drawings
    total_drawings = len([f for f in os.listdir(drawings_folder) if f.lower().endswith('.pdf')])
    session_info['drawings_count'] = total_drawings
    
    # Step 4: Process IFC files - prepare viewer and parse
    print("\n" + "="*60)
    print("🏢 ШАГ 4: ОБРАБОТКА IFC МОДЕЛЕЙ")
    print("="*60)
    
    ifc_files = [os.path.join(ifc_folder, f) for f in os.listdir(ifc_folder) if f.lower().endswith('.ifc')]
    ifc_viewer_info = []
    
    for ifc_path in ifc_files:
        if os.path.exists(ifc_path):
            # Step 4a: Prepare IFC for viewer
            print(f"\n📐 Подготовка IFC для вьюера: {os.path.basename(ifc_path)}")
            try:
                viewer_result = prepare_ifc_for_viewer(ifc_path, session_folder)
                ifc_viewer_info.append(viewer_result)
                session_info['ifc_viewer_info'] = ifc_viewer_info
                print(f"✓ IFC готов для просмотра: {viewer_result.get('filename', 'unknown')}")
            except Exception as e:
                print(f"❌ Ошибка подготовки IFC вьюера: {e}")
                session_info['ifc_viewer_error'] = str(e)
            
            # Step 4b: Parse IFC file
            print(f"\n📊 Парсинг IFC: {os.path.basename(ifc_path)}")
            try:
                ifc_result = parse_ifc_file(ifc_path, session_folder)
                results['ifc_results'] = ifc_result
                session_info['ifc_excel_file'] = ifc_result.get('excel_filename')
                
                # Create materials summary after IFC report
                if ifc_result.get('excel_filename'):
                    excel_path = os.path.join(session_folder, ifc_result['excel_filename'])
                    if os.path.exists(excel_path):
                        print(f"\n📈 Создание сводной таблицы по материалам: {excel_path}")
                        materials_excel_path = create_materials_summary(excel_path)
                        session_info['materials_excel_file'] = os.path.basename(materials_excel_path)
                        results['materials_summary'] = {
                            'filename': os.path.basename(materials_excel_path),
                            'path': str(materials_excel_path)
                        }
                        print(f"✓ Сводная таблица сохранена: {session_info['materials_excel_file']}")
            except Exception as e:
                print(f"❌ Ошибка парсинга IFC: {e}")
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
        'drawings_folder_count': total_drawings,
        'ifc_processed': results['ifc_results'] is not None,
        'pdf_classified': results['pdf_classification'] is not None,
        'materials_summary_created': results['materials_summary'] is not None
    }
    
    with open(session_info_path, 'w', encoding='utf-8') as f:
        json.dump(session_info, f, ensure_ascii=False, indent=2)
    
    return jsonify({
        'success': True,
        'session_id': session_id,
        'text_pages_count': len(results['text_pages']),
        'drawing_pages_count': len(results['drawing_pages']),
        'drawings_count': total_drawings,
        'ifc_processed': results['ifc_results'] is not None,
        'ifc_excel_file': session_info.get('ifc_excel_file'),
        'materials_excel_file': session_info.get('materials_excel_file'),
        'pdf_classification': results['pdf_classification'],
        'summary': session_info['results_summary']
    })


@app.route('/api/viewer/<session_id>/<path:filename>')
def serve_ifc_viewer_file(session_id, filename):
    """Serve IFC or geometry JSON file for the web viewer"""
    session_folder = os.path.join(UPLOAD_FOLDER, session_id)
    viewer_folder = os.path.join(session_folder, 'viewer')
    
    if not os.path.exists(viewer_folder):
        return jsonify({'error': 'Viewer folder not found'}), 404
    
    file_path = os.path.join(viewer_folder, filename)
    if not os.path.exists(file_path):
        return jsonify({'error': 'File not found'}), 404
    
    # Determine MIME type based on file extension
    if filename.lower().endswith('.json'):
        mimetype = 'application/json'
    elif filename.lower().endswith('.ifc'):
        mimetype = 'application/x-step'
    else:
        mimetype = 'application/octet-stream'
    
    response = send_from_directory(viewer_folder, filename, mimetype=mimetype)
    response.headers['Content-Type'] = mimetype
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
