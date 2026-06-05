from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import os
import sys
import uuid
import json
import shutil

# Add parent directory to path for scripts imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from scripts.files_classifier import classify_and_organize_files
from scripts.pdf_classifier_llm import classify_pdf_files as classify_pdf_llm
from scripts.draw_detector import extract_drawings_from_explanatory_note
from scripts.ifc_viewer import prepare_ifc_for_viewer
from scripts.ifc_parser import parse_ifc_file
from scripts.volume_statement_parser import parse_volume_statement_pdfs

# xlsx_parser module was removed - IFC parser now handles materials summary creation
# parse_and_aggregate_specification is no longer needed since ifc_parser.create_summary_table does this
# Stub function kept for backward compatibility but not used in process_files anymore
def parse_and_aggregate_specification(session_folder):
    """Stub function - materials summary is now created by ifc_parser"""
    materials_file = os.path.join(session_folder, 'materials_summary.xlsx')
    if os.path.exists(materials_file):
        return {
            'success': True,
            'source': 'ifc_parser',
            'output_file': 'materials_summary.xlsx',
            'aggregated_materials': 0  # Will be populated from results.json
        }
    return {
        'success': False,
        'error': 'xlsx_parser module was removed - materials summary is created from IFC',
        'aggregated_materials': 0
    }

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
        'materials_summary': None,
        'xlsx_parser_results': None
    }
    
    # Step 1: Organize files into folders (IFC, Excel, PDF)
    print("\n" + "="*60)
    print("📁 ШАГ 1: ОРГАНИЗАЦИЯ ФАЙЛОВ ПО ТИПАМ")
    print("="*60)
    
    file_org_results = classify_and_organize_files(session_folder)
    session_info['file_organization'] = {
        'ifc_count': len(file_org_results.get('ifc_models', [])),
        'excel_count': len(file_org_results.get('specification', [])),
        'pdf_count': len(file_org_results.get('pdf_documents', []))
    }
    
    # Step 2: Classify PDF files using LLM (gemma3:27b, first 3 pages only)
    print("\n" + "="*60)
    print("🤖 ШАГ 2: КЛАССИФИКАЦИЯ PDF ДОКУМЕНТОВ (LLM)")
    print("="*60)
    
    try:
        pdf_classification = classify_pdf_llm(session_folder)
        results['pdf_classification'] = {
            'categories': {cat: len(files) for cat, files in pdf_classification.items()},
            'files': pdf_classification
        }
        session_info['pdf_classification'] = results['pdf_classification']
        print("✅ Классификация PDF завершена")
    except Exception as e:
        print(f"❌ Ошибка классификации PDF: {e}")
        session_info['pdf_classification_error'] = str(e)
    
    # Step 3: Extract drawings from "explanatory_note" PDFs
    print("\n" + "="*60)
    print("🏗️ ШАГ 3: ИЗВЛЕЧЕНИЕ ЧЕРТЕЖЕЙ ИЗ ПОЯСНИТЕЛЬНОЙ ЗАПИСКИ")
    print("="*60)
    
    try:
        drawings_results = extract_drawings_from_explanatory_note(session_folder)
        results['text_pages'] = drawings_results.get('text_pages', [])
        results['drawing_pages'] = drawings_results.get('drawing_pages', [])
        session_info['drawings_count'] = drawings_results.get('total_drawings', 0)
        print(f"✅ Извлечено {drawings_results.get('total_drawings', 0)} чертежей")
    except Exception as e:
        print(f"❌ Ошибка извлечения чертежей: {e}")
        session_info['drawings_error'] = str(e)
    
    # Step 4: Process IFC files - prepare viewer and parse
    print("\n" + "="*60)
    print("🏢 ШАГ 4: ОБРАБОТКА IFC МОДЕЛЕЙ")
    print("="*60)
    
    ifc_folder = os.path.join(session_folder, 'ifc_models')
    ifc_viewer_info = []
    
    if os.path.exists(ifc_folder):
        ifc_files = [os.path.join(ifc_folder, f) for f in os.listdir(ifc_folder) if f.lower().endswith('.ifc')]
        
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
                except Exception as e:
                    print(f"❌ Ошибка парсинга IFC: {e}")
                    session_info['ifc_error'] = str(e)
    
    # Step 5: Check if materials summary was created by IFC parser
    print("\n" + "="*60)
    print("📊 ШАГ 5: ПРОВЕРКА СВODНОЙ ТАБЛИЦЫ МАТЕРИАЛОВ (IFC)")
    print("="*60)
    
    try:
        materials_file = os.path.join(session_folder, 'materials_summary.xlsx')
        if os.path.exists(materials_file):
            # IFC parser already created the materials summary
            session_info['materials_excel_file'] = 'materials_summary.xlsx'
            results['xlsx_parser_results'] = {
                'success': True,
                'source': 'ifc_parser',
                'output_file': 'materials_summary.xlsx'
            }
            print(f"✅ Сводная таблица материалов найдена: {session_info['materials_excel_file']}")
        else:
            print("⚠️ Файл materials_summary.xlsx не найден")
            results['xlsx_parser_results'] = {'success': False, 'error': 'Materials summary not found'}
    except Exception as e:
        print(f"❌ Ошибка проверки сводной таблицы: {e}")
        session_info['xlsx_parser_error'] = str(e)
        results['xlsx_parser_results'] = {'success': False, 'error': str(e)}
    
    # Step 6: Parse volume statement PDFs and extract material tables
    print("\n" + "="*60)
    print("📊 ШАГ 6: ПАРСИНГ ВЕДОМОСТЕЙ ОБЪЕМОВ (PDF)")
    print("="*60)

    try:
        volume_results = parse_volume_statement_pdfs(session_folder)
        results['volume_statement_results'] = volume_results
        if volume_results.get('success'):
            session_info['volume_statement_excel_file'] = 'volume_statement_summary.xlsx'
            print(f"✅ Ведомость объемов из PDF создана: {session_info['volume_statement_excel_file']}")
    except Exception as e:
        print(f"❌ Ошибка парсинга ведомостей объемов: {e}")
        session_info['volume_statement_error'] = str(e)

    # Save results
    results_path = os.path.join(session_folder, 'results.json')
    with open(results_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2, default=str)
    
    # Update session status - count files in explanatory_note and drawings folders
    session_info['status'] = 'completed'
    
    # Count files in explanatory_note folder (text pages)
    explanatory_folder = os.path.join(session_folder, 'explanatory_note')
    text_pages_count = 0
    if os.path.exists(explanatory_folder):
        text_pages_count = len([f for f in os.listdir(explanatory_folder) if f.lower().endswith('.pdf')])
    
    # Count files in drawings folder
    drawings_folder = os.path.join(session_folder, 'drawings')
    drawing_pages_count = 0
    if os.path.exists(drawings_folder):
        drawing_pages_count = len([f for f in os.listdir(drawings_folder) if f.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg'))])
    
    session_info['results_summary'] = {
        'text_pages_count': text_pages_count,
        'drawing_pages_count': drawing_pages_count,
        'drawings_folder_count': session_info.get('drawings_count', 0),
        'ifc_processed': results['ifc_results'] is not None,
        'pdf_classified': results['pdf_classification'] is not None,
        'xlsx_parsed': results['xlsx_parser_results'] is not None and results['xlsx_parser_results'].get('success', False)
    }
    
    with open(session_info_path, 'w', encoding='utf-8') as f:
        json.dump(session_info, f, ensure_ascii=False, indent=2)
    
    # Get aggregated_materials count from IFC results if available
    aggregated_materials = 0
    if results.get('ifc_results') and results['ifc_results'].get('aggregated_materials'):
        aggregated_materials = results['ifc_results']['aggregated_materials']
    
    return jsonify({
        'success': True,
        'session_id': session_id,
        'text_pages_count': text_pages_count,
        'drawing_pages_count': drawing_pages_count,
        'drawings_count': session_info.get('drawings_count', 0),
        'ifc_processed': results['ifc_results'] is not None,
        'ifc_excel_file': session_info.get('ifc_excel_file'),
        'materials_excel_file': session_info.get('materials_excel_file'),
        'pdf_classification': results['pdf_classification'],
        'xlsx_parser_results': results['xlsx_parser_results'],
        'ifc_results': results.get('ifc_results'),  # Add IFC results for materials count
        'aggregated_materials': aggregated_materials,
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


@app.route('/api/materials-summary/<session_id>')
def get_materials_summary(session_id):
    """Get materials summary data from Excel file"""
    import openpyxl
    
    session_folder = os.path.join(UPLOAD_FOLDER, session_id)
    if not os.path.exists(session_folder):
        return jsonify({'error': 'Session not found'}), 404
    
    materials_file = os.path.join(session_folder, 'materials_summary.xlsx')
    if not os.path.exists(materials_file):
        return jsonify({'error': 'Materials summary not found'}), 404
    
    try:
        wb = openpyxl.load_workbook(materials_file, data_only=True)
        ws = wb.active
        rows = []
        for row in ws.iter_rows(values_only=True):
            rows.append([str(cell) if cell is not None else '' for cell in row])
        
        return jsonify({'data': rows})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/building-height/<session_id>')
def get_building_height(session_id):
    """Get building height from building_info.json file"""
    
    session_folder = os.path.join(UPLOAD_FOLDER, session_id)
    if not os.path.exists(session_folder):
        return jsonify({'error': 'Session not found'}), 404
    
    building_info_file = os.path.join(session_folder, 'building_info.json')
    if not os.path.exists(building_info_file):
        return jsonify({'error': 'Building info not found'}), 404
    
    try:
        with open(building_info_file, 'r', encoding='utf-8') as f:
            building_info = json.load(f)
        
        # Get height from summary.total_height_from_zero_m
        height = building_info.get('summary', {}).get('total_height_from_zero_m')
        
        if height is not None:
            return jsonify({'height': height})
        else:
            return jsonify({'height': None})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/building-address/<session_id>')
def get_building_address(session_id):
    """Get building address from building_info.json file"""
    
    session_folder = os.path.join(UPLOAD_FOLDER, session_id)
    if not os.path.exists(session_folder):
        return jsonify({'error': 'Session not found'}), 404
    
    building_info_file = os.path.join(session_folder, 'building_info.json')
    if not os.path.exists(building_info_file):
        return jsonify({'error': 'Building info not found'}), 404
    
    try:
        with open(building_info_file, 'r', encoding='utf-8') as f:
            building_info = json.load(f)
        
        # Get address from summary.address
        address = building_info.get('summary', {}).get('address')
        
        if address is not None:
            return jsonify({'address': address})
        else:
            return jsonify({'address': None})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=6003)