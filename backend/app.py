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
from scripts.ifc_parser import parse_ifc_file
from scripts.filter_works_by_height import filter_works_by_height
from scripts.map_elements_to_works import map_elements_to_works

# xlsx_parser module was removed - IFC parser now handles materials summary creation

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
        'elements_json': None,
        'elements_json_results': None
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
                # Step 4: Parse IFC file
                print(f"\n📊 Парсинг IFC: {os.path.basename(ifc_path)}")
                try:
                    ifc_result = parse_ifc_file(ifc_path, session_folder)
                    results['ifc_results'] = ifc_result
                    session_info['ifc_excel_file'] = ifc_result.get('excel_filename')
                except Exception as e:
                    print(f"❌ Ошибка парсинга IFC: {e}")
                    session_info['ifc_error'] = str(e)
    
    # Step 5: Check if elements JSON was created by IFC parser
    print("\n" + "="*60)
    print("📊 ШАГ 5: ПРОВЕРКА СПИСКА ЭЛЕМЕНТОВ (IFC)")
    print("="*60)
    
    try:
        elements_json_file = os.path.join(session_folder, 'elements.json')
        if os.path.exists(elements_json_file):
            # IFC parser already created the elements JSON
            session_info['elements_json_file'] = 'elements.json'
            results['elements_json_results'] = {
                'success': True,
                'source': 'ifc_parser',
                'output_file': 'elements.json'
            }
            print(f"✅ Список элементов найден: {session_info['elements_json_file']}")
        else:
            print("⚠️ Файл elements.json не найден")
            results['elements_json_results'] = {'success': False, 'error': 'Elements JSON not found'}
    except Exception as e:
        print(f"❌ Ошибка проверки списка элементов: {e}")
        session_info['elements_json_error'] = str(e)
        results['elements_json_results'] = {'success': False, 'error': str(e)}
    
    # Step 6: Filter works list by building height
    print("\n" + "="*60)
    print("📐 ШАГ 6: ФИЛЬТРАЦИЯ ПЕРЕЧНЯ РАБОТ ПО ВЫСОТЕ ЗДАНИЯ")
    print("="*60)
    
    try:
        height_filter_result = filter_works_by_height(session_folder)
        if height_filter_result.get('success'):
            results['height_filter'] = height_filter_result
            session_info['works_list_filtered'] = True
            session_info['works_list_output_file'] = height_filter_result.get('output_file')
            print(f"✅ Перечень работ отфильтрован: {height_filter_result.get('output_file')}")
        else:
            print(f"⚠️ Фильтрация не выполнена: {height_filter_result.get('error', 'Unknown error')}")
            session_info['height_filter_error'] = height_filter_result.get('error', 'Unknown error')
    except Exception as e:
        print(f"❌ Ошибка фильтрации перечня работ: {e}")
        session_info['height_filter_error'] = str(e)

    # Step 7: Map elements to works
    print("\n" + "="*60)
    print("🔗 ШАГ 7: МАППИНГ ЭЛЕМЕНТОВ К РАБОТАМ")
    print("="*60)
    
    try:
        # Проверяем наличие необходимых файлов перед запуском маппинга
        elements_excel_file = os.path.join(session_folder, 'full_elements.xlsx')
        works_filtered_file = os.path.join(session_folder, 'Перечень работ КР_new.xlsx')
        
        if not os.path.exists(elements_excel_file):
            print(f"⚠️ Файл full_elements.xlsx не найден: {elements_excel_file}")
            session_info['mapping_error'] = 'Файл full_elements.xlsx не найден'
            results['mapping'] = {'success': False, 'error': 'Elements Excel file not found'}
        elif not os.path.exists(works_filtered_file):
            print(f"⚠️ Файл Перечень работ КР_new.xlsx не найден: {works_filtered_file}")
            session_info['mapping_error'] = 'Файл работ не найден'
            results['mapping'] = {'success': False, 'error': 'Works file not found'}
        else:
            mapping_result = map_elements_to_works(session_folder)
            if mapping_result.get('success'):
                results['mapping'] = mapping_result
                session_info['mapping_completed'] = True
                session_info['mapped_elements_works_file'] = mapping_result.get('output_file')
                print(f"✅ Маппинг завершен: {mapping_result.get('output_file')}")
                print(f"   Смаппировано {mapping_result.get('matched_elements', 0)} из {mapping_result.get('total_elements', 0)} элементов")
            else:
                print(f"⚠️ Маппинг не выполнен: {mapping_result.get('error', 'Unknown error')}")
                session_info['mapping_error'] = mapping_result.get('error', 'Unknown error')
    except Exception as e:
        print(f"❌ Ошибка маппинга: {e}")
        session_info['mapping_error'] = str(e)

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
        'elements_json_saved': results['elements_json_results'] is not None and results['elements_json_results'].get('success', False),
        'mapping_completed': results.get('mapping') is not None and results['mapping'].get('success', False)
    }
    
    with open(session_info_path, 'w', encoding='utf-8') as f:
        json.dump(session_info, f, ensure_ascii=False, indent=2)
    
    return jsonify({
        'success': True,
        'session_id': session_id,
        'text_pages_count': text_pages_count,
        'drawing_pages_count': drawing_pages_count,
        'drawings_count': session_info.get('drawings_count', 0),
        'ifc_processed': results['ifc_results'] is not None,
        'ifc_excel_file': session_info.get('ifc_excel_file'),
        'elements_json_file': session_info.get('elements_json_file'),
        'mapped_elements_works_file': session_info.get('mapped_elements_works_file'),
        'pdf_classification': results['pdf_classification'],
        'elements_json_results': results['elements_json_results'],
        'mapping_results': results.get('mapping'),
        'summary': session_info['results_summary']
    })


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


@app.route('/api/elements/<session_id>')
def get_elements(session_id):
    """Get elements data from JSON file"""
    
    session_folder = os.path.join(UPLOAD_FOLDER, session_id)
    if not os.path.exists(session_folder):
        return jsonify({'error': 'Session not found'}), 404
    
    elements_file = os.path.join(session_folder, 'elements.json')
    if not os.path.exists(elements_file):
        return jsonify({'error': 'Elements JSON not found'}), 404
    
    try:
        with open(elements_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return jsonify(data)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/mapped-elements/<session_id>')
def get_mapped_elements(session_id):
    """Get mapped elements works data from Excel file"""
    import openpyxl
    
    session_folder = os.path.join(UPLOAD_FOLDER, session_id)
    if not os.path.exists(session_folder):
        return jsonify({'error': 'Session not found'}), 404
    
    mapped_file = os.path.join(session_folder, 'mapped_elements_works.xlsx')
    if not os.path.exists(mapped_file):
        return jsonify({'error': 'Mapped elements file not found'}), 404
    
    try:
        wb = openpyxl.load_workbook(mapped_file, data_only=True)
        ws = wb.active
        rows = []
        for row in ws.iter_rows(values_only=True):
            rows.append([str(cell) if cell is not None else '' for cell in row])
        
        return jsonify({'data': rows})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=6003)