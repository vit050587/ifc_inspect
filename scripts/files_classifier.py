#!/usr/bin/env python3
"""
Скрипт классификации файлов по типам и распределения по папкам.
- IFC файлы -> ifc_models/
- Excel файлы -> specification/
- PDF файлы -> pdf_documents/
"""

import os
import shutil
from pathlib import Path
from typing import Dict, List, Any


def classify_and_organize_files(session_folder: str) -> Dict[str, List[Dict[str, Any]]]:
    """
    Распределяет файлы из корня сессии по папкам в зависимости от типа.
    
    Args:
        session_folder: Путь к папке сессии
        
    Returns:
        Словарь с результатами: {тип_файла: [список_файлов]}
    """
    
    results = {
        'ifc_models': [],
        'specification': [],
        'pdf_documents': []
    }
    
    # Создаем целевые папки
    ifc_folder = Path(session_folder) / 'ifc_models'
    excel_folder = Path(session_folder) / 'specification'
    pdf_folder = Path(session_folder) / 'pdf_documents'
    
    ifc_folder.mkdir(parents=True, exist_ok=True)
    excel_folder.mkdir(parents=True, exist_ok=True)
    pdf_folder.mkdir(parents=True, exist_ok=True)
    
    print("\n" + "="*60)
    print("📁 ШАГ 1: ОРГАНИЗАЦИЯ ФАЙЛОВ ПО ТИПАМ")
    print("="*60)
    
    # Находим все файлы в корне сессии (не в подпапках)
    session_path = Path(session_folder)
    files = [f for f in session_path.iterdir() if f.is_file()]
    
    print(f"🔍 Найдено файлов для обработки: {len(files)}")
    
    for file_path in files:
        filename = file_path.name
        suffix = file_path.suffix.lower()
        
        try:
            if suffix == '.ifc':
                # IFC файл -> ifc_models/
                dest_path = ifc_folder / filename
                if file_path != dest_path:
                    shutil.move(str(file_path), str(dest_path))
                    print(f"   🏗️ IFC: {filename} -> ifc_models/")
                
                results['ifc_models'].append({
                    'filename': filename,
                    'source_path': str(file_path.parent / filename),
                    'dest_path': str(dest_path),
                    'file_size': dest_path.stat().st_size
                })
                
            elif suffix in ['.xlsx', '.xls']:
                # Excel файл -> specification/
                dest_path = excel_folder / filename
                if file_path != dest_path:
                    shutil.move(str(file_path), str(dest_path))
                    print(f"   📊 Excel: {filename} -> specification/")
                
                results['specification'].append({
                    'filename': filename,
                    'source_path': str(file_path.parent / filename),
                    'dest_path': str(dest_path),
                    'file_size': dest_path.stat().st_size
                })
                
            elif suffix == '.pdf':
                # PDF файл -> pdf_documents/
                dest_path = pdf_folder / filename
                if file_path != dest_path:
                    shutil.move(str(file_path), str(dest_path))
                    print(f"   📄 PDF: {filename} -> pdf_documents/")
                
                results['pdf_documents'].append({
                    'filename': filename,
                    'source_path': str(file_path.parent / filename),
                    'dest_path': str(dest_path),
                    'file_size': dest_path.stat().st_size
                })
                
            else:
                # Другие файлы - оставляем в корне или логируем
                print(f"   ⚠️  Неизвестный тип файла: {filename} (оставлен в корне)")
                
        except Exception as e:
            print(f"   ❌ Ошибка при перемещении {filename}: {e}")
    
    # Выводим сводку
    print("\n" + "="*60)
    print("📊 СВОДКА ПО РАСПРЕДЕЛЕНИЮ ФАЙЛОВ:")
    print("="*60)
    print(f"🏗️ IFC модели: {len(results['ifc_models'])} файл(ов)")
    for f in results['ifc_models']:
        print(f"   • {f['filename']} ({f['file_size'] / 1024 / 1024:.1f} MB)")
    
    print(f"📊 Excel спецификации: {len(results['specification'])} файл(ов)")
    for f in results['specification']:
        print(f"   • {f['filename']} ({f['file_size'] / 1024 / 1024:.1f} MB)")
    
    print(f"📄 PDF документы: {len(results['pdf_documents'])} файл(ов)")
    for f in results['pdf_documents']:
        print(f"   • {f['filename']} ({f['file_size'] / 1024 / 1024:.1f} MB)")
    
    return results


def main(session_folder: str):
    """
    Основная функция классификации и организации файлов.
    
    Args:
        session_folder: Путь к папке сессии
    """
    print("="*60)
    print("📁 КЛАССИФИКАЦИЯ И ОРГАНИЗАЦИЯ ФАЙЛОВ")
    print("="*60)
    print(f"Папка сессии: {session_folder}")
    print("="*60)
    
    results = classify_and_organize_files(session_folder)
    
    return results


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Использование: python files_classifier.py <session_folder>")
        print("Пример: python files_classifier.py /path/to/uploads/session_id")
        sys.exit(1)
    
    session_folder = sys.argv[1]
    main(session_folder)
