#!/usr/bin/env python3
"""
Скрипт извлечения страниц с чертежами из PDF файлов пояснительной записки.
Использует логику определения чертежей по размеру страницы (> 30см).
Страницы <= 30см считаются текстовыми и остаются в пояснительной записке.
Все чертежи сохраняются в отдельную папку drawings/.
"""

import os
import fitz  # PyMuPDF
from pathlib import Path
from typing import List, Dict, Any, Tuple
import shutil


# Пороговый размер для определения чертежа (большая сторона в см)
# A4 = 29.7см, A3 = 42см, A2 = 59.4см
# Устанавливаем порог 30см - всё что больше A4 считается чертежом
DRAWING_MIN_SIZE_CM = 30.0


def correct_page_orientation(page: fitz.Page) -> int:
    """
    Определяет и корректирует ориентацию страницы.
    Все чертежи должны быть в альбомной ориентации (ширина > высоты).
    
    Логика:
    1. Получаем физические размеры страницы в см.
    2. Учитываем page.rotation для определения фактической видимой ориентации.
    3. Если видимая ширина < видимой высоты, значит страница портретная -> нужно повернуть на 90°.
    4. Возвращаем угол поворота для insert_pdf(rotate=...), чтобы получить итоговую альбомную ориентацию.
    """
    # Получаем текущий rotation страницы из PDF метаданных
    current_rotation = page.rotation
    
    # Получаем "сырые" размеры страницы в пунктах (без учета rotation)
    raw_width_pt = page.rect.width
    raw_height_pt = page.rect.height
    
    # Вычисляем фактические видимые размеры с учетом текущего rotation
    if current_rotation in [90, 270]:
        visible_width_pt = raw_height_pt
        visible_height_pt = raw_width_pt
    else:
        visible_width_pt = raw_width_pt
        visible_height_pt = raw_height_pt
    
    # Конвертируем в см для наглядности (1 pt = 1/72 inch, 1 inch = 2.54 cm)
    visible_width_cm = visible_width_pt * 2.54 / 72
    visible_height_cm = visible_height_pt * 2.54 / 72
    
    print(f"      📏 Размеры: {visible_width_cm:.1f}x{visible_height_cm:.1f}см (rotation={current_rotation}°)")
    
    # Определяем необходимость поворота
    rotation_adjustment = 0
    
    if visible_width_cm < visible_height_cm:
        # Страница фактически портретная, нужно повернуть на 90° по часовой стрелке
        rotation_adjustment = 90
        print(f"      📐 Страница портретная, требуется поворот на 90°")
    else:
        # Страница уже альбомная
        print(f"      ✅ Страница уже в альбомной ориентации")
        
        if current_rotation != 0:
            rotation_adjustment = (360 - current_rotation) % 360
            print(f"      🔄 Компенсация встроенного rotation {current_rotation}° -> {rotation_adjustment}°")
        else:
            rotation_adjustment = 0
            
    return rotation_adjustment


def extract_drawings_from_pdf(pdf_path: str, session_folder: str) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """
    Извлекает страницы из PDF как отдельные PDF файлы.
    Разделяет на текстовые страницы (<= 30см) и страницы с чертежами (> 30см).
    Чертежи сохраняются в папку drawings/, текстовые страницы остаются в explanatory_note/.
    Ориентация корректируется для альбомного формата.
    
    Args:
        pdf_path: Путь к PDF файлу
        session_folder: Папка сессии
        
    Returns:
        Кортеж из двух списков: (текстовые_страницы, страницы_с_чертежами)
    """
    text_pages = []
    drawing_pages = []
    
    try:
        doc = fitz.open(pdf_path)
        filename = os.path.basename(pdf_path)
        
        # Создаем папки для страниц
        drawings_dir = Path(session_folder) / "drawings"
        drawings_dir.mkdir(parents=True, exist_ok=True)
        
        print(f"🔍 Анализ файла {filename}... Критерий чертежа: > {DRAWING_MIN_SIZE_CM} см")
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            
            # Получаем размеры в см
            w_cm = page.rect.width * 2.54 / 72
            h_cm = page.rect.height * 2.54 / 72
            max_side = max(w_cm, h_cm)
            
            # Определяем необходимую коррекцию ориентации
            rotation = correct_page_orientation(page)
            
            # Создаем новый PDF с одной страницей
            new_doc = fitz.open()
            new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num, rotate=rotation)
            
            page_info = {
                'page_num': page_num + 1,
                'source_file': filename,
                'size': f"{w_cm:.1f}x{h_cm:.1f}cm",
                'width_cm': w_cm,
                'height_cm': h_cm
            }
            
            # Классифицируем страницу и сохраняем в соответствующую папку
            if max_side <= DRAWING_MIN_SIZE_CM:
                # Текстовая страница - остается в explanatory_note
                page_filename = f"{filename}_text_page_{page_num + 1}.pdf"
                page_path = Path(session_folder) / "explanatory_note" / page_filename
                new_doc.save(str(page_path))
                
                page_info['path'] = str(page_path.absolute())
                page_info['type'] = 'text'
                text_pages.append(page_info)
                print(f"   📄 Стр. {page_num + 1}: Текстовая ({w_cm:.1f}x{h_cm:.1f}см) -> explanatory_note/")
            else:
                # Страница с чертежом - сохраняем в drawings/
                page_filename = f"{filename}_drawing_page_{page_num + 1}.pdf"
                page_path = drawings_dir / page_filename
                new_doc.save(str(page_path))
                
                page_info['path'] = str(page_path.absolute())
                page_info['type'] = 'drawing'
                drawing_pages.append(page_info)
                print(f"   🏗️ Стр. {page_num + 1}: Чертеж ({w_cm:.1f}x{h_cm:.1f}см) -> drawings/")
            
            new_doc.close()
        
        doc.close()
        
        print(f"✅ Найдено {len(text_pages)} текстовых страниц и {len(drawing_pages)} страниц с чертежами")
        
        return text_pages, drawing_pages
        
    except Exception as e:
        print(f"❌ Ошибка при извлечении страниц из {pdf_path}: {e}")
        import traceback
        traceback.print_exc()
        return [], []


def extract_drawings_from_explanatory_note(session_folder: str) -> Dict[str, Any]:
    """
    Обрабатывает все PDF файлы в папке explanatory_note и извлекает чертежи.
    
    Args:
        session_folder: Путь к папке сессии
        
    Returns:
        Словарь с результатами обработки
    """
    
    results = {
        'text_pages': [],
        'drawing_pages': [],
        'processed_files': [],
        'total_drawings': 0
    }
    
    print("\n" + "="*60)
    print("🏗️ ШАГ 3: ИЗВЛЕЧЕНИЕ ЧЕРТЕЖЕЙ ИЗ ПОЯСНИТЕЛЬНОЙ ЗАПИСКИ")
    print("="*60)
    
    # Находим все PDF файлы в папке explanatory_note
    explanatory_folder = Path(session_folder) / 'explanatory_note'
    
    if not explanatory_folder.exists():
        print(f"⚠️  Папка explanatory_note не найдена: {explanatory_folder}")
        return results
    
    pdf_files = [f for f in explanatory_folder.glob("*.pdf") if f.is_file()]
    
    print(f"🔍 Найдено PDF файлов в explanatory_note: {len(pdf_files)}")
    
    for pdf_path in pdf_files:
        filename = pdf_path.name
        print(f"\n📄 Обработка файла: {filename}")
        
        try:
            text_pages, drawing_pages = extract_drawings_from_pdf(str(pdf_path), session_folder)
            
            results['text_pages'].extend(text_pages)
            results['drawing_pages'].extend(drawing_pages)
            results['processed_files'].append({
                'filename': filename,
                'text_pages_count': len(text_pages),
                'drawing_pages_count': len(drawing_pages)
            })
            
        except Exception as e:
            print(f"❌ Ошибка обработки {filename}: {e}")
    
    results['total_drawings'] = len(results['drawing_pages'])
    
    # Сохраняем результаты
    drawings_results_path = Path(session_folder) / "drawings_results.json"
    with open(drawings_results_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2, default=str)
    
    print(f"\n💾 Результаты сохранены в: {drawings_results_path}")
    
    # Выводим сводку
    print("\n" + "="*60)
    print("📊 СВОДКА ПО ИЗВЛЕЧЕНИЮ ЧЕРТЕЖЕЙ:")
    print("="*60)
    print(f"📄 Текстовых страниц: {len(results['text_pages'])}")
    print(f"🏗️ Страниц с чертежами: {len(results['drawing_pages'])}")
    print(f"📁 Обработано файлов: {len(results['processed_files'])}")
    
    for file_info in results['processed_files']:
        print(f"   • {file_info['filename']}: {file_info['text_pages_count']} текст., {file_info['drawing_pages_count']} чертежей")
    
    return results


def main(session_folder: str):
    """
    Основная функция извлечения чертежей из пояснительной записки.
    
    Args:
        session_folder: Путь к папке сессии
    """
    import json
    
    print("="*60)
    print("🏗️ ИЗВЛЕЧЕНИЕ ЧЕРТЕЖЕЙ ИЗ ПОЯСНИТЕЛЬНОЙ ЗАПИСКИ")
    print("="*60)
    print(f"Папка сессии: {session_folder}")
    print(f"Порог размера чертежа: {DRAWING_MIN_SIZE_CM} см")
    print("="*60)
    
    results = extract_drawings_from_explanatory_note(session_folder)
    
    return results


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Использование: python draw_detector.py <session_folder>")
        print("Пример: python draw_detector.py /path/to/uploads/session_id")
        sys.exit(1)
    
    session_folder = sys.argv[1]
    main(session_folder)
