#!/usr/bin/env python3
"""
Скрипт для маппинга таблицы элементов (full_elements.xlsx) с таблицей перечня работ (Перечень работ КР_new.xlsx).

Логика работы:
1. Читает full_elements.xlsx из папки сессии - таблица с элементами IFC модели
2. Читает Перечень работ КР_new.xlsx из папки сессии - отфильтрованный перечень работ
3. Сопоставляет элементы с работами по параметрам:
   - В элементах параметры могут быть в наименовании и других столбцах
   - В новой таблице параметры указаны в столбцах B (Наименование работ), E (Формула расчёта), F (Параметризация)
4. К одной строке элемента может быть смаплено несколько строк из перечня работ
5. Сохраняет результирующую таблицу в папке сессии
"""

import os
import re
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import pandas as pd


def extract_parameters_from_element(element_row):
    """
    Извлекает параметры из строки элемента IFC.
    
    Args:
        element_row: dict со всеми колонками элемента
        
    Returns:
        dict с извлеченными параметрами
    """
    params = {
        'name': '',
        'category': '',
        'material': '',
        'concrete_class': '',
        'frost': '',
        'water': '',
        'width': None,
        'height': None,
        'length': None,
        'thickness': None,
        'volume': None,
        'area': None,
        'underground': None
    }
    
    # Наименование
    name = element_row.get('Наименование', '') or ''
    params['name'] = str(name).upper() if name else ''
    
    # Категория (из classify_element)
    category = element_row.get('category', '') or ''
    params['category'] = str(category).upper() if category else ''
    
    # Материал
    material = element_row.get('Материал', '') or ''
    params['material'] = str(material).upper() if material else ''
    
    # Характеристики материала
    char_str = element_row.get('Характеристики материала', '') or ''
    if char_str:
        # Класс бетона (B30, B25, B35, B7.5 и т.д.)
        class_match = re.search(r'[BВ]\s?(\d+(?:[.,]\d+)?)', char_str.upper())
        if class_match:
            params['concrete_class'] = class_match.group(1).replace(',', '.')
        
        # Морозостойкость (F150, F100 и т.д.)
        frost_match = re.search(r'F(\d+)', char_str.upper())
        if frost_match:
            params['frost'] = frost_match.group(1)
        
        # Водонепроницаемость (W6, W8 и т.д.)
        water_match = re.search(r'W(\d+)', char_str.upper())
        if water_match:
            params['water'] = water_match.group(1)
    
    # Размеры
    params['width'] = element_row.get('Ширина, м')
    params['height'] = element_row.get('Высота, м')
    params['length'] = element_row.get('Длина, м')
    params['thickness'] = element_row.get('Толщина, м')
    
    # Объем и площадь
    params['volume'] = element_row.get('Объем, м³')
    params['area'] = element_row.get('Площадь, м²')
    
    # Подземный/Надземный
    underground_status = element_row.get('Подземный/Надземный', '')
    if underground_status:
        params['underground'] = 'ПОДЗЕМ' in str(underground_status).upper()
    
    return params


def extract_parameters_from_work(work_row):
    """
    Извлекает параметры из строки перечня работ.
    
    Args:
        work_row: dict со столбцами B, E, F
        
    Returns:
        dict с извлеченными параметрами
    """
    params = {
        'work_name': '',
        'formula': '',
        'parameters': '',
        'concrete_class': '',
        'thickness_min': None,
        'thickness_max': None,
        'height_min': None,
        'height_max': None,
        'area_min': None,
        'area_max': None,
        'perimeter_min': None,
        'perimeter_max': None,
        'diameter_min': None,
        'diameter_max': None,
        'is_underground': None  # True/False если указано
    }
    
    # Столбец B - Наименование работ
    work_name = work_row.get('Наименование работ', '') or ''
    params['work_name'] = str(work_name).upper() if work_name else ''
    
    # Столбец E - Формула расчёта
    formula = work_row.get('Формула расчёта объёмов работ и расхода материалов', '') or ''
    params['formula'] = str(formula).upper() if formula else ''
    
    # Столбец F - Параметризация
    parameters = work_row.get('Параметризация', '') or ''
    params['parameters'] = str(parameters).upper() if parameters else ''
    
    # Объединяем текст для парсинга
    combined_text = f"{params['work_name']} {params['parameters']}"
    
    # Класс бетона
    class_match = re.search(r'[BВ]\s?(\d+(?:[.,]\d+)?)', combined_text)
    if class_match:
        params['concrete_class'] = class_match.group(1).replace(',', '.')
    
    # Толщина t (например t>300 мм, 150<t<200 мм)
    thickness_match = re.search(r'(\d+(?:\.\d+)?)\s*<\s*t\s*<\s*(\d+(?:\.\d+)?)', combined_text)
    if thickness_match:
        params['thickness_min'] = float(thickness_match.group(1))
        params['thickness_max'] = float(thickness_match.group(2))
    else:
        t_gt_match = re.search(r't\s*>\s*(\d+(?:\.\d+)?)', combined_text)
        if t_gt_match:
            params['thickness_min'] = float(t_gt_match.group(1))
        
        t_lt_match = re.search(r't\s*<\s*(\d+(?:\.\d+)?)', combined_text)
        if t_lt_match:
            params['thickness_max'] = float(t_lt_match.group(1))
    
    # Высота H
    height_match = re.search(r'(\d+(?:\.\d+)?)\s*<\s*H\s*<\s*(\d+(?:\.\d+)?)', combined_text)
    if height_match:
        params['height_min'] = float(height_match.group(1))
        params['height_max'] = float(height_match.group(2))
    else:
        h_gt_match = re.search(r'H\s*>\s*(\d+(?:\.\d+)?)', combined_text)
        if h_gt_match:
            params['height_min'] = float(h_gt_match.group(1))
        
        h_lt_match = re.search(r'H\s*<\s*(\d+(?:\.\d+)?)', combined_text)
        if h_lt_match:
            params['height_max'] = float(h_lt_match.group(1))
    
    # Площадь S
    area_match = re.search(r'(\d+(?:\.\d+)?)\s*<\s*S\s*<\s*(\d+(?:\.\d+)?)', combined_text)
    if area_match:
        params['area_min'] = float(area_match.group(1))
        params['area_max'] = float(area_match.group(2))
    else:
        s_gt_match = re.search(r'S\s*>\s*(\d+(?:\.\d+)?)', combined_text)
        if s_gt_match:
            params['area_min'] = float(s_gt_match.group(1))
        
        s_lt_match = re.search(r'S\s*<\s*(\d+(?:\.\d+)?)', combined_text)
        if s_lt_match:
            params['area_max'] = float(s_lt_match.group(1))
    
    # Периметр P
    perimeter_match = re.search(r'(\d+(?:\.\d+)?)\s*<\s*P\s*<\s*(\d+(?:\.\d+)?)', combined_text)
    if perimeter_match:
        params['perimeter_min'] = float(perimeter_match.group(1))
        params['perimeter_max'] = float(perimeter_match.group(2))
    else:
        p_gt_match = re.search(r'P\s*>\s*(\d+(?:\.\d+)?)', combined_text)
        if p_gt_match:
            params['perimeter_min'] = float(p_gt_match.group(1))
        
        p_lt_match = re.search(r'P\s*<\s*(\d+(?:\.\d+)?)', combined_text)
        if p_lt_match:
            params['perimeter_max'] = float(p_lt_match.group(1))
    
    # Диаметр арматуры Ф (Фи)
    diameter_match = re.search(r'(\d+(?:\.\d+)?)\s*<=?\s*[ФФ]\s*<=?\s*(\d+(?:\.\d+)?)', combined_text)
    if diameter_match:
        params['diameter_min'] = float(diameter_match.group(1))
        params['diameter_max'] = float(diameter_match.group(2))
    else:
        phi_eq_match = re.search(r'[ФФ]\s*=\s*(\d+(?:\.\d+)?)', combined_text)
        if phi_eq_match:
            params['diameter_min'] = float(phi_eq_match.group(1))
            params['diameter_max'] = float(phi_eq_match.group(1))
        
        phi_gt_match = re.search(r'[ФФ]\s*>\s*(\d+(?:\.\d+)?)', combined_text)
        if phi_gt_match:
            params['diameter_min'] = float(phi_gt_match.group(1))
        
        phi_lt_match = re.search(r'[ФФ]\s*<\s*(\d+(?:\.\d+)?)', combined_text)
        if phi_lt_match:
            params['diameter_max'] = float(phi_lt_match.group(1))
    
    # Проверка на подземную часть
    if 'ПОДЗЕМН' in combined_text or 'ПОДЗЕМНАЯ ЧАСТЬ' in combined_text:
        params['is_underground'] = True
    
    return params


def check_element_work_match(element_params, work_params):
    """
    Проверяет соответствие элемента работе по параметрам.
    
    Args:
        element_params: dict параметров элемента
        work_params: dict параметров работы
        
    Returns:
        tuple (bool, score) - совпадение и оценка соответствия
    """
    score = 0
    max_score = 0
    
    # 1. Проверка класса бетона
    if work_params['concrete_class']:
        max_score += 3
        if element_params['concrete_class'] == work_params['concrete_class']:
            score += 3
    
    # 2. Проверка толщины
    if work_params['thickness_min'] is not None or work_params['thickness_max'] is not None:
        max_score += 2
        elem_thickness = element_params.get('thickness')
        if elem_thickness is not None:
            elem_t_mm = elem_thickness * 1000  # конвертируем в мм
            match = True
            if work_params['thickness_min'] is not None and elem_t_mm <= work_params['thickness_min']:
                match = False
            if work_params['thickness_max'] is not None and elem_t_mm >= work_params['thickness_max']:
                match = False
            if match:
                score += 2
    
    # 3. Проверка категории элемента vs наименования работы
    max_score += 3
    elem_cat = element_params.get('category', '').upper()
    work_name = work_params.get('work_name', '').upper()
    
    category_keywords = {
        'СТЕНА': ['СТЕН', 'WALL'],
        'КОЛОННА': ['КОЛОНН', 'COLUMN'],
        'БАЛКА': ['БАЛК', 'BEAM'],
        'ПЕРЕКРЫТИЕ_ПЛИТА': ['ПЛИТ', 'СЛАБ', 'ПЕРЕКРЫТИ'],
        'ФУНДАМЕНТ_ПЛИТА': ['ФУНДАМЕНТ', 'FOUNDATION'],
        'ЛЕСТНИЦА': ['ЛЕСТНИЧ', 'STAIR'],
    }
    
    for cat_key, keywords in category_keywords.items():
        if elem_cat == cat_key.upper() or cat_key in elem_cat:
            for kw in keywords:
                if kw in work_name:
                    score += 3
                    break
            break
    
    # 4. Проверка подземности
    if work_params['is_underground'] is not None:
        max_score += 2
        if element_params.get('underground') == work_params['is_underground']:
            score += 2
    
    # 5. Проверка диаметра арматуры (для арматурных работ)
    if work_params['diameter_min'] is not None or work_params['diameter_max'] is not None:
        max_score += 2
        # Ищем диаметр в имени элемента или характеристиках
        elem_name = element_params.get('name', '')
        diam_match = re.search(r'D\s*=?\s*(\d+(?:\.\d+)?)', elem_name)
        if diam_match:
            elem_diam = float(diam_match.group(1))
            match = True
            if work_params['diameter_min'] is not None and elem_diam < work_params['diameter_min']:
                match = False
            if work_params['diameter_max'] is not None and elem_diam > work_params['diameter_max']:
                match = False
            if match:
                score += 2
    
    # Нормализуем оценку
    if max_score == 0:
        return False, 0
    
    normalized_score = score / max_score if max_score > 0 else 0
    
    # Порог совпадения - 0.5 (50%)
    return normalized_score >= 0.5, normalized_score


def map_elements_to_works(session_folder):
    """
    Основная функция маппинга элементов к работам.
    
    Args:
        session_folder: Папка сессии
        
    Returns:
        Dict с результатами
    """
    print(f"\n{'='*60}")
    print("🔗 МАППИНГ ЭЛЕМЕНТОВ К РАБОТАМ")
    print(f"{'='*60}")
    
    # Шаг 1: Загрузка full_elements.xlsx
    elements_file = Path(session_folder) / "full_elements.xlsx"
    if not elements_file.exists():
        print(f"❌ Файл {elements_file} не найден")
        return {'success': False, 'error': 'full_elements.xlsx not found'}
    
    print(f"📁 Загрузка элементов из: {elements_file}")
    df_elements = pd.read_excel(str(elements_file))
    print(f"   ✅ Загружено {len(df_elements)} элементов")
    
    # Шаг 2: Загрузка Перечень работ КР_new.xlsx
    works_file = Path(session_folder) / "Перечень работ КР_new.xlsx"
    if not works_file.exists():
        print(f"❌ Файл {works_file} не найден")
        return {'success': False, 'error': 'Перечень работ КР_new.xlsx not found'}
    
    print(f"📁 Загрузка перечня работ из: {works_file}")
    df_works = pd.read_excel(str(works_file), sheet_name='ВОР КР+расценки')
    print(f"   ✅ Загружено {len(df_works)} работ")
    
    # Шаг 3: Маппинг
    print("\n🔍 Выполнение маппинга...")
    
    mapping_results = []
    matched_count = 0
    
    for idx, elem_row in df_elements.iterrows():
        element_params = extract_parameters_from_element(elem_row.to_dict())
        matched_works = []
        
        for work_idx, work_row in df_works.iterrows():
            work_row_dict = {
                'Наименование работ': work_row.get('Наименование работ'),
                'Формула расчёта объёмов работ и расхода материалов': work_row.get('Формула расчёта объёмов работ и расхода материалов'),
                'Параметризация': work_row.get('Параметризация')
            }
            work_params = extract_parameters_from_work(work_row_dict)
            
            is_match, score = check_element_work_match(element_params, work_params)
            
            if is_match:
                matched_works.append({
                    'work_index': work_idx,
                    'score': score,
                    'work_data': work_row.to_dict()
                })
        
        if matched_works:
            matched_count += 1
        
        # Сохраняем результат для элемента
        mapping_results.append({
            'element_index': idx,
            'element_name': elem_row.get('Наименование', ''),
            'matched_works': matched_works,
            'match_count': len(matched_works)
        })
        
        if (idx + 1) % 100 == 0:
            print(f"   Обработано {idx + 1} элементов...")
    
    print(f"\n   ✅ Смаппировано {matched_count} элементов из {len(df_elements)}")
    
    # Шаг 4: Создание результирующей таблицы
    print("\n📊 Создание результирующей таблицы...")
    
    result_rows = []
    
    for mapping in mapping_results:
        elem_name = mapping['element_name']
        matched_works = mapping['matched_works']
        
        if matched_works:
            # Для каждого смаппированного work создаем строку
            for work_match in matched_works:
                work_data = work_match['work_data']
                result_rows.append({
                    'Элемент IFC': elem_name,
                    '№ п/п работы': work_data.get('№ п/п'),
                    'Наименование работ': work_data.get('Наименование работ'),
                    'Ед. изм': work_data.get('Ед. изм'),
                    'IFC класс': work_data.get('IFC класс'),
                    'Формула расчёта': work_data.get('Формула расчёта объёмов работ и расхода материалов'),
                    'Параметризация': work_data.get('Параметризация'),
                    'Шифр ТСН': work_data.get('Шифр ТСН'),
                    'Наименование расценки/ресурса': work_data.get('Наименование расценки/ресурса'),
                    'Ед. изм.': work_data.get('Ед. изм.'),
                    'V по смете': work_data.get('V по смете'),
                    'Обозначения': work_data.get('Обозначения'),
                    'Оценка соответствия': round(work_match['score'], 2)
                })
        else:
            # Элемент без匹配的 работ
            result_rows.append({
                'Элемент IFC': elem_name,
                '№ п/п работы': None,
                'Наименование работ': None,
                'Ед. изм': None,
                'IFC класс': None,
                'Формула расчёта': None,
                'Параметризация': None,
                'Шифр ТСН': None,
                'Наименование расценки/ресурса': None,
                'Ед. изм.': None,
                'V по смете': None,
                'Обозначения': None,
                'Оценка соответствия': None
            })
    
    # Создаем DataFrame
    df_result = pd.DataFrame(result_rows)
    
    # Сохранение
    output_path = Path(session_folder) / "mapped_elements_works.xlsx"
    df_result.to_excel(str(output_path), index=False)
    print(f"✅ Результат сохранен: {output_path}")
    
    return {
        'success': True,
        'output_file': 'mapped_elements_works.xlsx',
        'output_path': str(output_path),
        'total_elements': len(df_elements),
        'matched_elements': matched_count,
        'total_works': len(df_works),
        'result_rows': len(result_rows)
    }


# Entry point when run directly
if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Использование: python map_elements_to_works.py <session_folder>")
        print("Пример: python map_elements_to_works.py /workspace/uploads/session_id")
        sys.exit(1)
    
    session_folder = sys.argv[1]
    result = map_elements_to_works(session_folder)
    
    if result.get('success'):
        print(f"\n✅ Маппинг завершен успешно!")
        print(f"   Всего элементов: {result['total_elements']}")
        print(f"   Смаппировано элементов: {result['matched_elements']}")
        print(f"   Строк в результате: {result['result_rows']}")
        print(f"   Результат: {result['output_file']}")
    else:
        print(f"\n❌ Ошибка: {result.get('error', 'Unknown error')}")
        sys.exit(1)
