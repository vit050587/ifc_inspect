#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для подбора видов работ к элементам из таблицы full_elements.xlsx
на основе таблицы Перечень работ КР_new.xlsx

Алгоритм:
1. Определяем IFC класс элемента по наименованию
2. Извлекаем характеристики материала (класс бетона, морозостойкость, водонепроницаемость)
3. Определяем геометрические параметры (толщина, площадь сечения, объем, высота и т.д.)
4. Подбираем виды работ по совпадению IFC класса, условий параметризации и характеристик материала
"""

import pandas as pd
import re
from typing import List, Dict, Optional, Tuple


def parse_concrete_grade(material: str, characteristics: str) -> Optional[str]:
    """Извлекает класс бетона из материала или характеристик"""
    if pd.isna(material) and pd.isna(characteristics):
        return None
    
    # Сначала пробуем извлечь из материала
    if pd.notna(material):
        match = re.search(r'В(\d+)', str(material))
        if match:
            return f"В{match.group(1)}"
    
    # Затем из характеристик
    if pd.notna(characteristics):
        match = re.search(r'Класс бетона:\s*В(\d+)', str(characteristics))
        if match:
            return f"В{match.group(1)}"
    
    return None


def parse_freeze_durability(characteristics: str) -> Optional[str]:
    """Извлекает марку морозостойкости F"""
    if pd.isna(characteristics):
        return None
    
    match = re.search(r'Морозостойкость:\s*F(\d+)', str(characteristics))
    if match:
        return f"F{match.group(1)}"
    
    # Пробуем из материала
    match = re.search(r'F(\d{3})', str(characteristics))
    if match:
        return f"F{match.group(1)}"
    
    return None


def parse_water_resist(characteristics: str) -> Optional[str]:
    """Извлекает марку водонепроницаемости W"""
    if pd.isna(characteristics):
        return None
    
    match = re.search(r'Водонепроницаемость:\s*W(\d+)', str(characteristics))
    if match:
        return f"W{match.group(1)}"
    
    # Пробуем из материала
    match = re.search(r'W(\d+)', str(characteristics))
    if match:
        return f"W{match.group(1)}"
    
    return None


def determine_ifc_class(row: pd.Series) -> str:
    """Определяет IFC класс элемента по наименованию и другим параметрам"""
    name = str(row.get('Наименование', ''))
    
    # Сопоставление по ключевым словам в наименовании
    if 'Колонн' in name or 'НесКол' in name:
        return 'IfcColumn'
    elif 'Балк' in name or 'Ригел' in name:
        return 'IfcBeam'
    elif 'Стен' in name or 'Диафрагм' in name:
        return 'IfcWall'
    elif 'Плит' in name or 'Перекрыт' in name or 'Термовкладыш' in name:
        return 'IfcSlab'
    elif 'Лестничн' in name or 'Марш' in name:
        return 'IfcStairFlight'
    elif 'Лестничн' in name and 'Площадк' in name:
        return 'IfcSlab'
    elif 'Фундамент' in name:
        return 'IfcFooting'
    elif 'Арматур' in name or 'Армирован' in name:
        return 'IfcReinforcingBar'
    elif 'Закладн' in name or 'Пластина' in name:
        return 'IfcPlate'
    elif 'Отверст' in name:
        return 'IfcOpeningElement'
    elif 'Гидрошпонк' in name:
        return 'IfcDiscreteAccessory'
    elif 'Труб' in name:
        return 'IfcBuildingElementProxy'
    elif 'Приям' in name:
        return 'IfcSlab'
    elif 'Рен' in name:
        return 'IfcCovering'
    
    # Если не нашли соответствие, возвращаем пустую строку
    return ''


def check_thickness_condition(element_thickness: float, condition: str) -> bool:
    """Проверяет условие по толщине"""
    if pd.isna(element_thickness) or element_thickness <= 0:
        return False
    
    # Переводим мм в м если нужно
    thickness_mm = element_thickness * 1000 if element_thickness < 10 else element_thickness * 1000
    
    try:
        if 't<' in condition or 't <' in condition:
            limit = float(re.search(r't\s*<\s*(\d+)', condition).group(1))
            return thickness_mm < limit
        elif 't>' in condition or 't >' in condition:
            limit = float(re.search(r't\s*>\s*(\d+)', condition).group(1))
            return thickness_mm > limit
        elif '<t<' in condition or '< t <' in condition:
            match = re.search(r'(\d+)\s*<\s*t\s*<\s*(\d+)', condition)
            if match:
                lower = float(match.group(1))
                upper = float(match.group(2))
                return lower < thickness_mm < upper
    except (AttributeError, ValueError):
        pass
    
    return False


def check_volume_condition(element_volume: float, condition: str) -> bool:
    """Проверяет условие по объему"""
    if pd.isna(element_volume) or element_volume <= 0:
        return False
    
    try:
        if 'V<' in condition or 'V <' in condition:
            limit = float(re.search(r'V\s*<\s*(\d+)', condition).group(1))
            return element_volume < limit
        elif 'V>' in condition or 'V >' in condition:
            limit = float(re.search(r'V\s*>\s*(\d+)', condition).group(1))
            return element_volume > limit
        elif '<V<' in condition or '< V <' in condition:
            match = re.search(r'(\d+)\s*<\s*V\s*<\s*(\d+)', condition)
            if match:
                lower = float(match.group(1))
                upper = float(match.group(2))
                return lower < element_volume < upper
    except (AttributeError, ValueError):
        pass
    
    return False


def check_area_condition(element_area: float, condition: str) -> bool:
    """Проверяет условие по площади"""
    if pd.isna(element_area) or element_area <= 0:
        return False
    
    try:
        if 'S<' in condition or 'S <' in condition:
            limit = float(re.search(r'S\s*<\s*([\d.]+)', condition).group(1))
            return element_area < limit
        elif 'S>' in condition or 'S >' in condition:
            limit = float(re.search(r'S\s*>\s*([\d.]+)', condition).group(1))
            return element_area > limit
        elif '<S<' in condition or '< S <' in condition:
            match = re.search(r'([\d.]+)\s*<\s*S\s*<\s*([\d.]+)', condition)
            if match:
                lower = float(match.group(1))
                upper = float(match.group(2))
                return lower < element_area < upper
    except (AttributeError, ValueError):
        pass
    
    return False


def check_rebar_diameter_condition(condition: str, diameter: Optional[float] = None) -> bool:
    """Проверяет условие по диаметру арматуры"""
    try:
        # Проверяем наличие условия по диаметру в формате Ф = X или X<=Ф<=Y
        if 'Ф =' in condition or 'Ф=' in condition:
            match = re.search(r'Ф\s*=\s*(\d+)', condition)
            if match:
                target_diameter = int(match.group(1))
                if diameter is not None:
                    return abs(diameter - target_diameter) < 1
                return True  # Если диаметр не указан, считаем что условие выполняется
        
        if '<=Ф<=' in condition or '<= Ф <=' in condition:
            match = re.search(r'(\d+)\s*<=\s*Ф\s*<=\s*(\d+)', condition)
            if match:
                lower = int(match.group(1))
                upper = int(match.group(2))
                if diameter is not None:
                    return lower <= diameter <= upper
                return True  # Если диаметр не указан, считаем что условие выполняется
    except (AttributeError, ValueError):
        pass
    
    return False


def check_concrete_condition(element_grade: str, element_freeze: str, element_water: str, 
                              condition: str) -> bool:
    """Проверяет условие по классу бетона и маркам"""
    if pd.isna(condition):
        return True
    
    condition_str = str(condition).upper()
    
    # Проверяем класс бетона
    grade_match = re.search(r'[BВ](\d+)', condition_str)
    if grade_match and element_grade:
        required_grade = f"В{grade_match.group(1)}"
        if required_grade != element_grade:
            return False
    
    # Проверяем морозостойкость
    freeze_match = re.search(r'F(\d+)', condition_str)
    if freeze_match and element_freeze:
        required_freeze = f"F{freeze_match.group(1)}"
        if required_freeze != element_freeze:
            return False
    
    # Проверяем водонепроницаемость
    water_match = re.search(r'W(\d+)', condition_str)
    if water_match and element_water:
        required_water = f"W{water_match.group(1)}"
        if required_water != element_water:
            return False
    
    return True


def get_element_parameters(row: pd.Series) -> Dict:
    """Извлекает параметры элемента из строки таблицы"""
    params = {
        'ifc_class': determine_ifc_class(row),
        'concrete_grade': parse_concrete_grade(row.get('Материал', ''), 
                                                row.get('Характеристики материала', '')),
        'freeze_durability': parse_freeze_durability(row.get('Характеристики материала', '')),
        'water_resist': parse_water_resist(row.get('Характеристики материала', '')),
        'thickness_m': row.get('Толщина, м', float('nan')),
        'width_m': row.get('Ширина, м', float('nan')),
        'height_m': row.get('Высота, м', float('nan')),
        'length_m': row.get('Длина, м', float('nan')),
        'volume_m3': row.get('Объем, м³', float('nan')),
        'area_m2': row.get('Площадь, м²', float('nan')),
    }
    
    # Вычисляем площадь сечения для колонн и балок
    if pd.notna(params['width_m']) and pd.notna(params['height_m']):
        params['cross_section_area'] = params['width_m'] * params['height_m']
    else:
        params['cross_section_area'] = float('nan')
    
    # Вычисляем периметр сечения
    if pd.notna(params['width_m']) and pd.notna(params['height_m']):
        params['perimeter'] = 2 * (params['width_m'] + params['height_m'])
    else:
        params['perimeter'] = float('nan')
    
    return params


def match_work_to_element(element_params: Dict, work_row: pd.Series) -> bool:
    """Проверяет, подходит ли вид работы к элементу"""
    
    # 1. Проверяем IFC класс
    work_ifc = str(work_row.get('IFC класс', ''))
    element_ifc = element_params['ifc_class']
    
    # Если в работе указано "не моделируется" или "-", то пропускаем проверку IFC
    if work_ifc not in ['не моделируется', 'чаще всего не моделируется', '-', '']:
        if element_ifc != work_ifc:
            return False
    
    # 2. Проверяем параметризацию
    parametrization = str(work_row.get('Параметризация', ''))
    if pd.notna(parametrization) and parametrization not in ['nan', '']:
        param_lines = parametrization.split('\n')
        
        for param_line in param_lines:
            param_line = param_line.strip()
            if not param_line:
                continue
            
            # Проверяем условия по толщине
            if 't' in param_line.lower() and ('<' in param_line or '>' in param_line):
                if not check_thickness_condition(element_params['thickness_m'], param_line):
                    return False
            
            # Проверяем условия по объему
            if 'v' in param_line.lower() and ('<' in param_line or '>' in param_line):
                if not check_volume_condition(element_params['volume_m3'], param_line):
                    return False
            
            # Проверяем условия по площади
            if 's' in param_line.lower() and ('<' in param_line or '>' in param_line):
                if not check_area_condition(element_params['area_m2'], param_line):
                    return False
            
            # Проверяем условия по бетону и маркам
            if re.search(r'[bв]\d+|f\d+|w\d+', param_line.lower()):
                if not check_concrete_condition(
                    element_params['concrete_grade'],
                    element_params['freeze_durability'],
                    element_params['water_resist'],
                    param_line
                ):
                    return False
    
    return True


def find_works_for_element(element_row: pd.Series, works_df: pd.DataFrame) -> List[Dict]:
    """Находит все подходящие виды работ для элемента"""
    
    element_params = get_element_parameters(element_row)
    matched_works = []
    
    for idx, work_row in works_df.iterrows():
        # Пропускаем заголовки разделов
        work_name = str(work_row.get('Наименование работ', ''))
        if pd.isna(work_row.get('IFC класс', '')) and pd.isna(work_row.get('Шифр ТСН', '')):
            continue
        
        if match_work_to_element(element_params, work_row):
            matched_works.append({
                'work_idx': idx,
                'work_name': work_row.get('Наименование работ', ''),
                'unit': work_row.get('Ед. изм', ''),
                'tsn_code': work_row.get('Шифр ТСН', ''),
                'description': work_row.get('Наименование расценки/ресурса', ''),
                'formula': work_row.get('Формула расчёта объёмов работ и расхода материалов', ''),
                'parametrization': work_row.get('Параметризация', ''),
            })
    
    return matched_works


def main():
    # Пути к файлам
    elements_file = 'uploads/b43c911d-191d-42e5-a953-810feb9bf2de/full_elements.xlsx'
    works_file = 'uploads/b43c911d-191d-42e5-a953-810feb9bf2de/Перечень работ КР_new.xlsx'
    output_file = 'uploads/b43c911d-191d-42e5-a953-810feb9bf2de/elements_with_works.xlsx'
    
    print("Загрузка таблиц...")
    df_elements = pd.read_excel(elements_file)
    df_works = pd.read_excel(works_file)
    
    print(f"Элементы: {len(df_elements)} строк")
    print(f"Виды работ: {len(df_works)} строк")
    
    # Предварительно обрабатываем виды работ - фильтруем заголовки разделов
    valid_works_mask = df_works['IFC класс'].notna() | df_works['Шифр ТСН'].notna()
    df_works_valid = df_works[valid_works_mask].copy()
    print(f"Валидных видов работ: {len(df_works_valid)}")
    
    # Группируем работы по IFC классам для ускорения поиска
    works_by_ifc = {}
    for ifc_class in df_works_valid['IFC класс'].unique():
        if pd.notna(ifc_class) and ifc_class not in ['не моделируется', 'чаще всего не моделируется', '-', '']:
            works_by_ifc[ifc_class] = df_works_valid[df_works_valid['IFC класс'] == ifc_class]
    
    # Работы без привязки к IFC классу (универсальные)
    works_universal = df_works_valid[df_works_valid['IFC класс'].isin(['не моделируется', 'чаще всего не моделируется', '-', ''])]
    
    # Обрабатываем элементы
    results = []
    
    print("\nПодбор видов работ для элементов...")
    for idx, element_row in df_elements.iterrows():
        if idx % 500 == 0:
            print(f"Обработано {idx} из {len(df_elements)} элементов...")
        
        element_params = get_element_parameters(element_row)
        element_ifc = element_params['ifc_class']
        
        # Собираем подходящие работы
        candidate_works = []
        
        # Добавляем работы для конкретного IFC класса
        if element_ifc in works_by_ifc:
            for _, work_row in works_by_ifc[element_ifc].iterrows():
                candidate_works.append(work_row)
        
        # Добавляем универсальные работы
        for _, work_row in works_universal.iterrows():
            candidate_works.append(work_row)
        
        # Проверяем каждую кандидатуру
        matched_works = []
        for work_row in candidate_works:
            if match_work_to_element(element_params, work_row):
                matched_works.append({
                    'work_idx': work_row.name,
                    'work_name': work_row.get('Наименование работ', ''),
                    'unit': work_row.get('Ед. изм', ''),
                    'tsn_code': work_row.get('Шифр ТСН', ''),
                    'description': work_row.get('Наименование расценки/ресурса', ''),
                    'formula': work_row.get('Формула расчёта объёмов работ и расхода материалов', ''),
                    'parametrization': work_row.get('Параметризация', ''),
                })
        
        if matched_works:
            for work in matched_works:
                result = {
                    'Element_Index': idx,
                    'Element_Name': element_row.get('Наименование', ''),
                    'Element_Material': element_row.get('Материал', ''),
                    'Element_Characteristics': element_row.get('Характеристики материала', ''),
                    'IFC_Class': element_ifc,
                    'Volume_m3': element_row.get('Объем, м³', ''),
                    'Thickness_m': element_row.get('Толщина, м', ''),
                    'Width_m': element_row.get('Ширина, м', ''),
                    'Height_m': element_row.get('Высота, м', ''),
                    'Work_Name': work['work_name'],
                    'Work_Unit': work['unit'],
                    'TSN_Code': work['tsn_code'],
                    'Work_Description': work['description'],
                    'Formula': work['formula'],
                    'Parametrization': work['parametrization'],
                }
                results.append(result)
    
    # Создаем DataFrame с результатами
    df_results = pd.DataFrame(results)
    
    print(f"\nНайдено {len(df_results)} соответствий элемент-работа")
    print(f"Уникальных элементов с работами: {df_results['Element_Index'].nunique()}")
    
    # Сохраняем результат
    df_results.to_excel(output_file, index=False)
    print(f"\nРезультаты сохранены в файл: {output_file}")
    
    # Выводим статистику
    print("\n=== Статистика по типам элементов ===")
    ifc_counts = df_results['IFC_Class'].value_counts()
    for ifc_class, count in ifc_counts.items():
        print(f"{ifc_class}: {count} работ")
    
    print("\n=== Примеры результатов (первые 10) ===")
    print(df_results.head(10).to_string())


if __name__ == '__main__':
    main()
