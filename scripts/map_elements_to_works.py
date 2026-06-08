#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для подбора видов работ к элементам из таблицы full_elements.xlsx
на основе таблицы Перечень работ КР_new.xlsx (версия 2)

Алгоритм:
1. Определяем IFC класс элемента по наименованию
2. Проверяем уровень элемента (подземный/надземный)
3. Извлекаем характеристики материала (класс бетона, морозостойкость, водонепроницаемость)
4. Определяем геометрические параметры (толщина, площадь сечения, объем, высота и т.д.)
5. Подбираем виды работ по совпадению IFC класса, условий параметризации и характеристик материала
6. Элементы без работ: термовкладыши, отверстия, трубы не получают виды работ

Ключевые особенности v2:
- Учет уровня элемента (подземный/надземный) для выбора работ
- Исключение элементов без работ (термовкладыши, отверстия, трубы)
- Проверка наличия IFC класса в таблице работ
- Расширенная проверка параметров через регулярные выражения
"""

import pandas as pd
import re
from typing import List, Dict, Optional, Tuple
import sys
import os


def parse_concrete_grade(material: str, characteristics: str) -> Optional[str]:
    """Извлекает класс бетона из материала или характеристик"""
    if pd.isna(material) and pd.isna(characteristics):
        return None
    
    # Сначала пробуем извлечь из материала
    if pd.notna(material):
        match = re.search(r'[ВB](\d+)', str(material))
        if match:
            return f"В{match.group(1)}"
    
    # Затем из характеристик
    if pd.notna(characteristics):
        match = re.search(r'Класс бетона:\s*[ВB](\d+)', str(characteristics))
        if match:
            return f"В{match.group(1)}"
        
        # Пробуем просто найти класс бетона в строке
        match = re.search(r'[ВB](\d{2})', str(characteristics))
        if match:
            return f"В{match.group(1)}"
    
    return None


def parse_freeze_durability(material: str, characteristics: str) -> Optional[str]:
    """Извлекает марку морозостойкости F"""
    text = ""
    if pd.notna(material):
        text += str(material) + " "
    if pd.notna(characteristics):
        text += str(characteristics)
    
    if not text.strip():
        return None
    
    # Ищем в характеристиках
    match = re.search(r'Морозостойкость:\s*F(\d+)', text)
    if match:
        return f"F{match.group(1)}"
    
    # Ищем просто FXXX в тексте
    match = re.search(r'F(\d{3})', text)
    if match:
        return f"F{match.group(1)}"
    
    return None


def parse_water_resist(material: str, characteristics: str) -> Optional[str]:
    """Извлекает марку водонепроницаемости W"""
    text = ""
    if pd.notna(material):
        text += str(material) + " "
    if pd.notna(characteristics):
        text += str(characteristics)
    
    if not text.strip():
        return None
    
    # Ищем в характеристиках
    match = re.search(r'Водонепроницаемость:\s*W(\d+)', text)
    if match:
        return f"W{match.group(1)}"
    
    # Ищем просто WX в тексте
    match = re.search(r'W(\d+)', text)
    if match:
        return f"W{match.group(1)}"
    
    return None


def parse_aggregate_type(material: str, characteristics: str) -> Optional[str]:
    """Извлекает тип заполнителя (гранитный щебень и т.д.)"""
    text = ""
    if pd.notna(material):
        text += str(material) + " "
    if pd.notna(characteristics):
        text += str(characteristics)
    
    if not text.strip():
        return None
    
    if 'гранитный щебень' in text.lower():
        return 'гранитный щебень'
    elif 'гравий' in text.lower():
        return 'гравий'
    elif 'известняк' in text.lower():
        return 'известняк'
    
    return None


def determine_ifc_class(row: pd.Series) -> str:
    """Определяет IFC класс элемента по наименованию и другим параметрам"""
    name = str(row.get('Наименование', ''))
    
    # Сначала проверяем на элементы без работ
    if 'Термовкладыш' in name or 'термовкладыш' in name:
        return 'IfcThermalInsert'  # Специальный класс для элементов без работ
    
    if 'Отверст' in name or 'отверст' in name or 'Проем' in name:
        return 'IfcOpeningElement'
    
    if 'Труб' in name and ('Гильз' in name or 'гильз' in name):
        return 'IfcBuildingElementProxy'
    
    # Сопоставление по ключевым словам в наименовании
    if 'Колонн' in name or 'НесКол' in name or 'колonn' in name.lower():
        return 'IfcColumn'
    elif 'Балк' in name or 'Ригел' in name or 'балк' in name.lower():
        return 'IfcBeam'
    elif 'Стен' in name or 'Диафрагм' in name or 'стен' in name.lower():
        return 'IfcWall'
    elif 'Плит' in name or 'Перекрыт' in name or 'Покрыт' in name or 'Фундаментн' in name:
        # Различаем фундаментную плиту и перекрытие
        if 'Фундамент' in name or 'фундамент' in name.lower() or 'ФП' in name:
            return 'IfcSlabFoundation'
        return 'IfcSlab'
    elif 'Лестничн' in name and 'Марш' in name or 'марш' in name.lower():
        return 'IfcStairFlight'
    elif 'Лестничн' in name and 'Площадк' in name:
        return 'IfcSlab'
    elif 'Фундамент' in name and 'плит' not in name.lower():
        return 'IfcFooting'
    elif 'Ростверк' in name or 'ростверк' in name.lower():
        return 'IfcFooting'
    elif 'Арматур' in name or 'Армирован' in name:
        return 'IfcReinforcingBar'
    elif 'Закладн' in name or 'Пластина' in name:
        return 'IfcPlate'
    elif 'Гидрошпонк' in name:
        return 'IfcDiscreteAccessory'
    elif 'Приям' in name:
        return 'IfcSlab'
    elif 'Рен' in name or 'стяжк' in name.lower() or 'Подготовка' in name:
        return 'IfcCovering'
    elif 'Лестниц' in name or 'сход' in name.lower():
        return 'IfcStairFlight'
    elif 'Пандус' in name or 'пандус' in name.lower():
        return 'IfcRamp'
    
    # Если не нашли соответствие, возвращаем пустую строку
    return ''


def is_element_without_works(ifc_class: str, name: str) -> bool:
    """Проверяет, является ли элемент элементом без работ"""
    # Элементы которые не получают виды работ
    no_works_classes = ['IfcOpeningElement', 'IfcBuildingElementProxy', 'IfcThermalInsert']
    
    if ifc_class in no_works_classes:
        return True
    
    # Дополнительная проверка по имени для термовкладышей
    if 'термовкладыш' in name.lower():
        return True
    
    return False


def get_element_level(row: pd.Series) -> str:
    """Определяет уровень элемента (подземный/надземный)"""
    level = row.get('Подземный/Надземный', '')
    if pd.isna(level):
        return 'Не определено'
    
    level_str = str(level).strip()
    if 'подземн' in level_str.lower() or 'подз.' in level_str.lower():
        return 'Подземный'
    elif 'надземн' in level_str.lower() or 'надз.' in level_str.lower():
        return 'Надземный'
    else:
        return level_str


def check_thickness_condition(element_thickness_m: float, condition: str) -> bool:
    """Проверяет условие по толщине (t в мм)"""
    if pd.isna(element_thickness_m) or element_thickness_m <= 0:
        return False
    
    # Переводим метры в миллиметры
    thickness_mm = element_thickness_m * 1000
    
    try:
        # t<XXX мм
        if re.search(r't\s*<\s*(\d+)', condition):
            limit = float(re.search(r't\s*<\s*(\d+)', condition).group(1))
            return thickness_mm < limit
        
        # t>XXX мм
        if re.search(r't\s*>\s*(\d+)', condition):
            limit = float(re.search(r't\s*>\s*(\d+)', condition).group(1))
            return thickness_mm > limit
        
        # XXX<t<YYY мм
        match = re.search(r'(\d+)\s*<\s*t\s*<\s*(\d+)', condition)
        if match:
            lower = float(match.group(1))
            upper = float(match.group(2))
            return lower < thickness_mm < upper
        
    except (AttributeError, ValueError):
        pass
    
    return False


def check_volume_condition(element_volume: float, condition: str) -> bool:
    """Проверяет условие по объему (V в м³)"""
    if pd.isna(element_volume) or element_volume <= 0:
        return False
    
    try:
        # V<XXX м3
        if re.search(r'V\s*<\s*(\d+)', condition):
            limit = float(re.search(r'V\s*<\s*(\d+)', condition).group(1))
            return element_volume < limit
        
        # V>XXX м3
        if re.search(r'V\s*>\s*(\d+)', condition):
            limit = float(re.search(r'V\s*>\s*(\d+)', condition).group(1))
            return element_volume > limit
        
        # XXX<V<YYY м3
        match = re.search(r'(\d+)\s*<\s*V\s*<\s*(\d+)', condition)
        if match:
            lower = float(match.group(1))
            upper = float(match.group(2))
            return lower < element_volume < upper
        
    except (AttributeError, ValueError):
        pass
    
    return False


def check_area_condition(element_area: float, condition: str) -> bool:
    """Проверяет условие по площади (S в м²)"""
    if pd.isna(element_area) or element_area <= 0:
        return False
    
    try:
        # S<XXX м2
        match = re.search(r'S\s*<\s*([\d.]+)\s*м?2?', condition)
        if match:
            limit = float(match.group(1))
            return element_area < limit
        
        # S>XXX м2
        match = re.search(r'S\s*>\s*([\d.]+)\s*м?2?', condition)
        if match:
            limit = float(match.group(1))
            return element_area > limit
        
        # XXX<S<YYY м2
        match = re.search(r'([\d.]+)\s*<\s*S\s*<\s*([\d.]+)', condition)
        if match:
            lower = float(match.group(1))
            upper = float(match.group(2))
            return lower < element_area < upper
        
    except (AttributeError, ValueError):
        pass
    
    return False


def check_cross_section_area_condition(width: float, height: float, condition: str) -> bool:
    """Проверяет условие по площади сечения (a или S для колонн/балок)"""
    if pd.isna(width) or pd.isna(height) or width <= 0 or height <= 0:
        return False
    
    cross_section = width * height  # м²
    
    try:
        # a<XXX мм (для колонн)
        match = re.search(r'a\s*<\s*(\d+)\s*мм', condition)
        if match:
            limit_mm = float(match.group(1))
            limit_m = limit_mm / 1000
            # Для квадратного сечения сравниваем сторону
            min_side = min(width, height)
            return min_side < limit_m
        
        # a>XXX мм
        match = re.search(r'a\s*>\s*(\d+)\s*мм', condition)
        if match:
            limit_mm = float(match.group(1))
            limit_m = limit_mm / 1000
            max_side = max(width, height)
            return max_side > limit_m
        
        # XXX<a<YYY мм
        match = re.search(r'(\d+)\s*<\s*a\s*<\s*(\d+)\s*мм', condition)
        if match:
            lower_mm = float(match.group(1))
            upper_mm = float(match.group(2))
            avg_side = (width + height) / 2
            return lower_mm/1000 < avg_side < upper_mm/1000
        
        # S<XXX м2 (площадь сечения)
        match = re.search(r'S\s*<\s*([\d.]+)\s*м2', condition)
        if match:
            limit = float(match.group(1))
            return cross_section < limit
        
        # S>XXX м2
        match = re.search(r'S\s*>\s*([\d.]+)\s*м2', condition)
        if match:
            limit = float(match.group(1))
            return cross_section > limit
        
        # XXX<S<YYY м2
        match = re.search(r'([\d.]+)\s*<\s*S\s*<\s*([\d.]+)\s*м2', condition)
        if match:
            lower = float(match.group(1))
            upper = float(match.group(2))
            return lower < cross_section < upper
        
    except (AttributeError, ValueError):
        pass
    
    return False


def check_perimeter_condition(width: float, height: float, condition: str) -> bool:
    """Проверяет условие по периметру (P в мм)"""
    if pd.isna(width) or pd.isna(height) or width <= 0 or height <= 0:
        return False
    
    perimeter_mm = 2 * (width + height) * 1000  # переводим в мм
    
    try:
        # P<XXX мм
        match = re.search(r'P\s*<\s*(\d+)\s*мм', condition)
        if match:
            limit = float(match.group(1))
            return perimeter_mm < limit
        
        # P>XXX мм
        match = re.search(r'P\s*>\s*(\d+)\s*мм', condition)
        if match:
            limit = float(match.group(1))
            return perimeter_mm > limit
        
    except (AttributeError, ValueError):
        pass
    
    return False


def check_height_condition(element_height: float, condition: str) -> bool:
    """Проверяет условие по высоте (H в м)"""
    if pd.isna(element_height) or element_height <= 0:
        return False
    
    try:
        # H<XXX м
        match = re.search(r'H\s*<\s*(\d+)\s*м', condition)
        if match:
            limit = float(match.group(1))
            return element_height < limit
        
        # H>XXX м
        match = re.search(r'H\s*>\s*(\d+)\s*м', condition)
        if match:
            limit = float(match.group(1))
            return element_height > limit
        
        # XXX<H<YYY м
        match = re.search(r'(\d+)\s*<\s*H\s*<\s*(\d+)\s*м', condition)
        if match:
            lower = float(match.group(1))
            upper = float(match.group(2))
            return lower < element_height < upper
        
    except (AttributeError, ValueError):
        pass
    
    return False


def check_rebar_diameter_condition(condition: str, diameter: Optional[float] = None) -> bool:
    """Проверяет условие по диаметру арматуры (Ф в мм)"""
    try:
        # Ф = X мм
        match = re.search(r'Ф\s*=\s*(\d+)', condition)
        if match:
            target_diameter = int(match.group(1))
            if diameter is not None:
                return abs(diameter - target_diameter) < 1
            return True  # Если диаметр не указан, считаем что условие выполняется
        
        # X<=Ф<=Y мм
        match = re.search(r'(\d+)\s*<=\s*Ф\s*<=\s*(\d+)', condition)
        if match:
            lower = int(match.group(1))
            upper = int(match.group(2))
            if diameter is not None:
                return lower <= diameter <= upper
            return True
        
    except (AttributeError, ValueError):
        pass
    
    return False


def check_concrete_condition(element_grade: str, element_freeze: str, element_water: str, 
                              element_aggregate: str, condition: str) -> bool:
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
    
    # Проверяем тип заполнителя
    if 'ГРАНИТНЫЙ ЩЕБЕНЬ' in condition_str.upper():
        if element_aggregate != 'гранитный щебень':
            return False
    
    return True


def get_element_parameters(row: pd.Series) -> Dict:
    """Извлекает параметры элемента из строки таблицы"""
    material = row.get('Материал', '')
    characteristics = row.get('Характеристики материала', '')
    
    params = {
        'ifc_class': determine_ifc_class(row),
        'name': row.get('Наименование', ''),
        'level': get_element_level(row),
        'concrete_grade': parse_concrete_grade(material, characteristics),
        'freeze_durability': parse_freeze_durability(material, characteristics),
        'water_resist': parse_water_resist(material, characteristics),
        'aggregate_type': parse_aggregate_type(material, characteristics),
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
        params['perimeter_m'] = 2 * (params['width_m'] + params['height_m'])
    else:
        params['perimeter_m'] = float('nan')
    
    # Определяем, является ли элемент без работ
    params['no_works'] = is_element_without_works(params['ifc_class'], params['name'])
    
    return params


def normalize_ifc_class(ifc_str: str) -> str:
    """Нормализует IFC класс для сравнения"""
    if pd.isna(ifc_str):
        return ''
    return str(ifc_str).strip().lower()


def ifc_class_matches(element_ifc: str, work_ifc: str) -> bool:
    """Проверяет совпадение IFC классов с учетом множественных значений"""
    if not work_ifc or work_ifc in ['не моделируется', 'чаще всего не моделируется', '-', '']:
        return True  # Универсальные работы подходят ко всем
    
    # Нормализуем для сравнения
    elem_ifc_norm = element_ifc.lower()
    work_ifc_norm = work_ifc.lower()
    
    # Работа может иметь несколько IFC классов через \n
    work_ifcs = [x.strip() for x in work_ifc_norm.split('\n')]
    
    for w_ifc in work_ifcs:
        if w_ifc in ['не моделируется', 'чаще всего не моделируется', '-', '']:
            continue
        
        # Прямое совпадение
        if elem_ifc_norm == w_ifc:
            return True
        
        # Особые случаи
        # IfcSlabFoundation может соответствовать IfcSlab или IfcFooting
        if elem_ifc_norm == 'ifcslabfoundation':
            if w_ifc in ['ifcslab', 'ifcfooting']:
                return True
        
        # IfcCovering (строчные буквы в таблице)
        if w_ifc == 'ifccovering' and elem_ifc_norm == 'ifccovering':
            return True
    
    return False


def match_work_to_element(element_params: Dict, work_row: pd.Series) -> bool:
    """Проверяет, подходит ли вид работы к элементу"""
    
    # 1. Проверяем, является ли элемент элементом без работ
    if element_params['no_works']:
        return False
    
    # 2. Проверяем IFC класс
    work_ifc = str(work_row.get('IFC класс', ''))
    element_ifc = element_params['ifc_class']
    
    if not ifc_class_matches(element_ifc, work_ifc):
        return False
    
    # 3. Проверяем параметризацию
    parametrization = str(work_row.get('Параметризация', ''))
    if pd.notna(parametrization) and parametrization not in ['nan', '', '-']:
        param_lines = parametrization.split('\n')
        
        for param_line in param_lines:
            param_line = param_line.strip()
            if not param_line or param_line in ['-', 'nan']:
                continue
            
            # Проверяем условия по толщине (t)
            if re.search(r'\bt\b', param_line.lower()) and ('<' in param_line or '>' in param_line):
                if not check_thickness_condition(element_params['thickness_m'], param_line):
                    return False
            
            # Проверяем условия по объему (V)
            if re.search(r'\bv\b', param_line.lower()) and ('<' in param_line or '>' in param_line):
                if not check_volume_condition(element_params['volume_m3'], param_line):
                    return False
            
            # Проверяем условия по площади (S)
            if re.search(r'\bs\b', param_line.lower()) and ('<' in param_line or '>' in param_line):
                # Сначала проверяем как площадь сечения для колонн/балок
                if element_params['ifc_class'] in ['ifccolumn', 'ifcbeam']:
                    if not check_cross_section_area_condition(
                        element_params['width_m'], 
                        element_params['height_m'], 
                        param_line
                    ):
                        return False
                else:
                    if not check_area_condition(element_params['area_m2'], param_line):
                        return False
            
            # Проверяем условия по периметру (P)
            if re.search(r'\bp\b', param_line.lower()) and ('<' in param_line or '>' in param_line):
                if not check_perimeter_condition(
                    element_params['width_m'], 
                    element_params['height_m'], 
                    param_line
                ):
                    return False
            
            # Проверяем условия по высоте (H)
            if re.search(r'\bh\b', param_line.lower()) and ('<' in param_line or '>' in param_line):
                if not check_height_condition(element_params['height_m'], param_line):
                    return False
            
            # Проверяем условия по бетону и маркам
            if re.search(r'[bв]\d+|f\d+|w\d+|щебень', param_line.lower()):
                if not check_concrete_condition(
                    element_params['concrete_grade'],
                    element_params['freeze_durability'],
                    element_params['water_resist'],
                    element_params['aggregate_type'],
                    param_line
                ):
                    return False
            
            # Проверяем условия по диаметру арматуры (Ф)
            if 'ф' in param_line.lower() or 'ф=' in param_line:
                if not check_rebar_diameter_condition(param_line):
                    return False
    
    return True


def find_works_for_element(element_row: pd.Series, works_df: pd.DataFrame, 
                           works_by_ifc: Dict, works_universal: pd.DataFrame) -> List[Dict]:
    """Находит все подходящие виды работ для элемента"""
    
    element_params = get_element_parameters(element_row)
    
    # Если элемент без работ, сразу возвращаем пустой список
    if element_params['no_works']:
        return []
    
    matched_works = []
    element_ifc = element_params['ifc_class'].lower()
    
    # Собираем подходящие работы
    candidate_works = []
    
    # Добавляем работы для конкретного IFC класса
    if element_ifc in works_by_ifc:
        for _, work_row in works_by_ifc[element_ifc].iterrows():
            candidate_works.append(work_row)
    
    # Добавляем универсальные работы
    if works_universal is not None:
        for _, work_row in works_universal.iterrows():
            candidate_works.append(work_row)
    
    # Проверяем каждую кандидатуру
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
    
    return matched_works


def load_and_prepare_works(works_file: str) -> Tuple[pd.DataFrame, Dict, pd.DataFrame]:
    """Загружает и подготавливает таблицу работ"""
    print(f"Загрузка таблицы работ из {works_file}...")
    
    # Пробуем разные варианты названия листа
    try:
        df_works = pd.read_excel(works_file, sheet_name='ВОР КР+расценки')
    except:
        try:
            df_works = pd.read_excel(works_file, sheet_name='Перечень работ КР_new')
        except:
            df_works = pd.read_excel(works_file)
    
    print(f"Виды работ: {len(df_works)} строк")
    
    # Предварительно обрабатываем виды работ - фильтруем заголовки разделов
    valid_works_mask = df_works['IFC класс'].notna() | df_works['Шифр ТСН'].notna()
    df_works_valid = df_works[valid_works_mask].copy()
    print(f"Валидных видов работ: {len(df_works_valid)}")
    
    # Группируем работы по IFC классам для ускорения поиска
    works_by_ifc = {}
    
    # Получаем уникальные IFC классы
    unique_ifcs = df_works_valid['IFC класс'].dropna().unique()
    
    for ifc_class in unique_ifcs:
        if pd.notna(ifc_class) and ifc_class not in ['не моделируется', 'чаще всего не моделируется', '-', '']:
            # Нормализуем IFC класс (приводим к нижнему регистру)
            ifc_lower = str(ifc_class).lower()
            
            # Разбиваем на отдельные классы если их несколько
            ifc_parts = [x.strip() for x in ifc_lower.split('\n')]
            
            for ifc_part in ifc_parts:
                if ifc_part and ifc_part not in ['не моделируется', 'чаще всего не моделируется', '-', '']:
                    if ifc_part not in works_by_ifc:
                        works_by_ifc[ifc_part] = df_works_valid[df_works_valid['IFC класс'].str.contains(ifc_class, na=False)]
    
    # Работы без привязки к IFC классу (универсальные)
    universal_mask = df_works_valid['IFC класс'].isin(['не моделируется', 'чаще всего не моделируется', '-', '']) | df_works_valid['IFC класс'].isna()
    works_universal = df_works_valid[universal_mask].copy()
    
    return df_works_valid, works_by_ifc, works_universal


def map_elements_to_works(session_folder=None):
    """Функция для подбора видов работ к элементам (точка входа для Flask)"""
    try:
        result = main(session_folder)
        return result
    except Exception as e:
        print(f"Ошибка при маппинге элементов к работам: {e}")
        return {'success': False, 'error': str(e)}


def main(session_folder=None):
    # Пути к файлам - используем аргументы командной строки или значения по умолчанию
    if len(sys.argv) >= 4:
        elements_file = sys.argv[1]
        works_file = sys.argv[2]
        output_file = sys.argv[3]
    elif session_folder:
        # Используем папку сессии
        elements_file = os.path.join(session_folder, 'full_elements.xlsx')
        works_file = os.path.join(session_folder, 'Перечень работ КР_new.xlsx')
        output_file = os.path.join(session_folder, 'mapped_elements_works.xlsx')
    else:
        # Значения по умолчанию
        elements_file = 'uploads/b43c911d-191d-42e5-a953-810feb9bf2de/full_elements.xlsx'
        works_file = 'uploads/b43c911d-191d-42e5-a953-810feb9bf2de/Перечень работ КР_new.xlsx'
        output_file = 'uploads/b43c911d-191d-42e5-a953-810feb9bf2de/mapped_elements_works.xlsx'
    
    # Проверяем существование файлов
    if not os.path.exists(elements_file):
        error_msg = f"Ошибка: файл элементов не найден: {elements_file}"
        print(error_msg)
        return {'success': False, 'error': error_msg}
    
    if not os.path.exists(works_file):
        error_msg = f"Ошибка: файл работ не найден: {works_file}"
        print(error_msg)
        return {'success': False, 'error': error_msg}
    
    print("=" * 60)
    print("Подбор видов работ к элементам (версия 2)")
    print("=" * 60)
    
    # Загружаем элементы
    print(f"\nЗагрузка таблицы элементов из {elements_file}...")
    df_elements = pd.read_excel(elements_file)
    print(f"Элементы: {len(df_elements)} строк")
    
    # Загружаем и подготавливаем работы
    df_works_valid, works_by_ifc, works_universal = load_and_prepare_works(works_file)
    
    # Обрабатываем элементы
    results = []
    elements_with_works = set()
    elements_without_works = set()
    
    print("\nПодбор видов работ для элементов...")
    for idx, element_row in df_elements.iterrows():
        if idx % 500 == 0:
            print(f"Обработано {idx} из {len(df_elements)} элементов...")
        
        element_params = get_element_parameters(element_row)
        
        # Если элемент без работ, пропускаем его
        if element_params['no_works']:
            elements_without_works.add(idx)
            continue
        
        # Находим подходящие работы
        matched_works = find_works_for_element(
            element_row, df_works_valid, works_by_ifc, works_universal
        )
        
        if matched_works:
            elements_with_works.add(idx)
            for work in matched_works:
                result = {
                    'Element_Index': idx,
                    'Element_Name': element_row.get('Наименование', ''),
                    'Element_Material': element_row.get('Материал', ''),
                    'Element_Characteristics': element_row.get('Характеристики материала', ''),
                    'Element_Level': element_params['level'],
                    'IFC_Class': element_params['ifc_class'],
                    'Concrete_Grade': element_params['concrete_grade'],
                    'Freeze_Durability': element_params['freeze_durability'],
                    'Water_Resist': element_params['water_resist'],
                    'Volume_m3': element_row.get('Объем, м³', ''),
                    'Thickness_m': element_row.get('Толщина, м', ''),
                    'Width_m': element_row.get('Ширина, м', ''),
                    'Height_m': element_row.get('Высота, м', ''),
                    'Length_m': element_row.get('Длина, м', ''),
                    'Area_m2': element_row.get('Площадь, м²', ''),
                    'Work_Name': work['work_name'],
                    'Work_Unit': work['unit'],
                    'TSN_Code': work['tsn_code'],
                    'Work_Description': work['description'],
                    'Formula': work['formula'],
                    'Parametrization': work['parametrization'],
                }
                results.append(result)
        else:
            # Элемент не имеет подходящих работ (нет в таблице работ)
            elements_without_works.add(idx)
    
    # Создаем DataFrame с результатами
    df_results = pd.DataFrame(results)
    
    print(f"\n{'=' * 60}")
    print(f"Результаты подбора видов работ")
    print(f"{'=' * 60}")
    print(f"Найдено {len(df_results)} соответствий элемент-работа")
    print(f"Уникальных элементов с работами: {len(elements_with_works)}")
    print(f"Элементов без работ (исключены): {len(elements_without_works)}")
    
    # Сохраняем результат
    df_results.to_excel(output_file, index=False)
    print(f"\nРезультаты сохранены в файл: {output_file}")
    
    # Выводим статистику
    print("\n" + "=" * 60)
    print("Статистика по типам элементов (IFC класс)")
    print("=" * 60)
    if len(df_results) > 0:
        ifc_counts = df_results['IFC_Class'].value_counts()
        for ifc_class, count in ifc_counts.items():
            print(f"{ifc_class}: {count} работ")
    
    print("\n" + "=" * 60)
    print("Статистика по уровням элементов")
    print("=" * 60)
    if len(df_results) > 0:
        level_counts = df_results['Element_Level'].value_counts()
        for level, count in level_counts.items():
            print(f"{level}: {count} работ")
    
    print("\n" + "=" * 60)
    print("Примеры результатов (первые 10)")
    print("=" * 60)
    if len(df_results) > 0:
        print(df_results[['Element_Name', 'IFC_Class', 'Element_Level', 'Concrete_Grade', 'Work_Name']].head(10).to_string())
    
    print("\n" + "=" * 60)
    print("Элементы без работ (первые 20)")
    print("=" * 60)
    if len(elements_without_works) > 0:
        without_works_list = list(elements_without_works)[:20]
        for idx in without_works_list:
            element_row = df_elements.iloc[idx]
            params = get_element_parameters(element_row)
            print(f"[{idx}] {element_row.get('Наименование', '')[:80]} | IFC: {params['ifc_class']} | Уровень: {params['level']}")
    
    print("\n" + "=" * 60)
    print("Работа завершена успешно!")
    print("=" * 60)
    
    return {
        'success': True,
        'output_file': os.path.basename(output_file),
        'total_elements': len(df_elements),
        'matched_elements': len(elements_with_works),
        'elements_without_works': len(elements_without_works),
        'total_matches': len(df_results)
    }


if __name__ == '__main__':
    import sys
    
    if len(sys.argv) >= 2:
        session_folder = sys.argv[1]
        result = main(session_folder)
        if not result.get('success'):
            sys.exit(1)
    else:
        print("Использование: python map_elements_to_works.py <session_folder>")
        print("Пример: python map_elements_to_works.py /workspace/uploads/session_id")
        # Для обратной совместимости запускаем без аргументов
        result = main()
        if result and not result.get('success'):
            sys.exit(1)
