#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для подбора видов работ к элементам из таблицы full_elements.xlsx
на основе таблицы Перечень работ КР_new.xlsx

Алгоритм v5 (ИТОГОВЫЙ АЛГОРИТМ ПОДБОРА ВИДОВ РАБОТ):

1. Определение уровня элемента:
   - Извлечь позицию из полей: ExpCheck_Slab:MGE_Position, ExpCheck_Wall:MGE_Position, 
     ExpCheck_Column:MGE_Position, ExpCheck_Beam:MGE_Position, ExpCheck_Ramp:MGE_Position
   - Распарсить префикс и номер секции (например ФПм-1.1-1 → префикс=ФП, секция=1)
   - Определить уровень:
     ◦ ФП/УФП/ФО/БП/ПП+секция1 → underground
     ◦ ПП/ПР/СТ/ПЛ/РП/П+секция≥2 → above
     ◦ Для элементов без позиции: по категории (Фундамент_* → underground, остальное → above)

2. Маппинг параметров:
   • t ← thickness/Width (мм)
   • S ← area_m2/NetSideArea (м²)
   • V ← volume_m3/NetVolume (м³)
   • Ф ← диаметр арматуры (мм)
   • a ← width/сечение (мм)
   • B30/F150/W6 ← concrete_class/frost_resistance/water_permeability
   Примечание: H (Height) не используется в условиях параметризации, так как высота здания учитывается на предыдущем этапе.

3. Порядок подбора работ:
   1. Фильтр по уровню (underground/above)
   2. Фильтр по IFC классу
   3. Применение условий параметризации
   4. Группировка: целые № п/п = работы, дробные = материалы
   5. Для is_reinforced=True добавить армирование

4. Пропускать: IfcBuildingElementProxy, IfcOpeningElement, элементы с volume=0 и area=0
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


def parse_position_to_level(position: str) -> Optional[str]:
    """
    Распарсить позицию элемента и определить уровень.
    
    Примеры позиций:
    - ФПм-1.1-1 → префикс=ФП, секция=1 → underground
    - УФП-1.2-3 → префикс=УФП, секция=1 → underground  
    - ФО-1.1-2 → префикс=ФО, секция=1 → underground
    - БП-1.1-1 → префикс=БП, секция=1 → underground
    - ПП-1.1-5 → префикс=ПП, секция=1 → underground
    - ПП-2.1-1 → префикс=ПП, секция=2 → above
    - ПР-2.1-3 → префикс=ПР, секция=2 → above
    - СТ-3.1-1 → префикс=СТ, секция=3 → above
    - ПЛ-4.1-2 → префикс=ПЛ, секция=4 → above
    - РП-5.1-1 → префикс=РП, секция=5 → above
    - П-6.1-3 → префикс=П, секция=6 → above
    
    Возвращает: 'underground', 'above' или None
    """
    if pd.isna(position) or not position or not isinstance(position, str):
        return None
    
    position = position.strip()
    if not position:
        return None
    
    # Извлекаем префикс (буквы до первого дефиса или цифры)
    # Учитываем что после букв может быть строчная буква (например ФПм)
    match = re.match(r'^([А-ЯA-Z]+[а-яa-z]?)([-–]|\s|\d)', position)
    if not match:
        return None
    
    prefix = match.group(1).upper()
    
    # Нормализуем префикс: убираем суффиксы типа "м" (ФПм -> ФП)
    # Оставляем только основные буквы префикса
    normalized_prefix = re.sub(r'[А-ЯA-Z]+$', '', prefix) or prefix
    
    # Извлекаем номер секции (первое число после префикса)
    # Формат может быть: ФПм-1.1-1, ФП-1.1, ПП2-1.1 и т.д.
    section_match = re.search(r'[-–]\s*(\d+)', position)
    if not section_match:
        # Пробуем найти число сразу после префикса (без дефиса)
        section_match = re.search(r'[А-ЯA-Z]+[а-яa-z]?(\d+)', position)
        if section_match:
            section_num = int(section_match.group(1))
        else:
            return None
    else:
        section_num = int(section_match.group(1))
    
    # Определяем уровень по префиксу и номеру секции
    underground_prefixes = ['ФП', 'УФП', 'ФО', 'БП']  # Фундаментная плита, Утепленная фундаментная плита, Фундаментные основания, Блочная плита
    
    # ПП с секцией 1 = подземный, с секцией >= 2 = надземный
    if normalized_prefix == 'ПП':
        if section_num == 1:
            return 'underground'
        else:
            return 'above'
    
    # Префиксы для подземной части
    if normalized_prefix in underground_prefixes:
        return 'underground'
    
    # Префиксы для надземной части (секция >= 2 или любой этаж кроме 1)
    above_prefixes = ['ПР', 'СТ', 'ПЛ', 'РП', 'П']  # Перекрытие кровли, Стены, Перекрытие лестничное, Рабочая поверхность, Плоский элемент
    
    if normalized_prefix in above_prefixes:
        if section_num >= 2:
            return 'above'
        elif section_num == 1 and prefix != 'ПП':
            # Для не-ПП префиксов секция 1 может быть подземной
            return 'underground'
    
    return None


def get_element_level(row: pd.Series) -> str:
    """
    Определяет уровень элемента (underground/above).
    
    Алгоритм:
    1. Извлечь позицию из полей: ExpCheck_Slab:MGE_Position, ExpCheck_Wall:MGE_Position, 
       ExpCheck_Column:MGE_Position, ExpCheck_Beam:MGE_Position, ExpCheck_Ramp:MGE_Position
    2. Распарсить префикс и номер секции
    3. Определить уровень по позиции
    4. Если позиции нет: по категории (Фундамент_* → underground, остальное → above)
    """
    # Шаг 1: Пытаемся извлечь позицию из различных полей
    position = None
    position_fields = [
        'ExpCheck_Slab:MGE_Position',
        'ExpCheck_Wall:MGE_Position', 
        'ExpCheck_Column:MGE_Position',
        'ExpCheck_Beam:MGE_Position',
        'ExpCheck_Ramp:MGE_Position'
    ]
    
    for field in position_fields:
        if field in row and pd.notna(row[field]):
            position = str(row[field]).strip()
            if position:
                break
    
    # Шаг 2: Пытаемся определить уровень по позиции
    if position:
        level_from_position = parse_position_to_level(position)
        if level_from_position:
            return level_from_position
    
    # Шаг 3: Если позиции нет или не удалось определить уровень, используем категорию
    category = row.get('Категория', '')
    if pd.notna(category):
        category_str = str(category).strip()
        if category_str.startswith('Фундамент_') or 'фундамент' in category_str.lower():
            return 'underground'
    
    # По умолчанию считаем элемент надземным
    return 'above'


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
    """Извлекает параметры элемента из строки таблицы
    
    Маппинг параметров согласно алгоритму v5:
    • t ← thickness/Width (мм)
    • S ← area_m2/NetSideArea (м²)
    • V ← volume_m3/NetVolume (м³)
    • Ф ← диаметр арматуры (мм)
    • a ← width/сечение (мм)
    • B30/F150/W6 ← concrete_class/frost_resistance/water_permeability
    Примечание: H (Height) не используется в условиях параметризации, так как высота здания учитывается на предыдущем этапе.
    """
    material = row.get('Материал', '')
    characteristics = row.get('Характеристики материала', '')
    
    # Получаем основные параметры
    thickness_m = row.get('Толщина, м', float('nan'))
    width_m = row.get('Ширина, м', float('nan'))
    height_m = row.get('Высота, м', float('nan'))
    length_m = row.get('Длина, м', float('nan'))
    volume_m3 = row.get('Объем, м³', float('nan'))
    area_m2 = row.get('Площадь, м²', float('nan'))
    
    # Дополнительные поля для маппинга S и V
    net_side_area = row.get('NetSideArea', float('nan'))
    net_volume = row.get('NetVolume', float('nan'))
    
    # Применяем маппинг для S (area_m2 или NetSideArea)
    if pd.isna(area_m2) and pd.notna(net_side_area):
        area_m2 = net_side_area
    
    # Применяем маппинг для V (volume_m3 или NetVolume)
    if pd.isna(volume_m3) and pd.notna(net_volume):
        volume_m3 = net_volume
    
    # Для толщины используем Width если Thickness не указан
    if pd.isna(thickness_m) and pd.notna(width_m):
        thickness_m = width_m
    
    params = {
        'ifc_class': determine_ifc_class(row),
        'name': row.get('Наименование', ''),
        'level': get_element_level(row),
        'concrete_grade': parse_concrete_grade(material, characteristics),
        'freeze_durability': parse_freeze_durability(material, characteristics),
        'water_resist': parse_water_resist(material, characteristics),
        'aggregate_type': parse_aggregate_type(material, characteristics),
        'thickness_m': thickness_m,
        'width_m': width_m,
        'height_m': height_m,
        'length_m': length_m,
        'volume_m3': volume_m3,
        'area_m2': area_m2,
        'is_reinforced': row.get('is_reinforced', False),  # Флаг армирования
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
    """Проверяет, подходит ли вид работы к элементу
    
    Алгоритм проверки (строго по порядку):
    1. Проверяем, является ли элемент элементом без работ
    2. Проверяем IFC класс
    3. Проверяем параметры (размеры, площадь, объем, толщина) - если есть условия в параметризации
    4. Проверяем материал элемента (класс бетона, морозостойкость, водонепроницаемость, заполнитель)
    5. Семантическая фильтрация по наименованию (отсечение явно неподходящих работ)
    """
    
    # 1. Проверяем, является ли элемент элементом без работ
    if element_params['no_works']:
        return False
    
    # 2. Проверяем IFC класс
    work_ifc = str(work_row.get('IFC класс', ''))
    element_ifc = element_params['ifc_class']
    
    if not ifc_class_matches(element_ifc, work_ifc):
        return False
    
    # 3. Проверяем параметры (размеры, площадь, объем, толщина)
    # Собираем все параметрические условия кроме материала
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
            
            # Проверяем условия по диаметру арматуры (Ф)
            if 'ф' in param_line.lower() or 'ф=' in param_line:
                if not check_rebar_diameter_condition(param_line):
                    return False
        
        # 4. Проверяем материал элемента (класс бетона, морозостойкость, водонепроницаемость, заполнитель)
        # Эта проверка выполняется ПОСЛЕ всех параметрических проверок
        for param_line in param_lines:
            param_line = param_line.strip()
            if not param_line or param_line in ['-', 'nan']:
                continue
            
            # Проверяем условия по бетону и маркам (только если есть соответствующие маркеры)
            if re.search(r'[bв]\d+|f\d+|w\d+|щебень', param_line.lower()):
                if not check_concrete_condition(
                    element_params['concrete_grade'],
                    element_params['freeze_durability'],
                    element_params['water_resist'],
                    element_params['aggregate_type'],
                    param_line
                ):
                    return False
    
    # 5. Семантическая фильтрация по наименованию (отсечение явно неподходящих работ)
    elem_name = element_params['name'].lower()
    work_name = str(work_row.get('Наименование работ (целые числа) материалы для этих работ (дробные числа)', '')).lower()
    
    # Определяем, является ли элемент элементом шва/профиля/герметика
    is_joint_element = any(x in elem_name for x in ['шов', 'заполнен', 'профиль', 'шпонк', 'герметик', 'набухающ'])
    
    # Специфичное правило для швов: отбрасываем общую гидроизоляцию поверхностей, если работа не про швы
    if is_joint_element:
        # Если работа про гидроизоляцию поверхностей (стены, фундаменты, полы) и не упоминает швы/стыки
        if any(x in work_name for x in ['гидроизоляц', 'изоляция']) and \
           any(x in work_name for x in ['стен', 'фундамент', 'пол', 'перекрыт', 'кровл', 'плит']) and \
           not any(x in work_name for x in ['шов', 'стык', 'примыкан', 'деформац', 'технологич']):
            return False
    
    # Правила исключения по типам работ
    exclusion_rules = [
        # Опалубка подходит только для монолитных элементов или самой опалубки
        {'work_keywords': ['опалубка'], 'elem_keywords': ['опалубк', 'монолит', 'бетон', 'плита', 'стена', 'колонн', 'ригель', 'балк']},
        # Арматура подходит для армирования, сеток, каркасов, монолита
        {'work_keywords': ['армирован', 'арматура', 'сетк', 'каркас'], 'elem_keywords': ['арматур', 'сетк', 'каркас', 'монолит', 'бетон', 'плита', 'стена', 'колонн']},
        # Гидроизоляция поверхностей (мембраны, рулоны, праймеры) НЕ подходит для элементов "Шов", "Профиль", "Герметик"
        {'work_keywords': ['праймер', 'мембран', 'рулон', 'битум', 'геотекстиль', 'полиэтилен', 'пленк', 'стяжк', 'плинтус'], 
         'elem_keywords': ['гидроизоляц', 'шпонк', 'шов', 'герметик', 'профиль', 'набухающ', 'инъекци', 'заполнен']},
        # Утеплитель подходит только для теплоизоляции
        {'work_keywords': ['утеплен', 'теплоизоляц', 'пеноплэкс', 'минеральн', 'экструд'], 'elem_keywords': ['утеплен', 'теплоизоляц', 'пеноплэкс', 'минеральн', 'экструд']},
        # Бетонирование подходит для конструктивов
        {'work_keywords': ['бетонирован', 'укладк', 'прием', 'вибрирован'], 'elem_keywords': ['бетон', 'монолит', 'плита', 'стена', 'колонн', 'ригель', 'балк', 'лестниц', 'марш', 'площадк']},
        # Демонтаж подходит только для старых конструкций
        {'work_keywords': ['демонтаж', 'разборк', 'снятие'], 'elem_keywords': ['демонтаж', 'разборк', 'существующ', 'стар', 'аварийн']},
        # Грунт/Подготовка подходит для оснований
        {'work_keywords': ['грунт', 'засыпк', 'подготовк песчан', 'планировк'], 'elem_keywords': ['грунт', 'подготовк', 'основан', 'песок', 'щебен', 'фундамент']}
    ]
    
    for rule in exclusion_rules:
        # Если в названии работы есть ключевые слова правила
        if any(kw in work_name for kw in rule['work_keywords']):
            # Проверяем, есть ли в названии элемента хоть что-то из допустимых
            if not any(kw in elem_name for kw in rule['elem_keywords']):
                return False
    
    return True


def find_works_for_element(element_row: pd.Series, works_df: pd.DataFrame, 
                           works_underground: Dict, works_aboveground: Dict, 
                           works_universal: pd.DataFrame) -> List[Dict]:
    """
    Находит все подходящие виды работ для элемента с учетом уровня (underground/above)
    
    Алгоритм v5 (строго по порядку):
    1. Фильтр по уровню (underground/above)
    2. Фильтр по IFC классу
    3. Применение условий параметризации
    4. Группировка: целые № п/п = работы, дробные = материалы
    5. Для is_reinforced=True добавить армирование
    
    Пропускать: IfcBuildingElementProxy, IfcOpeningElement, элементы с volume=0 и area=0
    """
    
    element_params = get_element_parameters(element_row)
    
    # Пропускаем элементы без работ
    if element_params['no_works']:
        return []
    
    # Пропускаем элементы с volume=0 и area=0
    volume = element_params.get('volume_m3', float('nan'))
    area = element_params.get('area_m2', float('nan'))
    if (pd.notna(volume) and volume == 0) or (pd.notna(area) and area == 0):
        return []
    
    matched_works = []
    element_ifc = element_params['ifc_class'].lower()
    element_level = element_params['level']  # 'underground' или 'above'
    
    # Собираем подходящие работы в зависимости от уровня элемента
    candidate_works = []
    has_specific_works = False  # Флаг: есть ли специфичные работы для этого IFC класса в нужном разделе
    
    # Шаг 1: Фильтр по уровню (underground/above)
    if element_level == 'underground':
        # Для подземных элементов берем работы из подземного раздела
        if element_ifc in works_underground:
            for _, work_row in works_underground[element_ifc].iterrows():
                candidate_works.append(work_row)
                has_specific_works = True
    elif element_level == 'above':
        # Для надземных элементов берем работы из надземного раздела
        if element_ifc in works_aboveground:
            for _, work_row in works_aboveground[element_ifc].iterrows():
                candidate_works.append(work_row)
                has_specific_works = True
    
    # Добавляем универсальные работы ТОЛЬКО если нет специфичных работ для этого IFC класса
    if not has_specific_works and works_universal is not None and len(works_universal) > 0:
        for _, work_row in works_universal.iterrows():
            candidate_works.append(work_row)
    
    # Проверяем каждую кандидатуру
    for work_row in candidate_works:
        # Шаг 2-3: Проверка IFC класса и параметров выполняется в match_work_to_element
        if match_work_to_element(element_params, work_row):
            matched_works.append({
                'work_idx': work_row.name,
                'work_name': work_row.get('Наименование работ (целые числа) материалы для этих работ (дробные числа)', ''),
                'unit': work_row.get('Ед. изм', ''),
                'tsn_code': work_row.get('Шифр ТСН', ''),
                'description': work_row.get('Наименование расценки/ресурса', ''),
                'formula': work_row.get('Формула расчёта объёмов работ и расхода материалов', ''),
                'parametrization': work_row.get('Параметризация', ''),
            })
    
    # Шаг 5: Для is_reinforced=True добавить армирование
    if element_params.get('is_reinforced', False):
        # Ищем работы по армированию в соответствующем разделе
        rebar_works = get_reinforcement_works(element_params, works_underground, works_aboveground, works_universal, element_level)
        matched_works.extend(rebar_works)
    
    return matched_works


def get_reinforcement_works(element_params: Dict, works_underground: Dict, works_aboveground: Dict, 
                            works_universal: pd.DataFrame, element_level: str) -> List[Dict]:
    """
    Получает работы по армированию для элемента с is_reinforced=True
    
    Возвращает список работ по армированию
    """
    reinforcement_works = []
    
    # Ищем работы с ключевыми словами "армирован", "арматура", "сетк"
    candidate_works = []
    
    if element_level == 'underground':
        for ifc_class, df in works_underground.items():
            for _, work_row in df.iterrows():
                work_name = str(work_row.get('Наименование работ (целые числа) материалы для этих работ (дробные числа)', '')).lower()
                if any(kw in work_name for kw in ['армирован', 'арматура', 'сетк', 'каркас']):
                    candidate_works.append(work_row)
    elif element_level == 'above':
        for ifc_class, df in works_aboveground.items():
            for _, work_row in df.iterrows():
                work_name = str(work_row.get('Наименование работ (целые числа) материалы для этих работ (дробные числа)', '')).lower()
                if any(kw in work_name for kw in ['армирован', 'арматура', 'сетк', 'каркас']):
                    candidate_works.append(work_row)
    
    # Также проверяем универсальные работы
    if works_universal is not None:
        for _, work_row in works_universal.iterrows():
            work_name = str(work_row.get('Наименование работ (целые числа) материалы для этих работ (дробные числа)', '')).lower()
            if any(kw in work_name for kw in ['армирован', 'арматура', 'сетк', 'каркас']):
                candidate_works.append(work_row)
    
    # Фильтруем работы по соответствию элементу
    for work_row in candidate_works:
        if match_work_to_element(element_params, work_row):
            reinforcement_works.append({
                'work_idx': work_row.name,
                'work_name': work_row.get('Наименование работ (целые числа) материалы для этих работ (дробные числа)', ''),
                'unit': work_row.get('Ед. изм', ''),
                'tsn_code': work_row.get('Шифр ТСН', ''),
                'description': work_row.get('Наименование расценки/ресурса', ''),
                'formula': work_row.get('Формула расчёта объёмов работ и расхода материалов', ''),
                'parametrization': work_row.get('Параметризация', ''),
            })
    
    return reinforcement_works


def load_and_prepare_works(works_file: str) -> Tuple[pd.DataFrame, Dict, Dict, pd.DataFrame]:
    """
    Загружает и подготавливает таблицу работ с разделением на подземные/надземные
    
    Возвращает:
    - df_works_valid: все валидные работы
    - works_underground: работы для подземной части (словарь по IFC классам)
    - works_aboveground: работы для надземной части (словарь по IFC классам)
    - works_universal: универсальные работы (без привязки к уровню)
    """
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
    
    # Определяем текущий раздел для каждой строки
    # Разделы определяются по строкам в колонке "Категория уровня" содержащим "Подземная часть здания" или "Надземная часть здания"
    # Важно: ищем именно заголовки высокого уровня, а не подразделы
    
    current_section = None  # 'underground', 'aboveground', 'other'
    section_column = []
    
    for idx, row in df_works.iterrows():
        # Используем колонку "Категория уровня" вместо "Наименование работ"
        category = str(row.get('Категория уровня', '')) if pd.notna(row.get('Категория уровня', '')) else ''
        
        # Проверяем является ли строка заголовком основного раздела
        # Ищем именно главные разделы, а не подразделы
        if 'Подземная часть здания' in category and 'Подраздел' not in category:
            current_section = 'underground'
        elif 'Надземная часть здания' in category and 'Подраздел' not in category:
            current_section = 'aboveground'
        # Подразделы наследуют текущий раздел
        
        section_column.append(current_section)
    
    df_works_valid['_section'] = section_column[:len(df_works_valid)]
    
    # Группируем работы по IFC классам для ускорения поиска с разделением на подземные/надземные
    works_underground = {}  # работы для подземной части
    works_aboveground = {}  # работы для надземной части
    
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
                    # Фильтруем работы по IFC классу и разделяем по секциям
                    mask = df_works_valid['IFC класс'].str.contains(ifc_class, na=False)
                    
                    # Подземные работы
                    underground_mask = mask & (df_works_valid['_section'] == 'underground')
                    if underground_mask.any():
                        if ifc_part not in works_underground:
                            works_underground[ifc_part] = df_works_valid[underground_mask].copy()
                    
                    # Надземные работы
                    aboveground_mask = mask & (df_works_valid['_section'] == 'aboveground')
                    if aboveground_mask.any():
                        if ifc_part not in works_aboveground:
                            works_aboveground[ifc_part] = df_works_valid[aboveground_mask].copy()
    
    # Работы без привязки к IFC классу (универсальные) - они применяются ко всем уровням
    universal_mask = df_works_valid['IFC класс'].isin(['не моделируется', 'чаще всего не моделируется', '-', '']) | df_works_valid['IFC класс'].isna()
    works_universal = df_works_valid[universal_mask].copy()
    
    print(f"Подземных работ (по IFC): {sum(len(v) for v in works_underground.values())}")
    print(f"Надземных работ (по IFC): {sum(len(v) for v in works_aboveground.values())}")
    print(f"Универсальных работ: {len(works_universal)}")
    
    return df_works_valid, works_underground, works_aboveground, works_universal


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
    # Проверяем, что скрипт запущен напрямую из командной строки с аргументами
    if session_folder is None and len(sys.argv) >= 4:
        elements_file = sys.argv[1]
        works_file = sys.argv[2]
        output_file = sys.argv[3]
    elif session_folder:
        # Используем папку сессии (вызов из Flask)
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
    print("Подбор видов работ к элементам (версия 3)")
    print("=" * 60)
    
    # Загружаем элементы
    print(f"\nЗагрузка таблицы элементов из {elements_file}...")
    df_elements = pd.read_excel(elements_file)
    print(f"Элементы: {len(df_elements)} строк")
    
    # Загружаем и подготавливаем работы с разделением на подземные/надземные
    df_works_valid, works_underground, works_aboveground, works_universal = load_and_prepare_works(works_file)
    
    # Оптимизация: предварительно фильтруем элементы без работ
    print("\nПредварительная фильтрация элементов...")
    valid_element_indices = []
    for idx in range(len(df_elements)):
        element_row = df_elements.iloc[idx]
        params = get_element_parameters(element_row)
        if not params['no_works'] and params['ifc_class']:
            valid_element_indices.append(idx)
    
    print(f"Элементов для обработки: {len(valid_element_indices)}")
    
    # Обрабатываем элементы
    results = []
    elements_with_works = set()
    elements_without_works = set()
    
    print("\nПодбор видов работ для элементов...")
    total = len(valid_element_indices)
    for batch_idx, idx in enumerate(valid_element_indices):
        if batch_idx % 200 == 0:
            print(f"Обработано {batch_idx} из {total} элементов...")
        
        element_row = df_elements.iloc[idx]
        element_params = get_element_parameters(element_row)
        
        # Находим подходящие работы с учетом уровня элемента
        matched_works = find_works_for_element(
            element_row, df_works_valid, works_underground, works_aboveground, works_universal
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
    
    # Добавляем исключенные элементы
    for idx in range(len(df_elements)):
        if idx not in valid_element_indices:
            elements_without_works.add(idx)
    
    # Создаем DataFrame с результатами
    df_results = pd.DataFrame(results)
    
    # Удаляем дубликаты работ для каждого элемента
    # Дубликат определяется по комбинации Element_Index и Work_Name
    if len(df_results) > 0:
        initial_count = len(df_results)
        df_results = df_results.drop_duplicates(subset=['Element_Index', 'Work_Name'], keep='first')
        duplicates_removed = initial_count - len(df_results)
        if duplicates_removed > 0:
            print(f"Удалено дубликатов работ: {duplicates_removed}")
    
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
