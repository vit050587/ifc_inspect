#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
IFC Extractor and Analyzer
Извлекает данные из IFC файла, рассчитывает объемы через геометрию,
классифицирует элементы, парсит свойства бетона и формирует сводную таблицу.
"""

import ifcopenshell
import ifcopenshell.util.element
import ifcopenshell.util.unit
import json
import pandas as pd
import re
import os
from collections import defaultdict
from typing import Dict, List, Any, Optional, Tuple

# Пути к файлам
IFC_FILE = "data/ifc модель КР.ifc"
JSON_OUTPUT = "data/ifc_extracted_detailed.json"
EXCEL_OUTPUT = "Data/summary.xlsx"

# Целевые значения для сравнения (из спецификации)
TARGET_VALUES = {
    "Стены": 3857.89,
    "Перекрытия": 4294.95,
    "Колонны": 94.28,
    "Фундаменты_плиты": 3042.14,
    "Бетон_неармированный": 619.08,
    "Лестничные_марши": 54.35,  # 39.13 + 13.15 + 1.55 + 0.52
    "Межэтажные_площадки": 70.08  # 34.74 + 2.54 + 24.79 + 8.01
}

def get_unit_scale(ifc_file: ifcopenshell.file) -> float:
    """Получить масштаб единиц из IFC файла."""
    # Пытаемся получить единицы измерения разными способами
    scale = 1.0  # По умолчанию метры
    
    try:
        # Способ 1: Через units в contexts
        contexts = ifc_file.by_type('IfcGeometricRepresentationContext')
        for context in contexts:
            if hasattr(context, 'Precision') and context.Precision:
                # Precision обычно около 1e-5 для метров
                if context.Precision < 1e-4:
                    scale = 1.0
                elif context.Precision < 1e-2:
                    scale = 0.001  # мм
        
        # Способ 2: Проверка через IfcUnitAssignment
        if hasattr(ifc_file, 'by_type'):
            unit_assignments = ifc_file.by_type('IfcUnitAssignment')
            for ua in unit_assignments:
                if hasattr(ua, 'Units'):
                    for unit in ua.Units:
                        if hasattr(unit, 'UnitType') and unit.UnitType == 'LENGTHUNIT':
                            if hasattr(unit, 'Name'):
                                if unit.Name == 'METRE':
                                    scale = 1.0
                                elif unit.Name == 'MILLIMETRE':
                                    scale = 0.001
                                elif unit.Name == 'CENTIMETRE':
                                    scale = 0.01
                            break
    except Exception as e:
        print(f"Предупреждение при определении единиц: {e}, используем метры по умолчанию")
    
    return scale

def calculate_volume_from_geometry(entity: ifcopenshell.entity_instance, scale: float) -> Optional[float]:
    """Рассчитать объем элемента через его геометрию."""
    try:
        shape = entity.Representation
        if not shape:
            return None
        
        volume = 0.0
        if hasattr(shape, 'Representations'):
            for rep in shape.Representations:
                if rep.RepresentationType == 'Solid':
                    for item in rep.Items:
                        if item.is_a('IfcExtrudedAreaSolid'):
                            # Объем = площадь профиля * глубина экструзии
                            profile = item.SweptArea
                            if profile.is_a('IfcArbitraryClosedProfileDef'):
                                area = calculate_profile_area(profile, scale)
                                depth = item.Depth * scale
                                volume += area * depth
                            elif profile.is_a('IfcRectangleProfileDef'):
                                width = profile.XDim * scale if profile.XDim else 0
                                height = profile.YDim * scale if profile.YDim else 0
                                area = width * height
                                depth = item.Depth * scale
                                volume += area * depth
                        elif item.is_a('IfcBooleanResult'):
                            # Для булевых операций пробуем приблизительный расчет
                            pass
                        elif item.is_a('IfcMappedItem'):
                            # Обработка mapped items
                            pass
        
        # Если не удалось рассчитать через Solid, пробуем bounding box
        if volume == 0.0:
            bbox = get_bounding_box(entity, scale)
            if bbox:
                volume = bbox['dx'] * bbox['dy'] * bbox['dz']
        
        return volume if volume > 0 else None
    except Exception as e:
        return None

def calculate_profile_area(profile: ifcopenshell.entity_instance, scale: float) -> float:
    """Рассчитать площадь профиля."""
    if profile.is_a('IfcArbitraryClosedProfileDef'):
        # Упрощенный расчет через координаты точек
        points = profile.OuterCurve.Points
        if points:
            coords = [(p.Coordinates[0] * scale, p.Coordinates[1] * scale) for p in points]
            # Формула площади многоугольника
            area = 0.0
            n = len(coords)
            for i in range(n):
                j = (i + 1) % n
                area += coords[i][0] * coords[j][1]
                area -= coords[j][0] * coords[i][1]
            return abs(area) / 2.0
    return 0.0

def get_bounding_box(entity: ifcopenshell.entity_instance, scale: float) -> Optional[Dict[str, float]]:
    """Получить ограничивающий короб элемента."""
    try:
        shape = entity.Representation
        if not shape:
            return None
        
        min_x, min_y, min_z = float('inf'), float('inf'), float('inf')
        max_x, max_y, max_z = float('-inf'), float('-inf'), float('-inf')
        
        def process_vertex(vertex):
            nonlocal min_x, min_y, min_z, max_x, max_y, max_z
            coords = vertex.Coordinates
            x, y, z = coords[0] * scale, coords[1] * scale, coords[2] * scale
            min_x = min(min_x, x)
            min_y = min(min_y, y)
            min_z = min(min_z, z)
            max_x = max(max_x, x)
            max_y = max(max_y, y)
            max_z = max(max_z, z)
        
        def traverse_geometry(obj):
            if obj.is_a('IfcVertexPoint'):
                process_vertex(obj)
            elif hasattr(obj, 'Coordinates'):
                process_vertex(obj)
            elif hasattr(obj, 'Points'):
                for p in obj.Points:
                    traverse_geometry(p)
            elif hasattr(obj, 'OuterCurve') and obj.OuterCurve:
                traverse_geometry(obj.OuterCurve)
            elif hasattr(obj, 'Items'):
                for item in obj.Items:
                    traverse_geometry(item)
            elif hasattr(obj, 'SweptArea'):
                traverse_geometry(obj.SweptArea)
            elif hasattr(obj, 'Position'):
                traverse_geometry(obj.Position)
        
        if hasattr(shape, 'Representations'):
            for rep in shape.Representations:
                if hasattr(rep, 'Items'):
                    for item in rep.Items:
                        traverse_geometry(item)
        
        if min_x == float('inf'):
            return None
        
        return {
            'dx': max_x - min_x,
            'dy': max_y - min_y,
            'dz': max_z - min_z,
            'min_x': min_x, 'min_y': min_y, 'min_z': min_z,
            'max_x': max_x, 'max_y': max_y, 'max_z': max_z
        }
    except Exception:
        return None

def parse_concrete_properties(name: str, properties: Dict[str, Any]) -> Dict[str, str]:
    """Распарсить свойства бетона из имени элемента и свойств IFC."""
    result = {
        'material': 'Бетон',
        'class': '',
        'frost_resistance': '',
        'waterproofing': '',
        'reinforced': True  # По умолчанию армированный
    }
    
    # Паттерны для поиска в имени
    patterns = {
        'class': [r'В(\d+)', r'B(\d+)', r'М(\d+)'],
        'frost_resistance': [r'F(\d+)', r'f(\d+)'],
        'waterproofing': [r'W(\d+)', r'w(\d+)'],
        'unreinforced': [r'неармированн', r'без арматуры', r'В7\.5', r'В5', r'M50', r'M75']
    }
    
    # Поиск в имени
    name_upper = name.upper()
    for pattern in patterns['class']:
        match = re.search(pattern, name_upper)
        if match:
            result['class'] = f"В{match.group(1)}"
            break
    
    for pattern in patterns['frost_resistance']:
        match = re.search(pattern, name_upper)
        if match:
            result['frost_resistance'] = f"F{match.group(1)}"
            break
    
    for pattern in patterns['waterproofing']:
        match = re.search(pattern, name_upper)
        if match:
            result['waterproofing'] = f"W{match.group(1)}"
            break
    
    for pattern in patterns['unreinforced']:
        if re.search(pattern, name_upper):
            result['reinforced'] = False
            break
    
    # Поиск в свойствах IFC
    for prop_name, prop_value in properties.items():
        prop_str = str(prop_value).upper()
        if not result['class']:
            for pattern in patterns['class']:
                match = re.search(pattern, prop_str)
                if match:
                    result['class'] = f"В{match.group(1)}"
                    break
        if not result['frost_resistance']:
            for pattern in patterns['frost_resistance']:
                match = re.search(pattern, prop_str)
                if match:
                    result['frost_resistance'] = f"F{match.group(1)}"
                    break
        if not result['waterproofing']:
            for pattern in patterns['waterproofing']:
                match = re.search(pattern, prop_str)
                if match:
                    result['waterproofing'] = f"W{match.group(1)}"
                    break
        if 'НЕАРМИРОВАНН' in prop_str or 'В7.5' in prop_str or 'B7.5' in prop_str:
            result['reinforced'] = False
    
    return result

def classify_slab(element: ifcopenshell.entity_instance, name: str) -> str:
    """Классифицировать плиту как фундаментную или перекрытие."""
    name_lower = name.lower()
    
    # Проверка по имени
    if any(keyword in name_lower for keyword in ['фп', 'фундамент', 'foundation', 'base']):
        return 'Фундаменты_плиты'
    elif any(keyword in name_lower for keyword in ['перекрытие', 'slab', 'floor', 'deck']):
        return 'Перекрытия'
    
    # Проверка по уровню Z (фундаменты обычно ниже)
    try:
        placement = element.ObjectPlacement
        if placement and placement.is_a('IfcLocalPlacement'):
            axis = placement.RelativePlacement.Location.Coordinates
            z_coord = axis[2] if len(axis) > 2 else 0
            # Если Z < -1.0 м, считаем фундаментом (настройка порога может потребоваться)
            if z_coord < -1.0:
                return 'Фундаменты_плиты'
    except Exception:
        pass
    
    # По умолчанию считаем перекрытием
    return 'Перекрытия'

def extract_ifc_data(ifc_path: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """Извлечь все данные из IFC файла."""
    ifc_file = ifcopenshell.open(ifc_path)
    scale = get_unit_scale(ifc_file)
    
    elements = []
    stats = defaultdict(lambda: {'count': 0, 'volume': 0.0, 'area': 0.0})
    
    # Получаем все строительные элементы разными способами
    all_elements = []
    
    # Способ 1: Через by_type для всех типов строительных элементов
    building_element_types = [
        'IfcWall', 'IfcColumn', 'IfcBeam', 'IfcSlab', 'IfcFooting', 
        'IfcStair', 'IfcMember', 'IfcPlate', 'IfcBuildingElementProxy'
    ]
    
    for elem_type in building_element_types:
        try:
            elems = ifc_file.by_type(elem_type)
            all_elements.extend(elems)
        except Exception:
            pass
    
    # Способ 2: Если есть traverse с аргументом
    if not all_elements:
        try:
            root_elements = ifc_file.by_type('IfcRoot')
            for root in root_elements:
                traversed = ifc_file.traverse(root)
                all_elements.extend(traversed)
        except Exception:
            # Если ничего не работает, берем все сущности
            all_elements = list(ifc_file)
    
    # Убираем дубликаты по ID
    seen_ids = set()
    unique_elements = []
    for elem in all_elements:
        if hasattr(elem, 'id') and elem.id() not in seen_ids:
            seen_ids.add(elem.id())
            unique_elements.append(elem)
    
    all_elements = unique_elements
    
    print(f"Всего найдено строительных элементов: {len(all_elements)}")
    
    for entity in all_elements:
        # Проверяем является ли элемент строительным (прямое или наследуемое)
        if not (entity.is_a('IfcBuildingElement') or 
                entity.is_a('IfcWall') or 
                entity.is_a('IfcColumn') or 
                entity.is_a('IfcBeam') or 
                entity.is_a('IfcSlab') or 
                entity.is_a('IfcFooting') or 
                entity.is_a('IfcStair') or 
                entity.is_a('IfcMember') or 
                entity.is_a('IfcPlate') or 
                entity.is_a('IfcBuildingElementProxy')):
            continue
        
        element_type = entity.is_a()
        name = getattr(entity, 'Name', '') or ''
        tag = getattr(entity, 'Tag', '') or ''
        description = getattr(entity, 'Description', '') or ''
        
        # Получаем свойства элемента
        properties = {}
        psets = ifcopenshell.util.element.get_psets(entity)
        for pset_name, props in psets.items():
            for prop_name, prop_value in props.items():
                properties[f"{pset_name}.{prop_name}"] = prop_value
        
        # Свойства материала
        material = ifcopenshell.util.element.get_material(entity)
        material_name = material.Name if material else ''
        
        # Расчет объема
        volume = None
        
        # Сначала пробуем взять из Qto_*Quantities (приоритет)
        for pset_name, props in psets.items():
            if pset_name.startswith('Qto_') and 'NetVolume' in props:
                vol_val = props['NetVolume']
                if isinstance(vol_val, (int, float)) and vol_val > 0:
                    # Проверяем порядок величины - если > 1000, вероятно в дм³
                    if vol_val > 1000:
                        volume = vol_val / 1000.0  # Конвертируем из дм³ в м³
                    else:
                        volume = vol_val
                    break
        
        # Если нет, пробуем Pset_ElementCommon
        if volume is None:
            if 'Pset_ElementCommon' in psets and 'NetVolume' in psets['Pset_ElementCommon']:
                vol_val = psets['Pset_ElementCommon']['NetVolume']
                if isinstance(vol_val, (int, float)) and vol_val > 0:
                    volume = vol_val
        
        # Если нет или объем подозрительно мал, считаем через геометрию
        if volume is None or volume < 0.001:
            calc_volume = calculate_volume_from_geometry(entity, scale)
            if calc_volume and calc_volume > 0.001:
                volume = calc_volume
        
        # Если все еще нет объема, используем bounding box
        if volume is None or volume < 0.001:
            bbox = get_bounding_box(entity, scale)
            if bbox:
                volume = bbox['dx'] * bbox['dy'] * bbox['dz']
        
        # Расчет площади поверхности (упрощенно через bounding box)
        area = None
        bbox = get_bounding_box(entity, scale)
        if bbox:
            area = 2 * (bbox['dx'] * bbox['dy'] + bbox['dx'] * bbox['dz'] + bbox['dy'] * bbox['dz'])
        
        # Парсинг свойств бетона
        full_name = f"{name} {description}"
        concrete_props = parse_concrete_properties(full_name, properties)
        
        # Классификация типа элемента
        category = element_type
        if element_type == 'IfcColumn':
            category = 'Колонны'
        elif element_type == 'IfcWall':
            category = 'Стены'
        elif element_type == 'IfcSlab':
            category = classify_slab(entity, full_name)
        elif element_type == 'IfcStair':
            category = 'Лестницы'
        elif element_type == 'IfcBeam':
            category = 'Балки'
        elif element_type == 'IfcFooting':
            category = 'Фундаменты_прочие'
        elif element_type == 'IfcMember':
            # Проверяем, не является ли это колонной
            if 'колонн' in full_name.lower() or 'column' in full_name.lower():
                category = 'Колонны'
            else:
                category = 'Прочие_элементы'
        
        # Пропускаем элементы без объема
        if volume is None or volume < 0.001:
            continue
        
        element_data = {
            'id': entity.id(),
            'type': element_type,
            'category': category,
            'name': name,
            'tag': tag,
            'description': description,
            'material': material_name or concrete_props['material'],
            'concrete_class': concrete_props['class'],
            'frost_resistance': concrete_props['frost_resistance'],
            'waterproofing': concrete_props['waterproofing'],
            'reinforced': concrete_props['reinforced'],
            'volume_m3': round(volume, 4),
            'area_m2': round(area, 4) if area else None,
            'properties': properties
        }
        
        elements.append(element_data)
        
        # Обновляем статистику
        key = f"{category}|{element_data['material']}|{element_data['concrete_class']}|{element_data['frost_resistance']}|{element_data['waterproofing']}|{element_data['reinforced']}"
        stats[key]['count'] += 1
        stats[key]['volume'] += volume
        if area:
            stats[key]['area'] += area
    
    return elements, dict(stats)

def create_summary_table(stats: Dict[str, Any]) -> pd.DataFrame:
    """Создать сводную таблицу из статистики."""
    rows = []
    
    for key, data in stats.items():
        parts = key.split('|')
        row = {
            'Категория': parts[0],
            'Материал': parts[1],
            'Класс_бетона': parts[2],
            'Морозостойкость': parts[3],
            'Водонепроницаемость': parts[4],
            'Армированный': 'Да' if parts[5] == 'True' else 'Нет',
            'Количество': data['count'],
            'Объем_м3': round(data['volume'], 4),
            'Площадь_м2': round(data['area'], 4) if data['area'] > 0 else None
        }
        rows.append(row)
    
    df = pd.DataFrame(rows)
    
    # Сортировка по категории и объему
    df = df.sort_values(['Категория', 'Объем_м3'], ascending=[True, False])
    
    return df

def add_comparison_with_targets(df: pd.DataFrame) -> pd.DataFrame:
    """Добавить сравнение с целевыми значениями."""
    # Группировка по категориям
    category_totals = df.groupby('Категория')['Объем_м3'].sum().to_dict()
    
    comparison = []
    for category, target in TARGET_VALUES.items():
        actual = category_totals.get(category, 0.0)
        diff = actual - target
        percent = (actual / target * 100) if target > 0 else 0
        comparison.append({
            'Категория': category,
            'Целевое_значение': target,
            'Фактическое_значение': round(actual, 4),
            'Разница': round(diff, 4),
            'Процент_от_цели': round(percent, 2)
        })
    
    comparison_df = pd.DataFrame(comparison)
    return comparison_df

def main():
    print(f"Открытие IFC файла: {IFC_FILE}")
    
    if not os.path.exists(IFC_FILE):
        print(f"Ошибка: Файл {IFC_FILE} не найден!")
        return
    
    # Извлечение данных
    elements, stats = extract_ifc_data(IFC_FILE)
    
    print(f"Извлечено элементов: {len(elements)}")
    
    # Сохранение в JSON
    output_data = {
        'metadata': {
            'source_file': IFC_FILE,
            'total_elements': len(elements),
            'extraction_date': pd.Timestamp.now().isoformat()
        },
        'elements': elements,
        'statistics': stats
    }
    
    with open(JSON_OUTPUT, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)
    
    print(f"Данные сохранены в JSON: {JSON_OUTPUT}")
    
    # Создание сводной таблицы
    df = create_summary_table(stats)
    
    # Добавление сравнения с целевыми значениями
    comparison_df = add_comparison_with_targets(df)
    
    # Сохранение в Excel с несколькими листами
    with pd.ExcelWriter(EXCEL_OUTPUT, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Детализация', index=False)
        comparison_df.to_excel(writer, sheet_name='Сравнение_с_целью', index=False)
        
        # Сводная таблица по категориям
        category_summary = df.groupby('Категория').agg({
            'Количество': 'sum',
            'Объем_м3': 'sum',
            'Площадь_м2': 'sum'
        }).reset_index()
        category_summary.to_excel(writer, sheet_name='Сводка_по_категориям', index=False)
    
    print(f"Сводная таблица сохранена в: {EXCEL_OUTPUT}")
    
    # Вывод результатов в консоль
    print("\n" + "="*80)
    print("СВОДКА ПО КАТЕГОРИЯМ")
    print("="*80)
    print(category_summary.to_string(index=False))
    
    print("\n" + "="*80)
    print("СРАВНЕНИЕ С ЦЕЛЕВЫМИ ЗНАЧЕНИЯМИ")
    print("="*80)
    print(comparison_df.to_string(index=False))
    
    # Проверка на недостающие элементы
    print("\n" + "="*80)
    print("АНАЛИЗ РАСХОЖДЕНИЙ")
    print("="*80)
    for _, row in comparison_df.iterrows():
        if abs(row['Разница']) > 1.0:  # Если расхождение больше 1 м³
            status = "НЕДОСТАЕТ" if row['Разница'] < 0 else "ИЗБЫТОК"
            print(f"{row['Категория']}: {status} {abs(row['Разница']):.2f} м³ ({row['Процент_от_цели']:.1f}% от цели)")

if __name__ == "__main__":
    main()
