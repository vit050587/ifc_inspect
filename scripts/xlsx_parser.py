#!/usr/bin/env python3
"""
Универсальный скрипт парсинга Excel файлов отчетов IFC (ifc_report.xlsx).
Извлекает ВСЕ материалы из таблицы с их параметрами и единицами измерения.

Особенности:
- Работает с любыми материалами (бетон, гидроизоляция, утеплитель, раствор, арматура и др.)
- Автоматически определяет тип единицы измерения (объем м³, площадь м², штуки, литры)
- Группирует данные по полному наименованию материала
- Считает количество элементов для каждого материала
- Создает сводную таблицу с агрегированными данными
"""

import os
import re
import json
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
from collections import defaultdict
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


# Mapping для основных типов элементов IFC к их Qto_ префиксам и колонкам объема/площади
IFC_ELEMENT_MAPPING = {
    'IfcWall': {'qto_prefix': 'Qto_WallBaseQuantities', 'volume_col': 'NetVolume', 'area_col': 'NetSideArea'},
    'IfcColumn': {'qto_prefix': 'Qto_ColumnBaseQuantities', 'volume_col': 'NetVolume', 'area_col': None},
    'IfcSlab': {'qto_prefix': 'Qto_SlabBaseQuantities', 'volume_col': 'NetVolume', 'area_col': 'NetArea'},
    'IfcStair': {'qto_prefix': 'Qto_StairBaseQuantities', 'volume_col': None, 'area_col': None},
    'IfcRamp': {'qto_prefix': 'Qto_RampBaseQuantities', 'volume_col': None, 'area_col': None},
    'IfcBeam': {'qto_prefix': 'Qto_BeamBaseQuantities', 'volume_col': 'NetVolume', 'area_col': None},
    'IfcFooting': {'qto_prefix': 'Qto_FootingBaseQuantities', 'volume_col': 'NetVolume', 'area_col': None},
    'IfcPile': {'qto_prefix': 'Qto_PileBaseQuantities', 'volume_col': 'NetVolume', 'area_col': None},
    'IfcRoof': {'qto_prefix': 'Qto_RoofBaseQuantities', 'volume_col': 'NetVolume', 'area_col': 'NetArea'},
    'IfcPlate': {'qto_prefix': 'Qto_PlateBaseQuantities', 'volume_col': None, 'area_col': 'NetArea'},
    'IfcMember': {'qto_prefix': 'Qto_MemberBaseQuantities', 'volume_col': None, 'area_col': None},
    'IfcCurtainWall': {'qto_prefix': 'Qto_CurtainWallBaseQuantities', 'volume_col': None, 'area_col': 'NetSideArea'},
    'IfcWindow': {'qto_prefix': 'Qto_WindowBaseQuantities', 'volume_col': None, 'area_col': None},
    'IfcDoor': {'qto_prefix': 'Qto_DoorBaseQuantities', 'volume_col': None, 'area_col': None},
    'IfcFurnishingElement': {'qto_prefix': 'Qto_FurnishingElementBaseQuantities', 'volume_col': None, 'area_col': None},
    'IfcBuildingElementProxy': {'qto_prefix': 'Qto_BuildingElementProxyBaseQuantities', 'volume_col': None, 'area_col': None},
    'IfcOpeningElement': {'qto_prefix': 'Qto_OpeningElementBaseQuantities', 'volume_col': None, 'area_col': 'Area'},
    'IfcReinforcingMesh': {'qto_prefix': 'Qto_ReinforcingMeshBaseQuantities', 'volume_col': None, 'area_col': None},
    'IfcBuildingStorey': {'qto_prefix': 'Qto_BuildingStoreyBaseQuantities', 'volume_col': None, 'area_col': None},
    'IfcBuilding': {'qto_prefix': 'Qto_BuildingBaseQuantities', 'volume_col': None, 'area_col': None},
}

# Паттерны для определения типа единицы измерения по названию материала
UNIT_PATTERNS = {
    'volume': [
        r'бетон', r'раствор', r'грунт', r'песок', r'щебень', r'керамзит',
        r'объем', r'volume', r'куб', r'м3', r'м³'
    ],
    'area': [
        r'гидроизоляц', r'пароизоляц', r'теплоизоляц', r'утеплител',
        r'мембран', r'пленк', r'покрыт', r'облицовк', r'штукатурк',
        r'площадь', r'area', r'м2', r'м²', r'квадрат'
    ],
    'length': [
        r'шнур', r'профиль', r'труба', r'кабель', r'провод', r'арматур',
        r'длина', r'length', r'погонный', r'м\.п\.', r'мп'
    ],
    'volume_liquid': [
        r'праймер', r'мастик', r'клей', r'герметик', r'жидк',
        r'литр', r'l', r'л\.'
    ],
    'pieces': [
        r'анкер', r'дюбель', r'саморез', r'болт', r'гайка', r'шайба',
        r'закладн', r'детал', r'элемент', r'издели', r'конструкц',
        r'сетк', r'каркас', r'штук', r'pcs', r'шт\.'
    ]
}

# Словарь соответствия типов материалов для группировки
MATERIAL_GROUPS = {
    'Бетон': ['бетон', 'B', 'В'],
    'Гидроизоляция': ['гидроизоляц', 'мембран', 'техноэласт', 'плантер'],
    'Утеплитель': ['утеплител', 'полистирол', 'carbon', 'пеноплэкс', 'пенопласт'],
    'Раствор': ['раствор', 'стяжка', 'М100', 'М150', 'М200'],
    'Праймер': ['праймер', 'битумный'],
    'Мастика': ['мастик', 'приклеивающ'],
    'Арматура': ['арматур', 'сетк', 'каркас', 'reinforc'],
    'Изоляция': ['пароизоляц', 'теплоизоляц'],
}


def detect_unit_type(material_name: str, has_volume: bool = False, has_area: bool = False) -> str:
    """
    Определяет тип единицы измерения для материала по его названию.
    
    Args:
        material_name: Название материала
        has_volume: Есть ли колонка объема для этого элемента
        has_area: Есть ли колонка площади для этого элемента
        
    Returns:
        Тип единицы: 'volume' (м³), 'area' (м²), 'length' (м), 
                     'volume_liquid' (л), 'pieces' (шт)
    """
    name_lower = material_name.lower()
    
    # Сначала проверяем по паттернам названия
    for unit_type, patterns in UNIT_PATTERNS.items():
        for pattern in patterns:
            if re.search(pattern, name_lower):
                return unit_type
    
    # Если не нашли по паттернам, используем доступные метрики
    if has_volume:
        return 'volume'
    if has_area:
        return 'area'
    
    # По умолчанию считаем что это штуки (для поштучных элементов)
    return 'pieces'


def get_unit_label(unit_type: str) -> str:
    """
    Возвращает текстовое обозначение единицы измерения.
    """
    labels = {
        'volume': 'м³',
        'area': 'м²',
        'length': 'м',
        'volume_liquid': 'л',
        'pieces': 'шт'
    }
    return labels.get(unit_type, 'шт')


def find_ifc_report_file(session_folder: str) -> Optional[str]:
    """Находит файл ifc_report.xlsx в папке сессии."""
    report_path = Path(session_folder) / 'ifc_report.xlsx'
    
    if not report_path.exists():
        print(f"⚠️  Файл ifc_report.xlsx не найден: {report_path}")
        return None
    
    return str(report_path)


def parse_excel_specification(excel_path: str) -> Dict[str, Any]:
    """
    Универсальный парсер Excel файла отчета IFC (лист "Элементы").
    Извлекает ВСЕ материалы из таблицы с их значениями и единицами измерения.
    """
    
    print(f"\n📊 Парсинг Excel отчета IFC: {os.path.basename(excel_path)}")
    
    try:
        # Читаем лист "Элементы"
        df = pd.read_excel(excel_path, sheet_name='Элементы', dtype=str)
        
        # Нормализуем имена колонок
        df.columns = [str(col).strip() for col in df.columns]
        
        # Ожидаемые имена колонок
        col_long_name = 'Element Specific:LongName'
        col_ifc_class = 'Ifc Class'
        col_concrete_grade = 'ExpCheck_MaterialConcrete:MGE_ConcreteGrade'
        
        # Проверяем наличие обязательных колонок
        if col_ifc_class not in df.columns:
            raise ValueError(f"Столбец '{col_ifc_class}' не найден в файле.")
        if col_long_name not in df.columns:
            col_long_name = 'Element Specific:Name'
            if col_long_name not in df.columns:
                raise ValueError(f"Столбец '{col_long_name}' не найден в файле")
        
        items = []
        skipped_rows = 0
        
        for idx, row in df.iterrows():
            ifc_class = row.get(col_ifc_class)
            if not ifc_class or pd.isna(ifc_class):
                skipped_rows += 1
                continue
            
            ifc_class = str(ifc_class).strip()
            
            # Пропускаем если класс не в маппинге
            if ifc_class not in IFC_ELEMENT_MAPPING:
                skipped_rows += 1
                continue
            
            long_name = row.get(col_long_name, '')
            if pd.isna(long_name):
                long_name = ''
            
            # Получаем марку бетона
            concrete_grade = row.get(col_concrete_grade, '')
            if pd.isna(concrete_grade) or concrete_grade == '0':
                concrete_grade = ''
            
            # Получаем конфиг для этого типа элемента
            element_config = IFC_ELEMENT_MAPPING[ifc_class]
            qto_prefix = element_config['qto_prefix']
            volume_col_name = f"{qto_prefix}:{element_config.get('volume_col')}" if element_config.get('volume_col') else None
            area_col_name = f"{qto_prefix}:{element_config.get('area_col')}" if element_config.get('area_col') else None
            
            # Получаем значения объема и площади
            volume_value = None
            area_value = None
            
            if volume_col_name and volume_col_name in df.columns:
                vol_val = row.get(volume_col_name)
                if vol_val and str(vol_val) != '0' and str(vol_val) != 'nan':
                    try:
                        volume_value = float(str(vol_val).replace(',', '.'))
                    except (ValueError, TypeError):
                        pass
            
            if area_col_name and area_col_name in df.columns:
                area_val = row.get(area_col_name)
                if area_val and str(area_val) != '0' and str(area_val) != 'nan':
                    try:
                        area_value = float(str(area_val).replace(',', '.'))
                    except (ValueError, TypeError):
                        pass
            
            # Определяем материал
            # Базовый материал из MGE_Material (если есть специфичная колонка)
            base_material_col = f'ExpCheck_{ifc_class[3:]}:MGE_Material'  # IfcWall -> ExpCheck_Wall:MGE_Material
            base_material = 'Бетон'  # значение по умолчанию
            
            if base_material_col in df.columns:
                mat_val = row.get(base_material_col)
                if mat_val and str(mat_val) != '0' and str(mat_val) != 'nan':
                    base_material = str(mat_val).strip()
            
            # Формируем полное название материала с маркой
            if concrete_grade:
                full_material = f"{base_material} {concrete_grade}"
            else:
                full_material = base_material
            
            # Определяем группу материала
            material_group = "Другой"
            for group, keywords in MATERIAL_GROUPS.items():
                for keyword in keywords:
                    if keyword.lower() in full_material.lower():
                        material_group = group
                        break
                if material_group != "Другой":
                    break
            
            # Добавляем запись с объемом (если есть)
            if volume_value is not None and volume_value > 0:
                unit_type = detect_unit_type(full_material, has_volume=True, has_area=False)
                items.append({
                    'ifc_class': ifc_class,
                    'element_name': long_name,
                    'material_full': full_material,
                    'material_group': material_group,
                    'unit_type': unit_type,
                    'unit_label': get_unit_label(unit_type),
                    'value': volume_value,
                    'metric_type': 'volume'
                })
            
            # Добавляем запись с площадью (если есть)
            if area_value is not None and area_value > 0:
                unit_type = detect_unit_type(full_material, has_volume=False, has_area=True)
                items.append({
                    'ifc_class': ifc_class,
                    'element_name': long_name,
                    'material_full': full_material,
                    'material_group': material_group,
                    'unit_type': unit_type,
                    'unit_label': get_unit_label(unit_type),
                    'value': area_value,
                    'metric_type': 'area'
                })
            
            # Если нет ни объема ни площади - добавляем как штуку
            if volume_value is None and area_value is None:
                items.append({
                    'ifc_class': ifc_class,
                    'element_name': long_name,
                    'material_full': full_material,
                    'material_group': material_group,
                    'unit_type': 'pieces',
                    'unit_label': 'шт',
                    'value': 1.0,
                    'metric_type': 'pieces'
                })
        
        print(f"   ✅ Извлечено {len(items)} записей о материалах (пропущено {skipped_rows} нерелевантных строк)")
        
        return {
            'success': True,
            'items': items,
            'total_items': len(items),
        }
        
    except Exception as e:
        print(f"   ❌ Ошибка при парсинге Excel: {e}")
        import traceback
        traceback.print_exc()
        return {'success': False, 'error': str(e), 'items': [], 'total_items': 0}


def aggregate_materials_data(items: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    """
    Агрегирует данные по полному наименованию материала + типу единицы.
    """
    
    aggregated = {}
    
    for item in items:
        material = item['material_full']
        unit_type = item['unit_type']
        key = f"{material}|{unit_type}"
        
        if key not in aggregated:
            aggregated[key] = {
                'material_full': material,
                'material_group': item['material_group'],
                'unit_type': unit_type,
                'unit_label': item['unit_label'],
                'total_value': 0.0,
                'element_count': 0,
                'elements_list': []
            }
        
        aggregated[key]['total_value'] += item['value']
        aggregated[key]['element_count'] += 1
        
        # Сохраняем пример имени элемента
        if len(aggregated[key]['elements_list']) < 3:
            elem_name = item.get('element_name', '')
            if elem_name and elem_name not in aggregated[key]['elements_list']:
                aggregated[key]['elements_list'].append(elem_name)
    
    return aggregated


def create_aggregated_excel(aggregated_data: Dict[str, Dict[str, Any]], output_path: str) -> str:
    """Создаёт Excel отчёт с агрегированными данными."""
    
    print(f"\n📈 Создание Excel отчета: {output_path}")
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Сводка по материалам"
    
    headers = [
        "№", "Материал", "Группа", "Ед. изм.",
        "Общее количество", "Кол-во элементов", "Примеры элементов"
    ]
    
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    
    # Пишем заголовки
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    
    # Сортируем: сначала объем/площадь, потом штуки
    sorted_items = sorted(
        aggregated_data.items(),
        key=lambda x: (
            0 if x[1]['unit_type'] in ['volume', 'area', 'length', 'volume_liquid'] else 1,
            x[1]['material_group'],
            x[1]['material_full']
        )
    )
    
    # Пишем данные
    row_num = 2
    for idx, (key, data) in enumerate(sorted_items, 1):
        examples = "; ".join(data['elements_list'][:3])
        if len(data['elements_list']) > 3:
            examples += f" ... и ещё {len(data['elements_list']) - 3}"
        
        ws.cell(row=row_num, column=1, value=idx).border = thin_border
        ws.cell(row=row_num, column=2, value=data['material_full']).border = thin_border
        ws.cell(row=row_num, column=3, value=data['material_group']).border = thin_border
        ws.cell(row=row_num, column=4, value=data['unit_label']).border = thin_border
        ws.cell(row=row_num, column=5, value=round(data['total_value'], 3)).border = thin_border
        ws.cell(row=row_num, column=6, value=data['element_count']).border = thin_border
        ws.cell(row=row_num, column=7, value=examples).border = thin_border
        row_num += 1
    
    # Автоширина колонок
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_len + 2, 60)
    
    wb.save(output_path)
    print(f"   ✅ Excel сохранён: {output_path}")
    return output_path


def parse_and_aggregate_specification(session_folder: str) -> Dict[str, Any]:
    """Основная функция: поиск, парсинг, агрегация и сохранение."""
    
    print("\n" + "="*60)
    print("📊 ПАРСИНГ ОТЧЕТА IFC (ifc_report.xlsx)")
    print("="*60)
    
    results = {
        'success': False, 
        'excel_file_found': None, 
        'parsed_items': 0,
        'aggregated_materials': 0, 
        'output_file': None,
        'materials_summary': {}
    }
    
    # Шаг 1: Найти файл ifc_report.xlsx
    excel_path = find_ifc_report_file(session_folder)
    if not excel_path:
        print("❌ Файл ifc_report.xlsx не найден")
        return results
    results['excel_file_found'] = excel_path
    
    # Шаг 2: Распарсить Excel
    parse_result = parse_excel_specification(excel_path)
    if not parse_result['success']:
        results['error'] = parse_result.get('error')
        return results
    results['parsed_items'] = parse_result['total_items']
    
    # Шаг 3: Агрегировать данные по материалам
    aggregated_data = aggregate_materials_data(parse_result['items'])
    results['aggregated_materials'] = len(aggregated_data)
    
    # Подготовка сводки для JSON
    materials_summary = {}
    for key, data in aggregated_data.items():
        material_name = data['material_full']
        if material_name not in materials_summary:
            materials_summary[material_name] = {
                'group': data['material_group'],
                'values': {}
            }
        materials_summary[material_name]['values'][data['unit_type']] = {
            'value': round(data['total_value'], 3),
            'unit': data['unit_label'],
            'element_count': data['element_count']
        }
    results['materials_summary'] = materials_summary
    
    # Шаг 4: Создать выходной Excel файл
    output_path = Path(session_folder) / "materials_summary.xlsx"
    create_aggregated_excel(aggregated_data, str(output_path))
    results['output_file'] = str(output_path)
    results['success'] = True
    
    # Шаг 5: Сохранить JSON для машинной обработки
    json_path = Path(session_folder) / "materials_summary.json"
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump({
            'success': True,
            'source_file': os.path.basename(excel_path),
            'total_items_parsed': parse_result['total_items'],
            'total_materials': len(aggregated_data),
            'output_excel': 'materials_summary.xlsx',
            'materials': materials_summary
        }, f, ensure_ascii=False, indent=2)
    
    print(f"\n💾 Данные сохранены в: {json_path}")
    
    # Сводка в консоль
    print("\n" + "="*60)
    print("📊 СВОДКА ПО МАТЕРИАЛАМ:")
    print("="*60)
    
    by_unit = defaultdict(list)
    for key, data in aggregated_data.items():
        by_unit[data['unit_label']].append((data['material_full'], data['total_value'], data['element_count']))
    
    for unit_label in ['м³', 'м²', 'м', 'л', 'шт']:
        if unit_label in by_unit:
            print(f"\n   [{unit_label}]:")
            for mat_name, total_val, elem_count in sorted(by_unit[unit_label], key=lambda x: -x[1]):
                print(f"      • {mat_name}: {total_val:.3f} {unit_label} ({elem_count} эл.)")
    
    return results


def main(session_folder: str):
    """Основная функция парсинга отчета IFC."""
    print("="*60)
    print("📊 ПАРСИНГ ОТЧЕТА IFC (ifc_report.xlsx)")
    print("="*60)
    print(f"Папка сессии: {session_folder}")
    print("="*60)
    
    results = parse_and_aggregate_specification(session_folder)
    
    return results


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Использование: python xlsx_parser.py <session_folder>")
        print("Пример: python xlsx_parser.py /workspace/uploads/session_id")
        sys.exit(1)
    
    session_folder = sys.argv[1]
    main(session_folder)
