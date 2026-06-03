#!/usr/bin/env python3
"""
Универсальный скрипт парсинга Excel файлов отчетов IFC (ifc_report.xlsx).
Извлекает ВСЕ материалы с их параметрами (объем, площадь, штуки) используя LLM gemma4:31b.

Особенности:
- Работает с любыми материалами (бетон, гидроизоляция, утеплитель, раствор, арматура и др.)
- Автоматически находит колонки с материалами по паттерну Qto_*:<номер>_<название материала>
- Определяет тип единицы измерения через LLM (объем м³, площадь м², штуки, литры)
- Группирует данные по полному наименованию материала с учетом подвидов
- Считает количество элементов для каждого материала
- Создает сводную таблицу с агрегированными данными
- Обрабатывает элементы без параметров объема/площади как поштучные
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
import requests


# URL Ollama API для работы с моделью gemma4:31b
OLLAMA_API_URL = "http://localhost:11434/api/generate"
LLM_MODEL = "gemma4:31b"


def call_llm(prompt: str, model: str = LLM_MODEL) -> Optional[str]:
    """
    Вызывает LLM-модель через Ollama API.
    
    Args:
        prompt: Текст запроса к модели
        model: Название модели
        
    Returns:
        Ответ модели или None при ошибке
    """
    try:
        payload = {
            "model": model,
            "prompt": prompt,
            "stream": False,
            "options": {
                "temperature": 0.1,
                "top_p": 0.9
            }
        }
        
        response = requests.post(OLLAMA_API_URL, json=payload, timeout=60)
        
        if response.status_code == 200:
            result = response.json()
            return result.get('response', '').strip()
        else:
            print(f"   ⚠️  LLM API error: {response.status_code}")
            return None
            
    except requests.exceptions.ConnectionError:
        print(f"   ⚠️  Не удалось подключиться к Ollama API ({OLLAMA_API_URL})")
        return None
    except Exception as e:
        print(f"   ⚠️  Ошибка при вызове LLM: {e}")
        return None


def analyze_material_with_llm(material_name: str) -> Dict[str, Any]:
    """
    Анализирует материал с помощью LLM gemma4:31b для определения:
    - Типа единицы измерения (volume/area/length/pieces/volume_liquid)
    - Группы материала (Бетон, Гидроизоляция, Утеплитель и т.д.)
    
    Args:
        material_name: Название материала
        
    Returns:
        Словарь с результатами анализа
    """
    prompt = f"""Проанализируй строительный материал и определи его характеристики.

Материал: "{material_name}"

Ответь ТОЛЬКО в формате JSON без дополнительных пояснений:
{{
    "unit_type": "volume" или "area" или "length" или "pieces" или "volume_liquid",
    "unit_label": "м³" или "м²" или "м" или "шт" или "л",
    "material_group": "Бетон" или "Гидроизоляция" или "Утеплитель" или "Раствор" или "Арматура" или "Праймер" или "Мастика" или "Изоляция" или "Другой"
}}

Правила определения типа:
- volume (м³): бетон, раствор, грунт, песок, щебень, керамзит, любые объемные материалы
- area (м²): гидроизоляция, пароизоляция, теплоизоляция, утеплитель, мембрана, пленка, покрытие, облицовка, штукатурка
- length (м): шнур, профиль, труба, кабель, провод, арматура (погонные метры)
- pieces (шт): анкер, дюбель, саморез, болт, гайка, шайба, закладная деталь, элемент, изделие, конструкция, сетка, каркас
- volume_liquid (л): праймер, мастика, клей, герметик, жидкость

Если материал содержит марку (например "Бетон B30 W6 F150"), определяй тип по основному слову ("Бетон" -> volume).
"""

    response = call_llm(prompt)
    
    # Если LLM вернул результат, парсим его
    if response:
        try:
            # Пытаемся найти JSON в ответе
            json_match = re.search(r'\{[^{}]*\}', response, re.DOTALL)
            if json_match:
                result = json.loads(json_match.group())
                return {
                    "unit_type": result.get("unit_type", "pieces"),
                    "unit_label": result.get("unit_label", "шт"),
                    "material_group": result.get("material_group", "Другой")
                }
        except (json.JSONDecodeError, AttributeError):
            pass
    
    # Fallback: используем паттерны если LLM недоступен или не вернул результат
    unit_patterns = {
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
    
    material_groups_fallback = {
        'Бетон': ['бетон', 'B', 'В'],
        'Гидроизоляция': ['гидроизоляц', 'мембран', 'техноэласт', 'плантер'],
        'Утеплитель': ['утеплител', 'полистирол', 'carbon', 'пеноплэкс', 'пенопласт'],
        'Раствор': ['раствор', 'стяжка', 'М100', 'М150', 'М200'],
        'Праймер': ['праймер', 'битумный'],
        'Мастика': ['мастик', 'приклеивающ'],
        'Арматура': ['арматур', 'сетк', 'каркас', 'reinforc'],
        'Изоляция': ['пароизоляц', 'теплоизоляц'],
    }
    
    name_lower = material_name.lower()
    
    # Определяем тип единицы измерения по паттернам
    unit_type = 'pieces'
    for utype, patterns in unit_patterns.items():
        for pattern in patterns:
            if re.search(pattern, name_lower):
                unit_type = utype
                break
        if unit_type != 'pieces':
            break
    
    # Определяем группу материала по паттернам
    material_group = "Другой"
    for group, keywords in material_groups_fallback.items():
        for keyword in keywords:
            if keyword.lower() in name_lower:
                material_group = group
                break
        if material_group != "Другой":
            break
    
    # Определяем label единицы измерения
    unit_labels = {
        'volume': 'м³',
        'area': 'м²',
        'length': 'м',
        'volume_liquid': 'л',
        'pieces': 'шт'
    }
    unit_label = unit_labels.get(unit_type, 'шт')
    
    return {
        "unit_type": unit_type,
        "unit_label": unit_label,
        "material_group": material_group
    }


# Mapping для основных типов элементов IFC к их Qto_ префиксам и колонкам объема/площади
IFC_ELEMENT_MAPPING = {
    'IfcWall': {'qto_prefix': 'Qto_WallBaseQuantities', 'volume_col': 'NetVolume', 'area_col': 'NetSideArea', 'length_col': 'Length'},
    'IfcColumn': {'qto_prefix': 'Qto_ColumnBaseQuantities', 'volume_col': 'NetVolume', 'area_col': 'OuterSurfaceArea', 'length_col': 'Length'},
    'IfcSlab': {'qto_prefix': 'Qto_SlabBaseQuantities', 'volume_col': 'NetVolume', 'area_col': 'NetArea', 'length_col': None},
    'IfcStair': {'qto_prefix': 'Qto_StairBaseQuantities', 'volume_col': None, 'area_col': None, 'length_col': None},
    'IfcRamp': {'qto_prefix': 'Qto_RampBaseQuantities', 'volume_col': None, 'area_col': None, 'length_col': None},
    'IfcBeam': {'qto_prefix': 'Qto_BeamBaseQuantities', 'volume_col': 'NetVolume', 'area_col': None, 'length_col': None},
    'IfcFooting': {'qto_prefix': 'Qto_FootingBaseQuantities', 'volume_col': 'NetVolume', 'area_col': None, 'length_col': None},
    'IfcPile': {'qto_prefix': 'Qto_PileBaseQuantities', 'volume_col': 'NetVolume', 'area_col': None, 'length_col': None},
    'IfcRoof': {'qto_prefix': 'Qto_RoofBaseQuantities', 'volume_col': 'NetVolume', 'area_col': 'NetArea', 'length_col': None},
    'IfcPlate': {'qto_prefix': 'Qto_PlateBaseQuantities', 'volume_col': None, 'area_col': 'NetArea', 'length_col': None},
    'IfcMember': {'qto_prefix': 'Qto_MemberBaseQuantities', 'volume_col': None, 'area_col': None, 'length_col': None},
    'IfcCurtainWall': {'qto_prefix': 'Qto_CurtainWallBaseQuantities', 'volume_col': None, 'area_col': 'NetSideArea', 'length_col': None},
    'IfcWindow': {'qto_prefix': 'Qto_WindowBaseQuantities', 'volume_col': None, 'area_col': None, 'length_col': None},
    'IfcDoor': {'qto_prefix': 'Qto_DoorBaseQuantities', 'volume_col': None, 'area_col': None, 'length_col': None},
    'IfcFurnishingElement': {'qto_prefix': 'Qto_FurnishingElementBaseQuantities', 'volume_col': None, 'area_col': None, 'length_col': None},
    'IfcBuildingElementProxy': {'qto_prefix': 'Qto_BuildingElementProxyBaseQuantities', 'volume_col': None, 'area_col': None, 'length_col': None},
    'IfcOpeningElement': {'qto_prefix': 'Qto_OpeningElementBaseQuantities', 'volume_col': None, 'area_col': 'Area', 'length_col': None},
    'IfcReinforcingMesh': {'qto_prefix': 'Qto_ReinforcingMeshBaseQuantities', 'volume_col': None, 'area_col': None, 'length_col': None},
    'IfcBuildingStorey': {'qto_prefix': 'Qto_BuildingStoreyBaseQuantities', 'volume_col': None, 'area_col': None, 'length_col': None},
    'IfcBuilding': {'qto_prefix': 'Qto_BuildingBaseQuantities', 'volume_col': None, 'area_col': None, 'length_col': None},
}


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
    Использует LLM gemma4:31b для анализа материалов.
    
    Особенности:
    - Находит колонки с материалами по паттерну Qto_*:<номер>_<название материала>
    - Параметры размерности могут быть в разных столбцах для разных категорий элементов
    - Обрабатывает материалы с подвидами (например "Бетон B30", "Бетон B35")
    """
    
    print(f"\n📊 Парсинг Excel отчета IFC: {os.path.basename(excel_path)}")
    print("   🤖 Используем модель gemma4:31b для анализа материалов...")
    
    try:
        # Читаем лист "Элементы"
        df = pd.read_excel(excel_path, sheet_name='Элементы', dtype=str)
        
        # Нормализуем имена колонок
        df.columns = [str(col).strip() for col in df.columns]
        
        # Ожидаемые имена колонок
        col_long_name = 'Element Specific:LongName'
        col_ifc_class = 'Ifc Class'
        
        # Проверяем наличие обязательных колонок
        if col_ifc_class not in df.columns:
            raise ValueError(f"Столбец '{col_ifc_class}' не найден в файле.")
        if col_long_name not in df.columns:
            col_long_name = 'Element Specific:Name'
            if col_long_name not in df.columns:
                raise ValueError(f"Столбец '{col_long_name}' не найден в файле")
        
        # Находим все колонки с материалами по паттерну Qto_*:<цифра>_<название>
        material_columns = []
        material_pattern = re.compile(r'Qto_\w+:(\d+)_([^_].*)$')
        
        for col in df.columns:
            match = material_pattern.search(col)
            if match:
                num = match.group(1)
                material_name = match.group(2).strip()
                material_columns.append({
                    'column': col,
                    'number': num,
                    'material_name': material_name
                })
        
        print(f"   📋 Найдено {len(material_columns)} колонок с материалами")
        
        items = []
        skipped_rows = 0
        llm_cache = {}  # Кэш для результатов LLM
        
        for idx, row in df.iterrows():
            ifc_class = row.get(col_ifc_class)
            if not ifc_class or pd.isna(ifc_class):
                skipped_rows += 1
                continue
            
            ifc_class = str(ifc_class).strip()
            
            long_name = row.get(col_long_name, '')
            if pd.isna(long_name):
                long_name = ''
            
            # Флаг: были ли найдены параметры для этой строки
            has_params = False
            
            # Проходим по всем колонкам с материалами
            for mat_col_info in material_columns:
                col_name = mat_col_info['column']
                material_name = mat_col_info['material_name']
                
                if col_name not in df.columns:
                    continue
                
                value = row.get(col_name)
                if value and str(value) != '0' and str(value) != 'nan' and str(value).strip():
                    try:
                        numeric_value = float(str(value).replace(',', '.'))
                        if numeric_value <= 0:
                            continue
                        
                        has_params = True
                        
                        # Используем LLM для анализа материала (с кэшированием)
                        if material_name not in llm_cache:
                            llm_result = analyze_material_with_llm(material_name)
                            llm_cache[material_name] = llm_result
                        else:
                            llm_result = llm_cache[material_name]
                        
                        material_group = llm_result['material_group']
                        unit_type = llm_result['unit_type']
                        unit_label = llm_result['unit_label']
                        
                        items.append({
                            'ifc_class': ifc_class,
                            'element_name': long_name,
                            'material_full': material_name,
                            'material_group': material_group,
                            'unit_type': unit_type,
                            'unit_label': unit_label,
                            'value': numeric_value,
                            'metric_type': unit_type,
                            'source_column': col_name
                        })
                        
                    except (ValueError, TypeError):
                        pass
            
            # Если нет ни одного параметра (объем/площадь/длина) - добавляем как штуку
            if not has_params:
                # Определяем базовый материал из MGE_Material если есть
                base_material_col = f'ExpCheck_{ifc_class[3:]}:MGE_Material'
                base_material = 'Неизвестный материал'
                
                if base_material_col in df.columns:
                    mat_val = row.get(base_material_col)
                    if mat_val and str(mat_val) != '0' and str(mat_val) != 'nan':
                        base_material = str(mat_val).strip()
                
                # Используем LLM для анализа материала
                if base_material not in llm_cache:
                    llm_result = analyze_material_with_llm(base_material)
                    llm_cache[base_material] = llm_result
                else:
                    llm_result = llm_cache[base_material]
                
                items.append({
                    'ifc_class': ifc_class,
                    'element_name': long_name,
                    'material_full': base_material,
                    'material_group': llm_result['material_group'],
                    'unit_type': 'pieces',
                    'unit_label': 'шт',
                    'value': 1.0,
                    'metric_type': 'pieces',
                    'source_column': None
                })
        
        print(f"   ✅ Извлечено {len(items)} записей о материалах (пропущено {skipped_rows} нерелевантных строк)")
        print(f"   🧠 Проанализировано {len(llm_cache)} уникальных материалов через LLM")
        
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
