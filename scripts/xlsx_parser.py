#!/usr/bin/env python3
"""
Скрипт парсинга Excel файлов отчетов IFC (ifc_report.xlsx).
Извлекает марку материала из отдельной колонки ExpCheck_MaterialConcrete:MGE_ConcreteGrade,
объём (в зависимости от типа элемента) или количество.
Агрегирует данные по маркам материалов (например, "Бетон В30", "Бетон В35").
Сохраняет результат в Excel файл в папке сессии.

Особенности:
- Работает с файлами ifc_report.xlsx (лист "Элементы")
- Использует колонку ExpCheck_MaterialConcrete:MGE_ConcreteGrade для марки бетона (В30, В35, В10, В25)
- Определяет тип элемента по столбцу Ifc Class
- Для каждого типа элемента использует соответствующую колонку объема:
  * IfcWall → Qto_WallBaseQuantities:NetVolume
  * IfcColumn → Qto_ColumnBaseQuantities:NetVolume
  * IfcSlab → Qto_SlabBaseQuantities:NetVolume
  * IfcStair, IfcRamp → только количество (нет объема)
- Агрегация по полной марке (например, "Бетон В30")
"""

import os
import re
import json
from pathlib import Path
from typing import Dict, List, Any, Optional
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


# Словарь соответствия: Ifc класс -> (столбец_материала, столбец_объёма)
# Используем точные имена колонок из ifc_report.xlsx
IFC_MAPPING = {
    'IfcWall': ('ExpCheck_Wall:MGE_Material', 'Qto_WallBaseQuantities:NetVolume'),
    'IfcColumn': ('ExpCheck_Column:MGE_Material', 'Qto_ColumnBaseQuantities:NetVolume'),
    'IfcSlab': ('ExpCheck_Slab:MGE_Material', 'Qto_SlabBaseQuantities:NetVolume'),
    'IfcStair': ('ExpCheck_Stair:MGE_Material', None),          # нет объёма, только количество
    'IfcRamp': ('ExpCheck_Ramp:MGE_Material', None),            # нет объёма, только количество
}

# Колонка с маркой бетона (общая для всех типов элементов)
CONCRETE_GRADE_COLUMN = 'ExpCheck_MaterialConcrete:MGE_ConcreteGrade'

# Регулярное выражение для поиска марки бетона (на случай если нужно будет извлекать из названия)
# Ищет В30, В35, B30, B35, В40, B40 и т.д. (цифры от 10 до 100)
GRADE_PATTERN = re.compile(r'[ВB]([1-9]\d?|100)')


def extract_material_grade(name: str, concrete_grade_value: Any) -> str:
    """
    Извлекает марку бетона из колонки ExpCheck_MaterialConcrete:MGE_ConcreteGrade.
    Если значение не указано или равно '0', пытается извлечь из наименования через regex.
    Примеры: "В30", "В35", "В10", "В25", "B30", "B35"
    Если марка не найдена, возвращает "марка не указана".
    
    Args:
        name: Наименование элемента (для резервного извлечения)
        concrete_grade_value: Значение из колонки марки бетона
    
    Returns:
        Марка бетона (например, "В30") или "марка не указана"
    """
    # Сначала пробуем взять из явной колонки
    if concrete_grade_value and not pd.isna(concrete_grade_value):
        grade_str = str(concrete_grade_value).strip()
        if grade_str and grade_str != '0':
            return grade_str
    
    # Если не удалось, пробуем извлечь из названия
    if isinstance(name, str):
        match = GRADE_PATTERN.search(name)
        if match:
            return match.group(0)
    
    return "марка не указана"


def find_ifc_report_file(session_folder: str) -> Optional[str]:
    """
    Находит файл ifc_report.xlsx в папке сессии.
    
    Args:
        session_folder: Путь к папке сессии
        
    Returns:
        Путь к Excel файлу или None
    """
    report_path = Path(session_folder) / 'ifc_report.xlsx'
    
    if not report_path.exists():
        print(f"⚠️  Файл ifc_report.xlsx не найден: {report_path}")
        return None
    
    return str(report_path)


def parse_excel_specification(excel_path: str) -> Dict[str, Any]:
    """
    Парсит Excel файл отчета IFC (лист "Элементы"), извлекая данные по элементам.
    Использует точные имена колонок из ifc_report.xlsx.
    
    Args:
        excel_path: Путь к Excel файлу
        
    Returns:
        Словарь с данными: {success, items, total_items}
    """
    
    print(f"\n📊 Парсинг Excel отчета IFC: {os.path.basename(excel_path)}")
    
    try:
        # Читаем лист "Элементы"
        df = pd.read_excel(excel_path, sheet_name='Элементы', dtype=str)
        
        # Нормализуем имена колонок: убираем лишние пробелы
        df.columns = [str(col).strip() for col in df.columns]
        
        # Ожидаемые имена колонок
        col_long_name = 'Element Specific:LongName'
        col_ifc_class = 'Ifc Class'
        
        # Проверяем наличие обязательных колонок
        if col_ifc_class not in df.columns:
            raise ValueError(f"Столбец '{col_ifc_class}' не найден в файле. Доступные: {list(df.columns)[:10]}...")
        if col_long_name not in df.columns:
            # Пробуем альтернативное имя
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
            if ifc_class not in IFC_MAPPING:
                # Неинтересный тип элемента (IfcBuildingStorey, IfcGroup и т.д.)
                skipped_rows += 1
                continue
            
            material_col, volume_col = IFC_MAPPING[ifc_class]
            
            # Получаем марку бетона из общей колонки
            concrete_grade = row.get(CONCRETE_GRADE_COLUMN)
            
            # Получаем наименование элемента
            long_name = row.get(col_long_name)
            
            # Извлекаем марку бетона
            grade = extract_material_grade(long_name if pd.notna(long_name) else '', concrete_grade)
            
            # Получаем базовый материал
            material = row.get(material_col) if material_col in df.columns else None
            if pd.isna(material) or not material:
                material = "Бетон"  # значение по умолчанию
            material = str(material).strip()
            
            # Формируем полное название материала с маркой
            if grade != "марка не указана":
                full_material = f"{material} {grade}"
            else:
                full_material = material
            
            # Объём или количество
            volume = None
            count = 1   # по умолчанию 1 элемент
            
            if volume_col and volume_col in df.columns:
                vol_val = row.get(volume_col)
                if pd.notna(vol_val) and vol_val:
                    try:
                        # Заменяем запятую на точку для корректного преобразования
                        vol_str = str(vol_val).replace(',', '.')
                        volume = float(vol_str)
                    except (ValueError, TypeError):
                        pass
            
            # Если объёма нет, то учитываем как количество
            if volume is None or volume == 0:
                count = 1
                volume = None
            
            items.append({
                'ifc_class': ifc_class,
                'material_full': full_material,
                'material_base': material,
                'grade': grade,
                'volume': volume,
                'count': count,
                'source_name': long_name if pd.notna(long_name) else ''
            })
        
        print(f"   ✅ Извлечено {len(items)} элементов (пропущено {skipped_rows} нерелевантных строк)")
        
        return {
            'success': True,
            'items': items,
            'total_items': len(items),
        }
        
    except Exception as e:
        print(f"   ❌ Ошибка при парсинге Excel: {e}")
        import traceback
        traceback.print_exc()
        return {'success': False, 'error': str(e), 'items': []}


def aggregate_materials_data(items: List[Dict[str, Any]]) -> Dict[str, Dict[str, float]]:
    """
    Агрегирует данные по полному наименованию материала (с маркой).
    Если есть объём – суммируем, иначе – суммируем количество.
    
    Args:
        items: Список элементов с данными
        
    Returns:
        Словарь: {материал_с_маркой: {'volume': ..., 'count': ..., 'elements_with_volume': ..., 'elements_without_volume': ...}}
    """
    
    aggregated = {}
    
    for item in items:
        material = item['material_full']
        volume = item.get('volume')
        count = item.get('count', 1)
        
        if material not in aggregated:
            aggregated[material] = {
                'volume': 0.0,
                'count': 0,
                'elements_with_volume': 0,
                'elements_without_volume': 0,
                'base_material': item['material_base'],
                'grade': item['grade']
            }
        
        if volume is not None and volume > 0:
            aggregated[material]['volume'] += volume
            aggregated[material]['elements_with_volume'] += 1
        else:
            aggregated[material]['count'] += count
            aggregated[material]['elements_without_volume'] += 1
    
    return aggregated


def create_aggregated_excel(aggregated_data: Dict[str, Dict[str, float]], output_path: str) -> str:
    """
    Создаёт Excel отчёт с агрегированными данными (улучшенное форматирование).
    
    Args:
        aggregated_data: Словарь с агрегированными данными
        output_path: Путь для сохранения файла
        
    Returns:
        Путь к сохраненному файлу
    """
    
    print(f"\n📈 Создание Excel отчета: {output_path}")
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Сводка по материалам"
    
    headers = [
        "№", "Материал (с маркой)", "Базовая марка", "Марка бетона",
        "Общий объём (м³)", "Кол-во элементов с объёмом",
        "Количество (шт)", "Кол-во элементов без объёма", "Примечание"
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
    
    # Пишем данные
    row_num = 2
    for idx, (material, data) in enumerate(sorted(aggregated_data.items()), 1):
        note = ""
        if data['elements_with_volume'] > 0:
            note = f"расчёт по объёму ({data['elements_with_volume']} эл.)"
            if data['elements_without_volume'] > 0:
                note += f" + {data['elements_without_volume']} эл. по количеству"
        else:
            note = f"расчёт по количеству ({data['count']} шт.)"
        
        ws.cell(row=row_num, column=1, value=idx).border = thin_border
        ws.cell(row=row_num, column=2, value=material).border = thin_border
        ws.cell(row=row_num, column=3, value=data['base_material']).border = thin_border
        ws.cell(row=row_num, column=4, value=data['grade']).border = thin_border
        ws.cell(row=row_num, column=5, value=round(data['volume'], 3) if data['volume'] > 0 else None).border = thin_border
        ws.cell(row=row_num, column=6, value=data['elements_with_volume'] if data['volume'] > 0 else None).border = thin_border
        ws.cell(row=row_num, column=7, value=data['count'] if data['count'] > 0 else None).border = thin_border
        ws.cell(row=row_num, column=8, value=data['elements_without_volume'] if data['elements_without_volume'] > 0 else None).border = thin_border
        ws.cell(row=row_num, column=9, value=note).border = thin_border
        row_num += 1
    
    # Автоширина колонок
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_len + 2, 50)
    
    wb.save(output_path)
    print(f"   ✅ Excel сохранён: {output_path}")
    return output_path


def parse_and_aggregate_specification(session_folder: str) -> Dict[str, Any]:
    """
    Основная функция: поиск, парсинг, агрегация и сохранение.
    Работает с файлом ifc_report.xlsx в папке сессии.
    
    Args:
        session_folder: Путь к папке сессии
        
    Returns:
        Словарь с результатами
    """
    
    print("\n" + "="*60)
    print("📊 ПАРСИНГ ОТЧЕТА IFC (ifc_report.xlsx)")
    print("="*60)
    
    results = {'success': False, 'excel_file_found': None, 'parsed_items': 0,
               'aggregated_materials': 0, 'output_file': None}
    
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
            'materials': aggregated_data
        }, f, ensure_ascii=False, indent=2)
    
    print(f"\n💾 Данные сохранены в: {json_path}")
    
    # Сводка в консоль
    print("\n" + "="*60)
    print("📊 СВОДКА ПО МАТЕРИАЛАМ (с марками бетона):")
    print("="*60)
    for material, data in sorted(aggregated_data.items()):
        if data['volume'] > 0:
            print(f"   • {material}: {data['volume']:.3f} м³ ({data['elements_with_volume']} эл.)")
        else:
            print(f"   • {material}: {data['count']} шт. ({data['elements_without_volume']} эл.)")
    return results


def main(session_folder: str):
    """
    Основная функция парсинга отчета IFC.
    
    Args:
        session_folder: Путь к папке сессии
    """
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
