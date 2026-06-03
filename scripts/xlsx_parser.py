#!/usr/bin/env python3
"""
Скрипт парсинга Excel файлов спецификаций.
Извлекает данные по элементам: марка материала, объем, площадь, количество.
Агрегирует данные: общий объем материала конкретной марки среди всех элементов.
Если объема нет - считается площадь. Если нет данных по размерам - считается количество.
Сохраняет результат в Excel файл в папке сессии.
"""

import os
import json
from pathlib import Path
from typing import Dict, List, Any, Optional
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


def find_excel_file_in_specification(session_folder: str) -> Optional[str]:
    """
    Находит Excel файл в папке specification.
    
    Args:
        session_folder: Путь к папке сессии
        
    Returns:
        Путь к Excel файлу или None
    """
    spec_folder = Path(session_folder) / 'specification'
    
    if not spec_folder.exists():
        print(f"⚠️  Папка specification не найдена: {spec_folder}")
        return None
    
    # Ищем файлы .xlsx и .xls
    excel_files = list(spec_folder.glob("*.xlsx")) + list(spec_folder.glob("*.xls"))
    
    if not excel_files:
        print(f"⚠️  Excel файлы не найдены в specification/")
        return None
    
    # Возвращаем первый найденный файл (можно расширить логику выбора)
    return str(excel_files[0])


def parse_excel_specification(excel_path: str) -> Dict[str, Any]:
    """
    Парсит Excel файл спецификации и извлекает данные по элементам.
    
    Args:
        excel_path: Путь к Excel файлу
        
    Returns:
        Словарь с данными: {марка_материала: {'volume': ..., 'area': ..., 'count': ...}}
    """
    
    print(f"\n📊 Парсинг Excel спецификации: {os.path.basename(excel_path)}")
    
    try:
        # Читаем все листы
        xls = pd.ExcelFile(excel_path)
        sheet_names = xls.sheet_names
        
        all_data = []
        
        for sheet_name in sheet_names:
            print(f"   📄 Чтение листа: {sheet_name}")
            
            # Читаем лист в DataFrame
            df = pd.read_excel(excel_path, sheet_name=sheet_name)
            
            # Нормализуем названия колонок (приводим к нижнему регистру, убираем пробелы)
            df.columns = [str(col).lower().strip() for col in df.columns]
            
            # Пытаемся найти колонки с нужными данными
            # Возможные названия колонок для марки материала
            material_cols = ['марка', 'material', 'марка материала', 'material grade', 'тип', 'type']
            material_col = None
            for col in material_cols:
                if col in df.columns:
                    material_col = col
                    break
            
            # Возможные названия колонок для объема
            volume_cols = ['объем', 'volume', 'v', 'объем (м3)', 'vol', 'куб.м', 'м3']
            volume_col = None
            for col in volume_cols:
                if col in df.columns:
                    volume_col = col
                    break
            
            # Возможные названия колонок для площади
            area_cols = ['площадь', 'area', 'a', 'площадь (м2)', 's', 'кв.м', 'м2']
            area_col = None
            for col in area_cols:
                if col in df.columns:
                    area_col = col
                    break
            
            # Возможные названия колонок для количества
            count_cols = ['количество', 'count', 'qty', 'шт', 'штук', 'число', 'n']
            count_col = None
            for col in count_cols:
                if col in df.columns:
                    count_col = col
                    break
            
            print(f"      Найдены колонки: марка={material_col}, объем={volume_col}, площадь={area_col}, количество={count_col}")
            
            # Извлекаем данные
            for idx, row in df.iterrows():
                # Пропускаем пустые строки
                if pd.isna(row).all():
                    continue
                
                item_data = {
                    'sheet': sheet_name,
                    'row': idx + 2,  # +2 т.к. нумерация с 1 и заголовок
                    'material': None,
                    'volume': None,
                    'area': None,
                    'count': 1  # По умолчанию считаем как 1 элемент
                }
                
                # Извлекаем марку материала
                if material_col and material_col in df.columns:
                    val = row.get(material_col)
                    if pd.notna(val):
                        item_data['material'] = str(val).strip()
                
                # Извлекаем объем
                if volume_col and volume_col in df.columns:
                    val = row.get(volume_col)
                    if pd.notna(val):
                        try:
                            item_data['volume'] = float(val)
                        except (ValueError, TypeError):
                            pass
                
                # Извлекаем площадь
                if area_col and area_col in df.columns:
                    val = row.get(area_col)
                    if pd.notna(val):
                        try:
                            item_data['area'] = float(val)
                        except (ValueError, TypeError):
                            pass
                
                # Извлекаем количество
                if count_col and count_col in df.columns:
                    val = row.get(count_col)
                    if pd.notna(val):
                        try:
                            item_data['count'] = int(float(val))
                        except (ValueError, TypeError):
                            pass
                
                # Добавляем только если есть хотя бы марка материала
                if item_data['material']:
                    all_data.append(item_data)
        
        print(f"   ✅ Извлечено {len(all_data)} элементов")
        
        return {
            'success': True,
            'items': all_data,
            'total_items': len(all_data),
            'sheets_processed': len(sheet_names)
        }
        
    except Exception as e:
        print(f"   ❌ Ошибка при парсинге Excel: {e}")
        import traceback
        traceback.print_exc()
        return {
            'success': False,
            'error': str(e),
            'items': [],
            'total_items': 0
        }


def aggregate_materials_data(items: List[Dict[str, Any]]) -> Dict[str, Dict[str, float]]:
    """
    Агрегирует данные по материалам: суммирует объем, площадь, количество для каждой марки.
    
    Логика:
    - Если у элемента есть объем - используем его
    - Если объема нет, но есть площадь - используем площадь
    - Если нет ни объема ни площади - считаем количество элементов
    
    Args:
        items: Список элементов с данными
        
    Returns:
        Словарь: {марка_материала: {'volume': ..., 'area': ..., 'count': ...}}
    """
    
    aggregated = {}
    
    for item in items:
        material = item.get('material')
        if not material:
            continue
        
        if material not in aggregated:
            aggregated[material] = {
                'volume': 0.0,
                'area': 0.0,
                'count': 0,
                'elements_with_volume': 0,
                'elements_with_area': 0
            }
        
        volume = item.get('volume')
        area = item.get('area')
        count = item.get('count', 1)
        
        if volume is not None and volume > 0:
            aggregated[material]['volume'] += volume
            aggregated[material]['elements_with_volume'] += 1
        elif area is not None and area > 0:
            aggregated[material]['area'] += area
            aggregated[material]['elements_with_area'] += 1
        else:
            # Нет данных по размерам - считаем количество
            aggregated[material]['count'] += count
    
    return aggregated


def create_aggregated_excel(aggregated_data: Dict[str, Dict[str, float]], output_path: str) -> str:
    """
    Создает Excel файл с агрегированными данными по материалам.
    
    Args:
        aggregated_data: Словарь с агрегированными данными
        output_path: Путь для сохранения файла
        
    Returns:
        Путь к сохраненному файлу
    """
    
    print(f"\n📈 Создание Excel отчета: {output_path}")
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Сводная по материалам"
    
    # Заголовки
    headers = [
        "№",
        "Марка материала",
        "Общий объем (м³)",
        "Элементов с объемом",
        "Общая площадь (м²)",
        "Элементов с площадью",
        "Количество (шт)",
        "Примечание"
    ]
    
    # Стили
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
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
        # Определяем примечание
        note = ""
        if data['elements_with_volume'] > 0:
            note = f"Расчет по объему ({data['elements_with_volume']} эл.)"
        elif data['elements_with_area'] > 0:
            note = f"Расчет по площади ({data['elements_with_area']} эл.)"
        else:
            note = f"Расчет по количеству ({data['count']} шт.)"
        
        ws.cell(row=row_num, column=1, value=idx).border = thin_border
        ws.cell(row=row_num, column=2, value=material).border = thin_border
        ws.cell(row=row_num, column=3, value=round(data['volume'], 3) if data['volume'] > 0 else None).border = thin_border
        ws.cell(row=row_num, column=4, value=data['elements_with_volume']).border = thin_border
        ws.cell(row=row_num, column=5, value=round(data['area'], 3) if data['area'] > 0 else None).border = thin_border
        ws.cell(row=row_num, column=6, value=data['elements_with_area']).border = thin_border
        ws.cell(row=row_num, column=7, value=data['count'] if data['volume'] <= 0 and data['area'] <= 0 else None).border = thin_border
        ws.cell(row=row_num, column=8, value=note).border = thin_border
        
        row_num += 1
    
    # Автоширина колонок
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column].width = adjusted_width
    
    # Сохраняем
    wb.save(output_path)
    print(f"   ✅ Excel сохранен: {output_path}")
    
    return output_path


def parse_and_aggregate_specification(session_folder: str) -> Dict[str, Any]:
    """
    Основная функция: находит Excel файл, парсит и агрегирует данные.
    
    Args:
        session_folder: Путь к папке сессии
        
    Returns:
        Словарь с результатами
    """
    
    print("\n" + "="*60)
    print("📊 ШАГ 6: ПАРСИНГ EXCEL СПЕЦИФИКАЦИИ И АГРЕГАЦИЯ ДАННЫХ")
    print("="*60)
    
    results = {
        'success': False,
        'excel_file_found': None,
        'parsed_items': 0,
        'aggregated_materials': 0,
        'output_file': None
    }
    
    # Шаг 1: Найти Excel файл
    excel_path = find_excel_file_in_specification(session_folder)
    
    if not excel_path:
        print("❌ Excel файл спецификации не найден")
        return results
    
    results['excel_file_found'] = excel_path
    
    # Шаг 2: Распарсить Excel
    parse_result = parse_excel_specification(excel_path)
    
    if not parse_result['success']:
        print("❌ Ошибка парсинга Excel файла")
        results['error'] = parse_result.get('error')
        return results
    
    results['parsed_items'] = parse_result['total_items']
    
    # Шаг 3: Агрегировать данные по материалам
    aggregated_data = aggregate_materials_data(parse_result['items'])
    results['aggregated_materials'] = len(aggregated_data)
    
    # Шаг 4: Создать выходной Excel файл
    output_filename = "materials_summary.xlsx"
    output_path = Path(session_folder) / output_filename
    
    create_aggregated_excel(aggregated_data, str(output_path))
    results['output_file'] = str(output_path)
    results['success'] = True
    
    # Шаг 5: Сохранить JSON с данными
    json_output_path = Path(session_folder) / "materials_summary.json"
    with open(json_output_path, 'w', encoding='utf-8') as f:
        json.dump({
            'success': True,
            'source_file': os.path.basename(excel_path),
            'total_items_parsed': parse_result['total_items'],
            'total_materials': len(aggregated_data),
            'output_excel': output_filename,
            'materials': aggregated_data
        }, f, ensure_ascii=False, indent=2)
    
    print(f"\n💾 Данные сохранены в: {json_output_path}")
    
    # Выводим сводку
    print("\n" + "="*60)
    print("📊 СВОДКА ПО МАТЕРИАЛАМ:")
    print("="*60)
    for material, data in sorted(aggregated_data.items()):
        if data['volume'] > 0:
            print(f"   • {material}: {data['volume']:.3f} м³ ({data['elements_with_volume']} эл.)")
        elif data['area'] > 0:
            print(f"   • {material}: {data['area']:.3f} м² ({data['elements_with_area']} эл.)")
        else:
            print(f"   • {material}: {data['count']} шт.")
    
    return results


def main(session_folder: str):
    """
    Основная функция парсинга спецификации.
    
    Args:
        session_folder: Путь к папке сессии
    """
    print("="*60)
    print("📊 ПАРСИНГ EXCEL СПЕЦИФИКАЦИИ")
    print("="*60)
    print(f"Папка сессии: {session_folder}")
    print("="*60)
    
    results = parse_and_aggregate_specification(session_folder)
    
    return results


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Использование: python xlsx_parser.py <session_folder>")
        print("Пример: python xlsx_parser.py /path/to/uploads/session_id")
        sys.exit(1)
    
    session_folder = sys.argv[1]
    main(session_folder)
