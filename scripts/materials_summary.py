#!/usr/bin/env python3
"""
Скрипт для создания сводной таблицы по материалам из IFC отчета.
Для каждого листа (кроме "Сводка") создает таблицу с агрегацией по материалам:
- Общий объем по материалу
- Объем по типу материала
- Количество элементов из этого материала

Для элементов без объема указывает количество и площадь.
Также создает лист "Сводка" с общим объемом по каждому материалу.
"""

import pandas as pd
import sys
from pathlib import Path


def process_sheet(df, sheet_name):
    """
    Обрабатывает один лист и возвращает агрегированные данные по материалам.
    
    Для элементов с объемом:
    - Суммарный объем по материалу
    - Объем по типу материала  
    - Количество элементов
    
    Для элементов без объема:
    - Количество элементов
    - Площадь
    """
    # Копируем данные для работы
    data = df.copy()
    
    # Заполняем NaN в колонке Материал для группировки
    material_col = 'Материал'
    volume_col = 'Объем (м³)'
    area_col = 'Площадь (м²)'
    
    # Заменяем NaN в названиях материалов на строку "Без материала"
    data[material_col] = data[material_col].fillna('Без материала')
    
    # Разделяем элементы с объемом и без объема
    has_volume = data[volume_col].notna() & (data[volume_col] != 0)
    no_volume = ~has_volume
    
    result_data = []
    
    # Обработка элементов с объемом
    if has_volume.any():
        vol_data = data[has_volume].copy()
        
        # Группировка по материалу
        grouped = vol_data.groupby(material_col).agg({
            volume_col: 'sum',
            'ID': 'count'
        }).reset_index()
        
        grouped.columns = ['Материал', 'Общий объем (м³)', 'Количество элементов']
        
        # Добавляем детали по объемам для каждого материала
        for material in grouped['Материал'].unique():
            mat_data = vol_data[vol_data[material_col] == material]
            
            # Получаем уникальный тип материала (первое значение)
            material_type = mat_data[material_col].iloc[0] if len(mat_data) > 0 else material
            
            total_vol = mat_data[volume_col].sum()
            count = len(mat_data)
            
            result_data.append({
                'Материал': material,
                'Тип материала': material_type,
                'Общий объем (м³)': round(total_vol, 3),
                'Объем по типу (м³)': round(total_vol, 3),
                'Количество элементов': count,
                'Площадь (м²)': None,
                'Примечание': 'Есть объем'
            })
    
    # Обработка элементов без объема
    if no_volume.any():
        no_vol_data = data[no_volume].copy()
        
        # Проверяем есть ли площадь
        has_area = no_vol_data[area_col].notna()
        
        if has_area.any():
            # Группировка по материалу для элементов с площадью
            area_grouped = no_vol_data[has_area].groupby(material_col).agg({
                area_col: 'sum',
                'ID': 'count'
            }).reset_index()
            
            for _, row in area_grouped.iterrows():
                material = row[material_col]
                total_area = row[area_col]
                count = row['ID']
                
                result_data.append({
                    'Материал': material,
                    'Тип материала': material,
                    'Общий объем (м³)': None,
                    'Объем по типу (м³)': None,
                    'Количество элементов': count,
                    'Площадь (м²)': round(total_area, 3),
                    'Примечание': 'Нет объема, есть площадь'
                })
        
        # Элементы без объема и без площади
        no_area = no_vol_data[~has_area]
        if len(no_area) > 0:
            count_only = no_area.groupby(material_col).agg({
                'ID': 'count'
            }).reset_index()
            
            for _, row in count_only.iterrows():
                material = row[material_col]
                count = row['ID']
                
                result_data.append({
                    'Материал': material,
                    'Тип материала': material,
                    'Общий объем (м³)': None,
                    'Объем по типу (м³)': None,
                    'Количество элементов': count,
                    'Площадь (м²)': None,
                    'Примечание': 'Только количество'
                })
    
    return pd.DataFrame(result_data)


def create_summary(all_data):
    """
    Создает сводную таблицу по всем материалам с общим объемом.
    """
    summary_data = []
    
    # Собираем все материалы и их объемы
    material_volumes = {}
    
    for sheet_name, df in all_data.items():
        if sheet_name == 'Сводка':
            continue
            
        # Обрабатываем каждый лист
        processed = process_sheet(df, sheet_name)
        
        if len(processed) > 0:
            for _, row in processed.iterrows():
                material = row['Материал']
                volume = row['Общий объем (м³)']
                
                if volume is not None:
                    if material not in material_volumes:
                        material_volumes[material] = 0
                    material_volumes[material] += volume
    
    # Создаем итоговую таблицу
    for material, total_volume in sorted(material_volumes.items()):
        summary_data.append({
            'Материал': material,
            'Общий объем (м³)': round(total_volume, 3)
        })
    
    return pd.DataFrame(summary_data)


def main(input_file, output_file=None):
    """
    Основная функция обработки IFC отчета.
    
    Args:
        input_file: Путь к входному файлу ifc_report.xlsx
        output_file: Путь к выходному файлу (по умолчанию добавляется _materials перед расширением)
    """
    input_path = Path(input_file)
    
    if not input_path.exists():
        print(f"Ошибка: Файл {input_file} не найден")
        sys.exit(1)
    
    # Читаем все листы из Excel файла
    xl_file = pd.ExcelFile(input_path)
    sheet_names = xl_file.sheet_names
    
    print(f"Найдено листов: {len(sheet_names)}")
    print(f"Листы: {', '.join(sheet_names)}")
    
    # Словарь для хранения данных по листам
    all_data = {}
    
    # Читаем каждый лист
    for sheet_name in sheet_names:
        df = pd.read_excel(input_path, sheet_name=sheet_name)
        all_data[sheet_name] = df
        print(f"\nЛист '{sheet_name}': {len(df)} строк")
    
    # Создаем новый Excel файл с результатами
    if output_file is None:
        output_path = input_path.parent / f"{input_path.stem}_materials{input_path.suffix}"
    else:
        output_path = Path(output_file)
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Обрабатываем каждый лист (кроме существующей Сводки)
        for sheet_name in sheet_names:
            if sheet_name == 'Сводка':
                continue
                
            df = all_data[sheet_name]
            processed_df = process_sheet(df, sheet_name)
            
            if len(processed_df) > 0:
                processed_df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Создан лист '{sheet_name}' с {len(processed_df)} материалами")
            else:
                # Если нет данных, создаем пустой лист с заголовками
                empty_df = pd.DataFrame(columns=[
                    'Материал', 'Тип материала', 'Общий объем (м³)', 
                    'Объем по типу (м³)', 'Количество элементов', 
                    'Площадь (м²)', 'Примечание'
                ])
                empty_df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Создан пустой лист '{sheet_name}'")
        
        # Создаем лист "Сводка по материалам"
        summary_df = create_summary(all_data)
        if len(summary_df) > 0:
            summary_df.to_excel(writer, sheet_name='Сводка по материалам', index=False)
            print(f"\nСоздан лист 'Сводка по материалам' с {len(summary_df)} материалами")
        else:
            print("\nНет данных для сводки по материалам")
    
    print(f"\nРезультат сохранен в: {output_path}")
    return output_path


if __name__ == '__main__':
    import argparse
    
    parser = argparse.ArgumentParser(
        description='Создание сводной таблицы по материалам из IFC отчета'
    )
    parser.add_argument(
        'input_file',
        help='Путь к файлу ifc_report.xlsx'
    )
    parser.add_argument(
        '-o', '--output',
        help='Путь к выходному файлу (по умолчанию создается рядом с входным)'
    )
    
    args = parser.parse_args()
    
    main(args.input_file, args.output)
