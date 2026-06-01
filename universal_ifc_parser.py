import ifcopenshell
import ifcopenshell.util.element
import ifcopenshell.util.unit
import pandas as pd
import os
from collections import defaultdict
from datetime import datetime

def get_unit_scale(ifc_file):
    """Получить масштаб единиц для корректного перевода размеров"""
    try:
        # Попытка использовать новый API
        units = ifcopenshell.util.unit.get_project_unit_scales(ifc_file)
        return units.get('LENGTHUNIT', 1.0)
    except AttributeError:
        # Старый API или альтернативный метод
        try:
            unit = ifcopenshell.util.unit.get_project_unit(ifc_file, 'LENGTHUNIT')
            if unit == 'METRE':
                return 1.0
            elif unit == 'MILLIMETRE':
                return 0.001
            elif unit == 'CENTIMETRE':
                return 0.01
            else:
                return 1.0
        except:
            return 1.0

def get_geometry_info(element, unit_scale):
    """Извлечь геометрию элемента (Bounding Box и основные размеры)"""
    info = {
        'Length_m': None,
        'Width_m': None,
        'Height_m': None,
        'Volume_m3': None,
        'Area_m2': None,
        'X_min': None, 'X_max': None,
        'Y_min': 'None', 'Y_max': None,
        'Z_min': None, 'Z_max': None
    }
    
    try:
        shape = element.Representation
        if shape:
            # Попытка получить Bounding Box через ifcopenshell
            # Примечание: это упрощенный метод, для сложной геометрии нужен более глубокий парсинг
            bb = ifcopenshell.util.element.get_bounding_box(element)
            if bb:
                info['X_min'] = round(bb[0][0] * unit_scale, 3)
                info['X_max'] = round(bb[1][0] * unit_scale, 3)
                info['Y_min'] = round(bb[0][1] * unit_scale, 3)
                info['Y_max'] = round(bb[1][1] * unit_scale, 3)
                info['Z_min'] = round(bb[0][2] * unit_scale, 3)
                info['Z_max'] = round(bb[1][2] * unit_scale, 3)
                
                dx = info['X_max'] - info['X_min']
                dy = info['Y_max'] - info['Y_min']
                dz = info['Z_max'] - info['Z_min']
                
                info['Length_m'] = round(dx, 3)
                info['Width_m'] = round(dy, 3)
                info['Height_m'] = round(dz, 3)
                
                # Эвристика для объема и площади (приближенно по боксу)
                if dx > 0 and dy > 0 and dz > 0:
                    info['Volume_m3'] = round(dx * dy * dz, 3)
                    # Площадь поверхности бокса (грубо)
                    info['Area_m2'] = round(2 * (dx*dy + dx*dz + dy*dz), 3)
    except Exception as e:
        pass
        
    return info

def get_all_properties(element):
    """Извлечь ВСЕ свойства элемента, включая Pset и пользовательские (MGE_)"""
    props = {}
    
    # 1. Стандартные атрибуты IFC (с проверкой существования атрибута)
    props['GlobalId'] = element.GlobalId if hasattr(element, 'GlobalId') else ""
    props['Name'] = element.Name if hasattr(element, 'Name') and element.Name else ""
    props['ObjectType'] = element.ObjectType if hasattr(element, 'ObjectType') and element.ObjectType else ""
    props['Tag'] = element.Tag if hasattr(element, 'Tag') and element.Tag else ""
    props['Description'] = element.Description if hasattr(element, 'Description') and element.Description else ""
    
    # 2. Материалы
    try:
        materials = ifcopenshell.util.element.get_materials(element, should_skip_usage=False)
    except TypeError:
        # Старая версия API
        materials = ifcopenshell.util.element.get_materials(element)
    
    mat_names = []
    for mat in materials:
        if hasattr(mat, 'Name') and mat.Name:
            mat_names.append(mat.Name)
        # Проверка на составные материалы (IfcMaterialLayerSet)
        if hasattr(mat, 'HasProperties'):
             for p_set in mat.HasProperties:
                 if hasattr(p_set, 'Properties'):
                     for prop in p_set.Properties:
                         if hasattr(prop, 'Name') and hasattr(prop, 'NominalValue'):
                             if prop.NominalValue:
                                 props[f'Mat_{prop.Name}'] = prop.NominalValue
    
    props['Materials'] = " | ".join(mat_names) if mat_names else ""

    # 3. Наборы свойств (Property Sets) - САМОЕ ВАЖНОЕ
    # Используем универсальный метод ifcopenshell для получения всех свойств
    all_property_sets = []
    
    # Метод 1: Прямые определения через IsDefinedBy
    if hasattr(element, 'IsDefinedBy'):
        for definition in element.IsDefinedBy:
            if definition.is_a('IfcRelDefinesByProperties'):
                # В IFC4 RelatingPropertyDefinition - это одиночный IfcPropertySet объект
                pset = definition.RelatingPropertyDefinition
                all_property_sets.append(pset)
            elif definition.is_a('IfcRelDefinesByType'):
                # Если определено через тип, берем свойства типа
                obj_type = definition.RelatingType
                if hasattr(obj_type, 'HasPropertySets'):
                    for p_set in obj_type.HasPropertySets:
                        all_property_sets.append(p_set)
    
    # Метод 2: Через IsTypedBy (для IFC4)
    if hasattr(element, 'IsTypedBy'):
        for type_def in element.IsTypedBy:
            if type_def.is_a('IfcRelDefinesByType'):
                obj_type = type_def.RelatingType
                if hasattr(obj_type, 'HasPropertySets'):
                    for p_set in obj_type.HasPropertySets:
                        all_property_sets.append(p_set)
    
    # Метод 3: Прямой поиск HasPropertySets у элемента (редко, но бывает)
    if hasattr(element, 'HasPropertySets'):
        for p_set in element.HasPropertySets:
            all_property_sets.append(p_set)

    # Парсинг свойств из наборов
    for p_set in all_property_sets:
        # В IFC4 используется HasProperties, в старых версиях Properties
        prop_list = None
        if hasattr(p_set, 'HasProperties'):
            prop_list = p_set.HasProperties
        elif hasattr(p_set, 'Properties'):
            prop_list = p_set.Properties
            
        if not prop_list:
            continue
            
        pset_name = p_set.Name if hasattr(p_set, 'Name') else "UnknownPset"
        
        for prop in prop_list:
            prop_name = prop.Name if hasattr(prop, 'Name') else "UnknownProp"
            value = None
            
            # Извлечение значения разными способами
            if hasattr(prop, 'NominalValue'):
                val_obj = prop.NominalValue
                if val_obj:
                    if hasattr(val_obj, 'wrappedValue'):
                        value = val_obj.wrappedValue
                    else:
                        value = str(val_obj)
            elif hasattr(prop, 'Values') and prop.Values:
                # Для списков значений
                value = ", ".join([str(v.wrappedValue) if hasattr(v, 'wrappedValue') else str(v) for v in prop.Values])
            
            # Формируем уникальное имя ключа
            final_key = prop_name
            if not prop_name.startswith("MGE_") and not prop_name.startswith("Reference"):
                 short_pset = pset_name.replace("Pset_", "").replace(" ", "_")
                 final_key = f"{short_pset}_{prop_name}"
            
            # Очистка ключа от недопустимых символов
            final_key = "".join([c if c.isalnum() or c in "_-" else "_" for c in final_key])
            
            if value is not None:
                props[final_key] = value

    return props

def parse_ifc_universal(file_path, output_excel):
    print(f"🚀 Запуск универсального парсера для: {file_path}")
    
    if not os.path.exists(file_path):
        print(f"❌ Файл не найден: {file_path}")
        return

    try:
        f = ifcopenshell.open(file_path)
    except Exception as e:
        print(f"❌ Ошибка открытия файла: {e}")
        return

    products = f.by_type('IfcProduct')
    print(f"🔍 Найдено элементов: {len(products)}")
    
    unit_scale = get_unit_scale(f)
    
    all_data = []
    all_keys = set()
    
    # Сбор всех возможных ключей заранее (чтобы столбцы были одинаковыми)
    # Это ресурсоемко, но гарантирует полноту таблицы
    print("📋 Сканирование структуры данных...")
    temp_props = []
    for i, element in enumerate(products):
        if i % 500 == 0:
            print(f"   Обработано {i}/{len(products)} элементов (сканирование)...")
        
        props = get_all_properties(element)
        geom = get_geometry_info(element, unit_scale)
        combined = {**props, **geom}
        combined['Class'] = element.is_a()
        temp_props.append(combined)
        all_keys.update(combined.keys())
    
    print(f"✅ Найдено уникальных параметров: {len(all_keys)}")
    
    # Формирование финальной таблицы с упорядочиванием колонок
    # Сначала важные, потом остальные по алфавиту
    priority_cols = ['Class', 'Name', 'Tag', 'ObjectType', 'Materials', 
                     'Length_m', 'Width_m', 'Height_m', 'Volume_m3', 'Area_m2',
                     'Z_min', 'Z_max', 'GlobalId']
    
    # Отделяем приоритетные от остальных
    other_cols = sorted(list(all_keys - set(priority_cols)))
    
    # Особая сортировка для MGE и Reference свойств, чтобы они были рядом
    mge_cols = [c for c in other_cols if 'MGE' in c or 'Reference' in c or 'Concrete' in c or 'Reinforcement' in c]
    other_clean = [c for c in other_cols if c not in mge_cols]
    
    final_columns = priority_cols + sorted(mge_cols) + other_clean
    
    print("📝 Формирование DataFrame...")
    df = pd.DataFrame(temp_props, columns=final_columns)
    
    # Заполнение пустот
    df = df.fillna('')
    
    # Сохранение в Excel
    print(f"💾 Сохранение в {output_excel}...")
    try:
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='All_Elements_Full', index=False)
            
            # Добавляем сводку на первый лист
            summary_data = {
                'Metric': ['Total Elements', 'Unique Classes', 'File', 'Date Processed'],
                'Value': [len(df), df['Class'].nunique(), os.path.basename(file_path), datetime.now().strftime("%Y-%m-%d %H:%M")]
            }
            df_summary = pd.DataFrame(summary_data)
            
            # Вставляем сводку перед основной таблицей в том же файле? 
            # Лучше отдельным листом для чистоты
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            
            # Группировка по классам для удобства
            class_counts = df['Class'].value_counts().reset_index()
            class_counts.columns = ['Class', 'Count']
            class_counts.to_excel(writer, sheet_name='Counts_by_Class', index=False)
            
        print(f"✅ Успешно сохранено в {output_excel}")
        print(f"📊 Строк данных: {len(df)}")
        print(f"📏 Столбцов (параметров): {len(df.columns)}")
        
        # Вывод примера найденных MGE свойств
        mge_sample = [c for c in df.columns if 'MGE' in c][:10]
        if mge_sample:
            print(f"🔍 Пример найденных MGE свойств: {', '.join(mge_sample)}")
            
    except Exception as e:
        print(f"❌ Ошибка сохранения Excel: {e}")

if __name__ == "__main__":
    input_file = "К01_КР_П_R19.ifc"
    output_file = "ifc_full_universal_report.xlsx"
    
    parse_ifc_universal(input_file, output_file)
