import ifcopenshell
import ifcopenshell.util.element
import ifcopenshell.util.unit
import json
import os
import math
from collections import defaultdict
import pandas as pd
import numpy as np
import re

# Пути к файлам
IFC_FILE = "data/ifc модель КР.ifc"
JSON_OUTPUT = "data/ifc_extracted_detailed.json"
EXCEL_OUTPUT = "Data/summary.xlsx"

def get_unit_scale(ifc_file):
    """Получить масштаб единиц из IFC файла"""
    try:
        units = ifcopenshell.util.unit.get_unit_scale(ifc_file)
        return units
    except:
        return 1.0

def calculate_volume_by_geometry(shape, unit_scale):
    """Расчет объема через bounding box с учетом масштаба"""
    try:
        bbox_min = shape.bounding_box.lower_corner
        bbox_max = shape.bounding_box.upper_corner
        
        dx = (bbox_max[0] - bbox_min[0]) * unit_scale
        dy = (bbox_max[1] - bbox_min[1]) * unit_scale
        dz = (bbox_max[2] - bbox_min[2]) * unit_scale
        
        return abs(dx * dy * dz)
    except Exception as e:
        return 0.0

def parse_concrete_properties(name, psets):
    """Парсинг свойств бетона из имени и PSets"""
    props = {
        "material": "Бетон",
        "class": "",
        "frost": "",
        "water": "",
        "reinforced": True,
        "type_detail": ""
    }
    
    full_text = f"{name} "
    for pset_name, pset_data in psets.items():
        for prop_name, prop_val in pset_data.items():
            if isinstance(prop_val, str):
                full_text += f"{prop_val} "
    
    full_text_upper = full_text.upper()
    
    # Проверка на не бетонные материалы подготовки
    if "ПЕСОК" in full_text_upper or "SAND" in full_text_upper:
        props["material"] = "Песок"
        props["type_detail"] = "Песок"
        props["reinforced"] = False
    elif "ЩЕБЕНЬ" in full_text_upper or "GRAVEL" in full_text_upper or "CRUSHED STONE" in full_text_upper:
        props["material"] = "Щебень"
        props["type_detail"] = "Щебень"
        props["reinforced"] = False
    elif "ПОДГОТОВКА" in full_text_upper and ("В7.5" in full_text_upper or "B7.5" in full_text_upper or "НЕАРМИРОВАННЫЙ" in full_text_upper):
        props["type_detail"] = "Бетон неармированный (подготовка)"
        props["reinforced"] = False

    # Поиск класса бетона
    class_match = re.search(r'[BВ]\s?(\d+(?:[.,]\d+)?)', full_text_upper)
    if class_match:
        val = class_match.group(1).replace(',', '.')
        props["class"] = f"В{val}"
        if float(val) < 15:
             props["reinforced"] = False

    # Поиск морозостойкости
    frost_match = re.search(r'F\s?(\d+)', full_text_upper)
    if frost_match:
        props["frost"] = f"F{frost_match.group(1)}"
    
    # Поиск водонепроницаемости
    water_match = re.search(r'W\s?(\d+)', full_text_upper)
    if water_match:
        props["water"] = f"W{water_match.group(1)}"
        
    return props

def classify_element(element, name):
    """Классификация элемента на более детальные подтипы"""
    element_type = element.is_a()
    name_upper = name.upper() if name else ""
    
    if element_type == "IfcSlab":
        # Приоритет 1: Лестничные площадки (должны быть отдельно)
        if "ЛЕСТНИЧ" in name_upper or "STAIR" in name_upper or "ПЛОЩАДКА" in name_upper:
            return "Лестница_Площадка"
        
        # Приоритет 2: Подготовка фундамента (песок, щебень, тощий бетон, стяжка)
        if "ПОДГОТОВКА" in name_upper or "СТЯЖКА" in name_upper:
            if "ПЕСОК" in name_upper or "SAND" in name_upper or "ЦЕМЕНТНО-ПЕСЧАНАЯ" in name_upper:
                return "Фундамент_Подготовка_Песок"
            elif "ЩЕБЕНЬ" in name_upper or "GRAVEL" in name_upper or "CRUSHED" in name_upper:
                return "Фундамент_Подготовка_Щебень"
            else:
                return "Фундамент_Подготовка_Бетон"
        
        # Приоритет 3: Фундаментные плиты (ищем явные маркеры фундамента)
        # Важно: проверяем наличие "ФП" но не в контексте подготовки
        if ("ФУНДАМЕНТ" in name_upper and "ПОДГОТОВКА" not in name_upper) or \
           "FOUNDATION SLAB" in name_upper or \
           "BASE SLAB" in name_upper or \
           "ФУНДАМЕНТНАЯ ПЛИТА" in name_upper:
            return "Фундамент_Плита"
        
        # Проверяем "ФП" отдельно - это может быть как фундаментная плита так и "перекрытие с ФП"
        # Если есть "ФП" и нет "перекрытие" - считаем фундаментом
        if "ФП" in name_upper and "ПЕРЕКРЫТИЕ" not in name_upper and "ПОДГОТОВКА" not in name_upper:
            return "Фундамент_Плита"
        
        # Приоритет 4: Перекрытия (по умолчанию для плит)
        return "Перекрытие_Плита"
    
    if element_type == "IfcWall":
        return "Стена"
    if element_type == "IfcColumn":
        return "Колонна"
    if element_type in ["IfcStair", "IfcStairFlight"]:
        return "Лестница_Марш"
    if element_type == "IfcBeam":
        return "Балка"
        
    return element_type

def extract_ifc_data(ifc_path):
    ifc_file = ifcopenshell.open(ifc_path)
    unit_scale = get_unit_scale(ifc_file)
    
    elements_data = []
    all_elements = ifc_file.by_type("IfcProduct")
    
    print(f"Всего элементов в файле: {len(all_elements)}")
    
    for element in all_elements:
        if not element.Representation:
            continue
            
        name = element.Name if element.Name else "Без имени"
        tag = element.Tag if element.Tag else ""
        element_type = element.is_a()
        
        psets = {}
        try:
            pset_definitions = ifcopenshell.util.element.get_psets(element)
            for pset_name, props in pset_definitions.items():
                psets[pset_name] = props
        except:
            pass
            
        volume = 0.0
        area = 0.0
        shape = None
        
        # Приоритет №1: Объем из свойств
        found_volume_in_props = False
        for pset_name, props in psets.items():
            if "NetVolume" in props:
                vol = props["NetVolume"]
                if isinstance(vol, (int, float)):
                    volume = vol
                    found_volume_in_props = True
                    break
            if "GrossVolume" in props and not found_volume_in_props:
                vol = props["GrossVolume"]
                if isinstance(vol, (int, float)):
                    if vol > 10000: 
                        volume = vol / 1000.0
                    else:
                        volume = vol
                    found_volume_in_props = True
                    break
        
        # Если объема нет в свойствах, пробуем рассчитать по геометрии
        if not found_volume_in_props:
            try:
                settings = ifcopenshell.geom.settings()
                settings.set(settings.USE_PYTHON_OPENCASCADE, True)
                shape = ifcopenshell.geom.create_shape(settings, element)
                if shape:
                    volume = calculate_volume_by_geometry(shape, unit_scale)
            except Exception as e:
                pass

        category = classify_element(element, name)
        material_props = parse_concrete_properties(name, psets)
        
        item = {
            "id": element.id(),
            "type": element_type,
            "category": category,
            "name": name,
            "tag": tag,
            "volume_m3": round(volume, 4),
            "area_m2": round(area, 4),
            "material": material_props["material"],
            "concrete_class": material_props["class"],
            "frost_resistance": material_props["frost"],
            "water_permeability": material_props["water"],
            "is_reinforced": material_props["reinforced"],
            "material_detail": material_props["type_detail"]
        }
        elements_data.append(item)
        
    return elements_data

def create_summary_table(data):
    """Создание сводной таблицы"""
    df = pd.DataFrame(data)
    
    group_cols = [
        "category", 
        "material", 
        "concrete_class", 
        "frost_resistance", 
        "water_permeability", 
        "is_reinforced"
    ]
    
    summary = df.groupby(group_cols, dropna=False).agg(
        count=("id", "count"),
        total_volume_m3=("volume_m3", "sum"),
        total_area_m2=("area_m2", "sum")
    ).reset_index()
    
    summary["total_volume_m3"] = summary["total_volume_m3"].round(2)
    summary["total_area_m2"] = summary["total_area_m2"].round(2)
    
    summary = summary.sort_values(by=["category", "material", "concrete_class", "total_volume_m3"], ascending=[True, True, True, False])
    
    return summary

def main():
    if not os.path.exists(IFC_FILE):
        print(f"Ошибка: Файл {IFC_FILE} не найден!")
        return

    print("Извлечение данных из IFC...")
    data = extract_ifc_data(IFC_FILE)
    
    with open(JSON_OUTPUT, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"Подробные данные сохранены в {JSON_OUTPUT}")
    
    print("Формирование сводной таблицы...")
    summary_df = create_summary_table(data)
    
    os.makedirs(os.path.dirname(EXCEL_OUTPUT), exist_ok=True)
    summary_df.to_excel(EXCEL_OUTPUT, index=False, sheet_name="Сводная")
    print(f"Сводная таблица сохранена в {EXCEL_OUTPUT}")
    
    # Удаление листа сравнения если он существует (требование пользователя)
    try:
        with pd.ExcelWriter(EXCEL_OUTPUT, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            # Перезаписываем только лист "Сводная", удаляя другие если они были добавлены ранее
            book = writer.book
            if "Сравнение с целью" in book.sheetnames:
                del book["Сравнение с целью"]
            summary_df.to_excel(writer, index=False, sheet_name="Сводная")
    except Exception as e:
        print(f"Предупреждение при очистке лишних листов: {e}")

    print("\n--- ИТОГИ ПО КАТЕГОРИЯМ ---")
    totals = summary_df.groupby("category").agg(
        total_count=("count", "sum"),
        total_vol=("total_volume_m3", "sum"),
        total_area=("total_area_m2", "sum")
    )
    print(totals.to_string())
    
    targets = {
        "Стена": 3857.89,
        "Перекрытие_Плита": 4294.95,
        "Колонна": 94.28,
        "Фундамент_Плита": 3042.14,
        "Фундамент_Подготовка_Бетон": 619.08,
        "Фундамент_Подготовка_Песок": 0.0
    }
    
    print("\n--- СРАВНЕНИЕ С ЦЕЛЕВЫМИ ЗНАЧЕНИЯМИ ---")
    for cat, target_vol in targets.items():
        if cat in totals.index:
            actual_vol = totals.loc[cat, "total_vol"]
            diff = actual_vol - target_vol
            percent = (actual_vol / target_vol) * 100 if target_vol != 0 else 0
            print(f"{cat}: Факт={actual_vol:.2f}, Цель={target_vol:.2f}, Разница={diff:.2f} ({percent:.1f}%)")
        else:
            found_similar = [k for k in totals.index if cat.split('_')[0] in k]
            if found_similar:
                 print(f"{cat}: Не найдено точное совпадение. Найдено похожие: {found_similar}. Цель={target_vol:.2f}")
            else:
                print(f"{cat}: Не найдено в модели! Цель={target_vol:.2f}")

if __name__ == "__main__":
    main()
