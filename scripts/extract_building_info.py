#!/usr/bin/env python3
"""
IFC Building Info Extractor - Скрипт для извлечения информации об объекте строительства из IFC файла.
Извлекает: адрес, высоту от нулевой отметки, этажность, площадь застройки и другие параметры здания.
"""

import ifcopenshell
import ifcopenshell.util.element
import json
import sys
from pathlib import Path
from typing import Dict, List, Any, Optional


def get_property_value(psets: Dict, pset_name: str, prop_name: str) -> Any:
    """Получить значение свойства из psets"""
    if pset_name in psets:
        props = psets[pset_name]
        if prop_name in props:
            val = props[prop_name]
            if val is not None:
                return val
    return None


def extract_building_info(ifc_path: str) -> Dict[str, Any]:
    """
    Извлечь информацию об объекте строительства из IFC файла.
    
    Args:
        ifc_path: Путь к IFC файлу
        
    Returns:
        Словарь с информацией о здании
    """
    result = {
        "file": str(ifc_path),
        "project": {},
        "building": {},
        "storeys": [],
        "summary": {}
    }
    
    try:
        f = ifcopenshell.open(ifc_path)
    except Exception as e:
        return {"error": f"Не удалось открыть файл: {str(e)}"}
    
    # === Извлечение данных проекта ===
    projects = f.by_type("IfcProject")
    if projects:
        project = projects[0]
        result["project"]["name"] = getattr(project, 'Name', None)
        result["project"]["long_name"] = getattr(project, 'LongName', None)
        result["project"]["description"] = getattr(project, 'Description', None)
    
    # === Извлечение данных здания ===
    buildings = f.by_type("IfcBuilding")
    if buildings:
        building = buildings[0]
        
        # Базовые атрибуты
        result["building"]["name"] = getattr(building, 'Name', None)
        result["building"]["long_name"] = getattr(building, 'LongName', None)
        result["building"]["elevation_ref_height"] = getattr(building, 'ElevationOfRefHeight', None)
        result["building"]["elevation_terrain"] = getattr(building, 'ElevationOfTerrain', None)
        
        # Извлечение свойств через IsDefinedBy
        psets_data = {}
        if hasattr(building, 'IsDefinedBy'):
            for rel in building.IsDefinedBy:
                if hasattr(rel, 'RelatingPropertyDefinition'):
                    pset_def = rel.RelatingPropertyDefinition
                    pset_name = getattr(pset_def, 'Name', 'Unknown')
                    
                    if hasattr(pset_def, 'HasProperties'):
                        for prop in pset_def.HasProperties:
                            prop_name = getattr(prop, 'Name', None)
                            prop_value = getattr(prop, 'NominalValue', None)
                            
                            if prop_name and prop_value is not None:
                                # Извлекаем само значение из NominalValue
                                if hasattr(prop_value, 'wrappedValue'):
                                    actual_value = prop_value.wrappedValue
                                else:
                                    actual_value = prop_value
                                
                                if pset_name not in psets_data:
                                    psets_data[pset_name] = {}
                                psets_data[pset_name][prop_name] = actual_value
        
        # Сохраняем все найденные свойства
        result["building"]["properties"] = psets_data
        
        # Ключевые параметры для summary
        if "ExpCheck_Building" in psets_data:
            exp_props = psets_data["ExpCheck_Building"]
            result["building"]["address"] = exp_props.get("MGE_BuildingAddress")
            result["building"]["customer"] = exp_props.get("MGE_Customer")
            result["building"]["designer"] = exp_props.get("MGE_Designer")
            result["building"]["project_name"] = exp_props.get("MGE_ProjectName")
            result["building"]["object_name"] = exp_props.get("MGE_ObjectName")
            result["building"]["project_code"] = exp_props.get("MGE_ProjectCode")
            result["building"]["korpus"] = exp_props.get("MGE_Korpus")
            result["building"]["section"] = exp_props.get("MGE_Section")
            result["building"]["num_of_section"] = exp_props.get("MGE_NumOfSection")
            result["building"]["functional_use"] = exp_props.get("MGE_FunctionalUse")
            
            # Высота от нулевой отметки (разница между референсной высотой и уровнем земли)
            ref_height = exp_props.get("MGE_ElevationOfRefHeight")
            terrain_height = exp_props.get("MGE_ElevationOfTerrain")
            if ref_height is not None and terrain_height is not None:
                # Обычно MGE_ElevationOfRefHeight - это абсолютная высота в мм
                # Конвертируем в метры
                result["building"]["ref_height_m"] = float(ref_height) / 1000.0 if ref_height else None
                result["building"]["terrain_height_m"] = float(terrain_height) / 1000.0 if terrain_height else None
        
        if "Pset_BuildingCommon" in psets_data:
            common_props = psets_data["Pset_BuildingCommon"]
            result["building"]["number_of_storeys"] = common_props.get("NumberOfStoreys")
            result["building"]["construction_method"] = common_props.get("ConstructionMethod")
            result["building"]["fire_protection_class"] = common_props.get("FireProtectionClass")
            result["building"]["is_landmarked"] = common_props.get("IsLandmarked")
    
    # === Извлечение информации об этажах ===
    storeys = f.by_type("IfcBuildingStorey")
    above_ground_storeys = []
    below_ground_storeys = []
    
    for storey in storeys:
        elevation = getattr(storey, 'Elevation', None)
        storey_info = {
            "name": getattr(storey, 'Name', None),
            "elevation": float(elevation) / 1000.0 if elevation is not None else None  # Конвертация в метры
        }
        
        # Определяем, надземный или подземный этаж
        if elevation is not None:
            if elevation > 0:
                above_ground_storeys.append(storey_info)
            else:
                below_ground_storeys.append(storey_info)
        else:
            above_ground_storeys.append(storey_info)
        
        result["storeys"].append(storey_info)
    
    # Сортировка этажей по высоте
    result["storeys"].sort(key=lambda x: x["elevation"] if x["elevation"] is not None else 0)
    
    # === Формирование сводной информации (summary) ===
    total_storeys = len(above_ground_storeys) + len(below_ground_storeys)
    
    # Расчет высоты здания
    max_elevation = None
    min_elevation = None
    
    for storey in result["storeys"]:
        elev = storey["elevation"]
        if elev is not None:
            if max_elevation is None or elev > max_elevation:
                max_elevation = elev
            if min_elevation is None or elev < min_elevation:
                min_elevation = elev
    
    # Общая высота здания от нулевой отметки
    if max_elevation is not None:
        result["summary"]["total_height_from_zero_m"] = round(max_elevation, 3)
    
    # Глубина подземной части
    if min_elevation is not None and min_elevation < 0:
        result["summary"]["underground_depth_m"] = round(abs(min_elevation), 3)
    
    # Полная высота (от низа подземной части до верха)
    if max_elevation is not None and min_elevation is not None:
        result["summary"]["total_height_full_m"] = round(max_elevation - min_elevation, 3)
    
    # Этажность
    result["summary"]["above_ground_storeys"] = len(above_ground_storeys)
    result["summary"]["below_ground_storeys"] = len(below_ground_storeys)
    result["summary"]["total_storeys"] = total_storeys
    
    # Если есть данные из Pset_BuildingCommon
    if result["building"].get("number_of_storeys"):
        result["summary"]["storeys_from_pset"] = result["building"]["number_of_storeys"]
    
    # Адрес
    result["summary"]["address"] = result["building"].get("address")
    
    # Попытка вычислить площадь застройки
    # Ищем все элементы на первом этаже (с минимальной положительной высотой или нулевой)
    ground_level_elements = []
    if result["storeys"]:
        # Находим уровень земли (первый надземный этаж)
        ground_level = None
        for storey in result["storeys"]:
            if storey["elevation"] is not None and storey["elevation"] >= 0:
                ground_level = storey["elevation"]
                break
        
        if ground_level is not None:
            # Ищем плиты перекрытия или другие элементы на этом уровне
            slabs = f.by_type("IfcSlab")
            for slab in slabs[:100]:  # Ограничиваем количество для производительности
                psets = ifcopenshell.util.element.get_psets(slab)
                storey_name = psets.get("Storey", None)
                
                # Проверяем, относится ли элемент к уровню земли
                if storey_name:
                    for storey in result["storeys"]:
                        if storey["name"] == storey_name and abs(storey["elevation"] - ground_level) < 0.1:
                            # Получаем геометрию плиты
                            try:
                                shape = ifcopenshell.util.element.get_representation(slab, 'Body')
                                if shape:
                                    bbox_min = shape.bounding_box.lower_corner
                                    bbox_max = shape.bounding_box.upper_corner
                                    
                                    dx = (bbox_max[0] - bbox_min[0]) / 1000.0  # в метрах
                                    dy = (bbox_max[1] - bbox_min[1]) / 1000.0  # в метрах
                                    
                                    area = abs(dx * dy)
                                    if area > 10:  # Фильтр маленьких элементов
                                        ground_level_elements.append({
                                            "name": getattr(slab, 'Name', None),
                                            "area": round(area, 2)
                                        })
                            except:
                                pass
            
            # Вычисляем примерную площадь застройки как сумму площадей плит на уровне земли
            if ground_level_elements:
                # Берем максимальную площадь или сумму уникальных областей
                areas = [el["area"] for el in ground_level_elements]
                if areas:
                    # Для простоты берем максимальную площадь как оценку площади застройки
                    result["summary"]["footprint_area_approx_m2"] = round(max(areas), 2)
                    result["summary"]["footprint_elements_count"] = len(ground_level_elements)
    
    return result


def print_summary(info: Dict[str, Any]):
    """Вывести краткую сводку по зданию"""
    print("\n" + "="*80)
    print("ИНФОРМАЦИЯ ОБ ОБЪЕКТЕ СТРОИТЕЛЬСТВА")
    print("="*80)
    
    if "error" in info:
        print(f"ОШИБКА: {info['error']}")
        return
    
    summary = info.get("summary", {})
    building = info.get("building", {})
    project = info.get("project", {})
    
    print("\n📋 ПРОЕКТ:")
    print(f"   Название: {project.get('name', 'Н/Д')}")
    if project.get('long_name'):
        print(f"   Полное название: {project.get('long_name')}")
    
    print("\n🏢 ЗДАНИЕ:")
    print(f"   Название: {building.get('name', 'Н/Д')}")
    print(f"   Адрес: {summary.get('address', building.get('address', 'Н/Д'))}")
    
    print("\n📏 ВЫСОТНЫЕ ПАРАМЕТРЫ:")
    if summary.get('total_height_from_zero_m'):
        print(f"   Высота от нулевой отметки: {summary['total_height_from_zero_m']} м")
    if summary.get('underground_depth_m'):
        print(f"   Глубина подземной части: {summary['underground_depth_m']} м")
    if summary.get('total_height_full_m'):
        print(f"   Полная высота (от низа до верха): {summary['total_height_full_m']} м")
    if building.get('ref_height_m'):
        print(f"   Абсолютная высота референсной отметки: {building['ref_height_m']} м")
    if building.get('terrain_height_m'):
        print(f"   Абсолютная высота уровня земли: {building['terrain_height_m']} м")
    
    print("\n🏗️ ЭТАЖНОСТЬ:")
    print(f"   Надземных этажей: {summary.get('above_ground_storeys', 'Н/Д')}")
    print(f"   Подземных этажей: {summary.get('below_ground_storeys', 'Н/Д')}")
    print(f"   Всего этажей: {summary.get('total_storeys', 'Н/Д')}")
    if summary.get('storeys_from_pset'):
        print(f"   Этажность (из Pset): {summary['storeys_from_pset']}")
    
    print("\n📐 ПЛОЩАДЬ ЗАСТРОЙКИ:")
    if summary.get('footprint_area_approx_m2'):
        print(f"   Площадь застройки (прибл.): {summary['footprint_area_approx_m2']} м²")
        print(f"   Количество элементов для расчета: {summary.get('footprint_elements_count', 0)}")
    else:
        print("   Не удалось определить (требуется дополнительный анализ геометрии)")
    
    print("\n📊 ДОПОЛНИТЕЛЬНАЯ ИНФОРМАЦИЯ:")
    if building.get('project_code'):
        print(f"   Код проекта: {building['project_code']}")
    if building.get('customer'):
        print(f"   Заказчик: {building['customer']}")
    if building.get('designer'):
        print(f"   Проектировщик: {building['designer']}")
    if building.get('functional_use'):
        print(f"   Функциональное назначение: {building['functional_use']}")
    if building.get('construction_method'):
        print(f"   Метод строительства: {building['construction_method']}")
    if building.get('fire_protection_class'):
        print(f"   Класс пожарной защиты: {building['fire_protection_class']}")
    
    print("\n" + "="*80)


def main():
    """Основная функция"""
    if len(sys.argv) < 2:
        # Путь по умолчанию
        ifc_path = "/workspace/data/ifc модель КР.ifc"
    else:
        ifc_path = sys.argv[1]
    
    # Проверка существования файла
    if not Path(ifc_path).exists():
        print(f"ОШИБКА: Файл не найден: {ifc_path}")
        sys.exit(1)
    
    print(f"Обработка файла: {ifc_path}")
    
    # Извлечение информации
    info = extract_building_info(ifc_path)
    
    # Вывод сводки
    print_summary(info)
    
    # Сохранение полного результата в JSON
    output_path = Path(ifc_path).parent / "building_info.json"
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(info, f, ensure_ascii=False, indent=2)
    
    print(f"\n✅ Полный результат сохранен в: {output_path}")


if __name__ == "__main__":
    main()
