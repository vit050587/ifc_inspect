#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для подбора видов работ к элементам из таблицы full_elements.xlsx
на основе таблицы Перечень работ КР_new.xlsx

Алгоритм:
1. Берем элемент из таблицы элементов
2. Смотрим категорию уровня элемента в колонке "Категория уровня" таблицы full_elements.xlsx
3. Ищем в справочнике работ (Перечень работ КР_new.xlsx) в колонке "Категория уровня" наше значение (наличие текста)
4. Отбираем все виды работ и материалы соответствующие значениям уровня

Далее:
5. Смотрим на класс элемента в колонке "Ifc Class" таблицы элементов и сопоставляем с колонкой «IFC класс подраздела» в таблице справочника работ
6. Отбираем из уже ранее отобранных видов работ те что соответствуют еще и по классу

Далее:
7. Смотрим на имя в колонке Element Specific:Name таблицы элементов и семантически пытаемся сопоставить со значением колонки «Подраздел» из справочника работ 
   из тех которые уже были отобраны ранее двумя этапами (по уровню и классу) и оставляем те виды работ и материалов которые соответствуют

Далее:
8. Смотрим на отобранные виды работы по нашему выбранному элементу и если там в колонке «№ п/п» есть работы с одинаковым числом (целым), 
   то необходимо выбрать одно исходя из параметров в колонке «Параметризация» - смотрим условия как они считаются в колонке 
   «Формула расчёта объёмов работ и расхода материалов» и единицы измерения в колонке «Ед. изм» в которых нужно высчитывать 
   в справочнике и пробуем посчитать требуемые размерности у нашего выбранного элемента, после чего оставляем наиболее соответствующий вид работы.

Далее:
9. Смотрим на «№ п/п» с нецелочисленным значением (это означает материал для вида работы) если таковой имеется, 
   то смотрим на колонку «Наименование работ (целые числа) материалы для этих работ (дробные числа)» в справочнике 
   и смотреть на материал изготовления нашего элемента и считать его объем в той единице измерения в которой указано в колонке «Ед. изм»

Записывать списки работ по данному элементу в новый файл и брать следующий элемент чтобы так же пройти на соответствие
"""

import pandas as pd
import re
from typing import List, Dict, Optional, Tuple
import sys
import os


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
                           works_by_level: Dict, works_universal: pd.DataFrame) -> List[Dict]:
    """
    Находит все подходящие виды работ для элемента по алгоритму v6
    
    Алгоритм v6 (строго по порядку):
    1. Смотрим категорию уровня элемента в колонке "Категория уровня" таблицы элементов
    2. Ищем в справочнике работ в колонке "Категория уровня" наше значение (наличие текста)
    3. Отбираем все виды работ и материалы соответствующие значениям уровня
    4. Смотрим на класс элемента в колонке "Ifc Class" и сопоставляем с колонкой «IFC класс подраздела»
    5. Отбираем из уже ранее отобранных видов работ те что соответствуют еще и по классу
    6. Смотрим на имя элемента и семантически сопоставляем со значением колонки «Подраздел»
    7. Если есть работы с одинаковым целым № п/п, выбираем одну по параметрам параметризации
    8. Для дробных № п/п (материалы) смотрим материал элемента и считаем объем
    
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
    element_ifc = element_params['ifc_class']
    element_level = element_params['level']  # категория уровня из таблицы элементов
    element_name = element_params['name']
    element_subdivision = element_params.get('subdivision', '')  # Подраздел элемента если есть
    
    # Шаг 1-3: Фильтр по категории уровня
    # Ищем в справочнике работ в колонке "Категория уровня" наше значение (наличие текста)
    candidate_works = []
    
    # Проверяем точное совпадение категории уровня
    if element_level in works_by_level:
        for _, work_row in works_by_level[element_level].iterrows():
            candidate_works.append(work_row)
    
    # Если не нашли точного совпадения, пробуем найти частичное совпадение (наличие текста)
    if not candidate_works:
        for level_key, level_works in works_by_level.items():
            if element_level in level_key or level_key in element_level:
                for _, work_row in level_works.iterrows():
                    candidate_works.append(work_row)
                break
    
    # Если все еще нет кандидатов, пробуем универсальные работы
    if not candidate_works and works_universal is not None and len(works_universal) > 0:
        for _, work_row in works_universal.iterrows():
            candidate_works.append(work_row)
    
    # Шаг 4: Фильтр по IFC классу подраздела
    # Сопоставляем значение с колонкой «IFC класс подраздела» в таблице справочника работ
    if candidate_works and element_ifc:
        filtered_by_ifc = []
        element_ifc_lower = element_ifc.lower()
        
        for work_row in candidate_works:
            work_ifc_subclass = str(work_row.get('IFC класс подраздела', ''))
            if work_ifc_subclass:
                # Проверяем наличие IFC класса элемента в IFC классе подраздела работы
                work_ifcs = [x.strip().lower() for x in work_ifc_subclass.replace('\\n', ',').split(',')]
                if any(element_ifc_lower in wifc or wifc in element_ifc_lower for wifc in work_ifcs if wifc):
                    filtered_by_ifc.append(work_row)
        
        candidate_works = filtered_by_ifc
    
    # Шаг 5: Семантическое сопоставление по Подразделу
    # Смотрим на имя элемента и сопоставляем со значением колонки «Подраздел» из справочника работ
    if candidate_works and element_name:
        filtered_by_subdivision = []
        element_name_lower = element_name.lower()
        
        for work_row in candidate_works:
            work_subdivision = str(work_row.get('Подраздел', ''))
            if work_subdivision:
                # Семантическое сопоставление: проверяем наличие ключевых слов
                # Пример: "Фундаментная плита" в названии элемента → ищем "Фундаментная плита" в подразделе
                if work_subdivision.lower() in element_name_lower or element_name_lower in work_subdivision.lower():
                    filtered_by_subdivision.append(work_row)
                else:
                    # Проверяем по ключевым словам
                    subdivision_keywords = {
                        'фундаментн': ['фундамент', 'плита', 'фп'],
                        'стен': ['стена', 'стен', 'монолит'],
                        'колонн': ['колонн', 'кол'],
                        'балк': ['балк', 'ригел'],
                        'перекрыт': ['перекрыт', 'плит', 'покрыт'],
                        'лестниц': ['лестниц', 'марш', 'площадк'],
                        'приям': ['приям', 'приямок'],
                    }
                    
                    match_found = False
                    for subdiv_key, keywords in subdivision_keywords.items():
                        if subdiv_key in work_subdivision.lower():
                            if any(kw in element_name_lower for kw in keywords):
                                match_found = True
                                break
                    
                    if match_found:
                        filtered_by_subdivision.append(work_row)
                    else:
                        # Если не нашли совпадения по подразделу, оставляем работу если нет строгого фильтра
                        filtered_by_subdivision.append(work_row)
            else:
                filtered_by_subdivision.append(work_row)
        
        candidate_works = filtered_by_subdivision
    
    # Шаг 6-7: Проверка параметризации и выбор работ
    for work_row in candidate_works:
        if match_work_to_element(element_params, work_row):
            matched_works.append({
                'work_idx': work_row.name,
                'work_name': work_row.get('Наименование работ (целые числа) материалы для этих работ (дробные числа)', ''),
                'unit': work_row.get('Ед. изм', ''),
                'tsn_code': work_row.get('Шифр ТСН', ''),
                'description': work_row.get('Наименование расценки/ресурса', ''),
                'formula': work_row.get('Формула расчёта объёмов работ и расхода материалов', ''),
                'parametrization': work_row.get('Параметризация', ''),
                'pp_number': work_row.get('№ п/п', None),
            })
    
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


def load_and_prepare_works(works_file: str) -> Tuple[pd.DataFrame, Dict, pd.DataFrame]:
    """
    Загружает и подготавливает таблицу работ с группировкой по Категория уровня
    
    Алгоритм v6:
    1. Загружаем справочник работ (Перечень работ КР_new.xlsx)
    2. Группируем работы по колонке "Категория уровня" для быстрого поиска
    3. Также группируем по IFC классу подраздела внутри каждой категории уровня
    
    Возвращает:
    - df_works_valid: все валидные работы
    - works_by_level: словарь работ по категориям уровня (ключ = категория уровня, значение = DataFrame)
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
    
    # Группируем работы по Категория уровня для ускорения поиска
    # Алгоритм v6: смотрим на справочник работ и в колонке категория уровня ищем наше значение (наличие текста)
    works_by_level = {}  # словарь: ключ = категория уровня, значение = DataFrame с работами
    
    # Получаем уникальные категории уровня
    unique_levels = df_works_valid['Категория уровня'].dropna().unique()
    
    for level in unique_levels:
        if pd.notna(level) and level not in ['не моделируется', 'чаще всего не моделируется', '-', '']:
            level_str = str(level).strip()
            if level_str:
                # Фильтруем работы по категории уровня
                mask = df_works_valid['Категория уровня'] == level
                if mask.any():
                    works_by_level[level_str] = df_works_valid[mask].copy()
    
    # Работы без привязки к категории уровня (универсальные)
    universal_mask = df_works_valid['Категория уровня'].isin(['не моделируется', 'чаще всего не моделируется', '-', '']) | df_works_valid['Категория уровня'].isna()
    works_universal = df_works_valid[universal_mask].copy()
    
    print(f"Категорий уровня с работами: {len(works_by_level)}")
    for level, works in works_by_level.items():
        print(f"  - {level}: {len(works)} работ")
    print(f"Универсальных работ: {len(works_universal)}")
    
    return df_works_valid, works_by_level, works_universal


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
    print("Подбор видов работ к элементам (версия 6)")
    print("=" * 60)
    
    # Загружаем элементы
    print(f"\nЗагрузка таблицы элементов из {elements_file}...")
    df_elements = pd.read_excel(elements_file)
    print(f"Элементы: {len(df_elements)} строк")
    
    # Загружаем и подготавливаем работы с группировкой по Категория уровня
    df_works_valid, works_by_level, works_universal = load_and_prepare_works(works_file)
    results = []
    elements_with_works = set()
    elements_without_works = set()
    
    print("\nПодбор видов работ для элементов (по алгоритму v6)...")
    total = len(df_elements)
    valid_element_indices = set()
    
    for idx, element_row in enumerate(df_elements.itertuples(index=True)):
        # Получаем параметры элемента
        element_params = get_element_parameters(element_row)
        
        # Пропускаем элементы без работ
        if element_params['no_works']:
            continue
        
        # Пропускаем элементы с volume=0 и area=0
        volume = element_params.get('volume_m3', float('nan'))
        area = element_params.get('area_m2', float('nan'))
        if (pd.notna(volume) and volume == 0) or (pd.notna(area) and area == 0):
            continue
        
        valid_element_indices.add(idx)
        
        # Находим подходящие работы по алгоритму v6
        matched_works = find_works_for_element(
            element_row, df_works_valid, works_by_level, works_universal
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
