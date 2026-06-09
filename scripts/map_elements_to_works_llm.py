#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для подбора видов работ к элементам из таблицы full_elements.xlsx
на основе таблицы Перечень работ КР_new.xlsx с использованием LLM (Ollama + Gemma2:27b)

Алгоритм с применением LLM (строго по порядку):
1. Определить уровень элемента (Подземный/Надземный) для фильтрации разделов работ
2. По классу элемента выбрать в отфильтрованном по уровню разделе соответствующий подраздел видов работ
   с таким же классом (Стены->Стены, Колонны->Колонны, Лестницы->Лестницы, Перекрытия/Покрытия->Плиты перекрытия/Покрытия, Фундаменты плиты->Фундаментная плита)
3. Сопоставить работы согласно параметрам видов работ и соответствующим параметрам элемента 
   (размеры, площадь, площадь сечения, ширина или толщина) - только в тех видах работ где они указаны
4. Из отобранных работ подобрать виды работ по материалам и его параметрам (класс бетона, морозостойкость, водонепроницаемость)
   из которого состоит элемент (если такие условия по материалу присутствуют у вида работы)

Обрабатываются только классы: IfcWall, IfcColumn, IfcStairFlight, IfcSlab, IfcSlabFoundation

Использует Ollama API для подключения к локальной модели.
"""

import os
import json
import requests
import pandas as pd
import re
from typing import List, Dict, Optional, Tuple
import sys


# Настройки из переменных окружения
OLLAMA_BASE_URL = os.getenv("OLLAMA_BASE_URL", "http://localhost:11434")
LLM_MODEL = os.getenv("DRAWING_VALIDATION_MODEL", "gemma2:27b")

# Классы элементов для обработки
TARGET_IFC_CLASSES = ['IfcWall', 'IfcColumn', 'IfcStairFlight', 'IfcSlab', 'IfcSlabFoundation']

# Маппинг классов элементов к названиям подразделов
CLASS_TO_SUBSECTION_MAP = {
    'IfcWall': ['Стен', 'Стена'],
    'IfcColumn': ['Колонн', 'Колонна'],
    'IfcStairFlight': ['Лестниц', 'Лестничн', 'Лестница'],
    'IfcSlab': ['Перекрыт', 'Покрыт', 'Плит', 'Плита'],
    'IfcSlabFoundation': ['Фундаментн', 'Фундамент', 'Плита', 'фундаментная плита']
}

# Маппинг уровней к разделам
LEVEL_TO_SECTION_MAP = {
    'Подземный': ['Подземн', 'подземная', 'Фундамент'],
    'Надземный': ['Надземн', 'надземная', 'Цоколь']
}


def parse_concrete_grade(material: str, characteristics: str) -> Optional[str]:
    """Извлекает класс бетона из материала или характеристик"""
    if pd.isna(material) and pd.isna(characteristics):
        return None
    
    text = ""
    if pd.notna(material):
        text += str(material) + " "
    if pd.notna(characteristics):
        text += str(characteristics)
    
    match = re.search(r'[ВB](\d+)', text)
    if match:
        return f"В{match.group(1)}"
    
    return None


def parse_freeze_durability(material: str, characteristics: str) -> Optional[str]:
    """Извлекает марку морозостойкости F"""
    text = ""
    if pd.notna(material):
        text += str(material) + " "
    if pd.notna(characteristics):
        text += str(characteristics)
    
    match = re.search(r'F(\d{3})', text)
    if match:
        return f"F{match.group(1)}"
    
    return None


def parse_water_resist(material: str, characteristics: str) -> Optional[str]:
    """Извлекает марку водонепроницаемости W"""
    text = ""
    if pd.notna(material):
        text += str(material) + " "
    if pd.notna(characteristics):
        text += str(characteristics)
    
    match = re.search(r'W(\d+)', text)
    if match:
        return f"W{match.group(1)}"
    
    return None


def determine_ifc_class(row: pd.Series) -> str:
    """Определяет IFC класс элемента по наименованию и другим параметрам"""
    name = str(row.get('Наименование', ''))
    
    # Только целевые классы для обработки
    if 'Колонн' in name or 'НесКол' in name:
        return 'IfcColumn'
    elif 'Стен' in name or 'Диафрагм' in name:
        return 'IfcWall'
    elif 'Плит' in name or 'Перекрыт' in name or 'Покрыт' in name:
        if 'Фундамент' in name or 'ФП' in name or 'фундаментн' in name.lower():
            return 'IfcSlabFoundation'
        return 'IfcSlab'
    elif 'Лестничн' in name and 'Марш' in name:
        return 'IfcStairFlight'
    elif 'Фундамент' in name and 'плит' not in name.lower():
        return 'IfcSlabFoundation'  # Treat footing as slab foundation
    
    return ''


def get_element_level(row: pd.Series) -> str:
    """Определяет уровень элемента (подземный/надземный)"""
    level = row.get('Подземный/Надземный', '')
    if pd.isna(level):
        return 'Не определено'
    
    level_str = str(level).strip()
    if 'подземн' in level_str.lower() or 'подз.' in level_str.lower():
        return 'Подземный'
    elif 'надземн' in level_str.lower() or 'надз.' in level_str.lower():
        return 'Надземный'
    else:
        return level_str


def is_target_element(ifc_class: str) -> bool:
    """Проверяет, является ли элемент целевым классом для обработки"""
    return ifc_class in TARGET_IFC_CLASSES


def get_element_parameters(row: pd.Series) -> Dict:
    """Извлекает параметры элемента из строки таблицы"""
    material = row.get('Материал', '')
    characteristics = row.get('Характеристики материала', '')
    
    ifc_class = determine_ifc_class(row)
    name = row.get('Наименование', '')
    
    params = {
        'ifc_class': ifc_class,
        'name': name,
        'level': get_element_level(row),
        'concrete_grade': parse_concrete_grade(material, characteristics),
        'freeze_durability': parse_freeze_durability(material, characteristics),
        'water_resist': parse_water_resist(material, characteristics),
        'thickness_m': row.get('Толщина, м', float('nan')),
        'width_m': row.get('Ширина, м', float('nan')),
        'height_m': row.get('Высота, м', float('nan')),
        'length_m': row.get('Длина, м', float('nan')),
        'volume_m3': row.get('Объем, м³', float('nan')),
        'area_m2': row.get('Площадь, м²', float('nan')),
        'material': material,
        'characteristics': characteristics,
        'is_target': is_target_element(ifc_class)
    }
    
    return params


def format_element_for_prompt(element_params: Dict) -> str:
    """Форматирует параметры элемента для промпта LLM"""
    lines = [
        f"Наименование: {element_params['name']}",
        f"IFC класс: {element_params['ifc_class']}",
        f"Уровень: {element_params['level']}",
    ]
    
    if pd.notna(element_params['concrete_grade']):
        lines.append(f"Класс бетона: {element_params['concrete_grade']}")
    
    if pd.notna(element_params['freeze_durability']):
        lines.append(f"Морозостойкость: {element_params['freeze_durability']}")
    
    if pd.notna(element_params['water_resist']):
        lines.append(f"Водонепроницаемость: {element_params['water_resist']}")
    
    if pd.notna(element_params['thickness_m']):
        lines.append(f"Толщина: {element_params['thickness_m']} м")
    
    if pd.notna(element_params['width_m']) and pd.notna(element_params['height_m']):
        lines.append(f"Сечение: {element_params['width_m']} x {element_params['height_m']} м")
    
    if pd.notna(element_params['volume_m3']):
        lines.append(f"Объем: {element_params['volume_m3']} м³")
    
    if pd.notna(element_params['area_m2']):
        lines.append(f"Площадь: {element_params['area_m2']} м²")
    
    if element_params['material']:
        lines.append(f"Материал: {element_params['material']}")
    
    if element_params['characteristics']:
        lines.append(f"Характеристики: {element_params['characteristics']}")
    
    return "\n".join(lines)


def find_section_range_by_level(works_df: pd.DataFrame, level: str) -> tuple:
    """
    Находит диапазон строк в таблице работ, соответствующий уровню элемента
    
    Args:
        works_df: DataFrame с работами
        level: Уровень элемента ('Подземный', 'Надземный')
        
    Returns:
        (start_idx, end_idx) или None если не найдено
    """
    # Определяем ключевые слова для поиска раздела по уровню
    if level == 'Подземный':
        keywords = ['Подземн', 'подземная', 'Фундамент']
    elif level == 'Надземный':
        keywords = ['Надземн', 'надземная', 'Цоколь']
    else:
        return (0, len(works_df))
    
    start_idx = None
    end_idx = None
    
    for idx, row in works_df.iterrows():
        num = str(row.get('№ п/п', '')) if pd.notna(row.get('№ п/п', '')) else ''
        name = str(row.get('Наименование работ', '')) if pd.notna(row.get('Наименование работ', '')) else ''
        
        # Ищем начало раздела по ключевым словам в № п/п или Наименовании
        text_to_check = num + ' ' + name
        if 'Раздел' in num:
            # Это новый раздел - сбрасываем
            if start_idx is not None and end_idx is None:
                # Предыдущий раздел закончился
                end_idx = idx - 1
        
        # Проверяем, подходит ли раздел по уровню
        has_keyword = any(kw.lower() in text_to_check.lower() for kw in keywords)
        if 'Раздел' in num and has_keyword:
            start_idx = idx
            end_idx = None  # Сбрасываем конец, пока ищем следующий раздел
    
    # Если нашли начало, но не нашли конец - берем до конца файла или до следующего раздела
    if start_idx is not None and end_idx is None:
        # Ищем следующий раздел
        for idx in range(start_idx + 1, len(works_df)):
            row = works_df.iloc[idx]
            num = str(row.get('№ п/п', '')) if pd.notna(row.get('№ п/п', '')) else ''
            if 'Раздел' in num:
                end_idx = idx - 1
                break
        if end_idx is None:
            end_idx = len(works_df) - 1
    
    if start_idx is not None:
        return (start_idx, end_idx)
    
    # Если не нашли точного совпадения, возвращаем весь файл
    return (0, len(works_df) - 1)


def find_subsection_by_class(works_df: pd.DataFrame, ifc_class: str, start_idx: int, end_idx: int) -> tuple:
    """
    Находит подраздел в пределах диапазона, соответствующий классу элемента
    
    Args:
        works_df: DataFrame с работами
        ifc_class: IFC класс элемента
        start_idx: Начальный индекс диапазона
        end_idx: Конечный индекс диапазона
        
    Returns:
        (sub_start_idx, sub_end_idx) или None если не найдено
    """
    # Получаем ключевые слова для данного класса
    keywords = CLASS_TO_SUBSECTION_MAP.get(ifc_class, [])
    if not keywords:
        return (start_idx, end_idx)
    
    sub_start = None
    sub_end = None
    
    for idx in range(start_idx, min(end_idx + 1, len(works_df))):
        row = works_df.iloc[idx]
        num = str(row.get('№ п/п', '')) if pd.notna(row.get('№ п/п', '')) else ''
        
        # Ищем Подраздел с нужными ключевыми словами
        if 'Подраздел' in num:
            # Если уже нашли предыдущий подраздел, завершаем его
            if sub_start is not None and sub_end is None:
                sub_end = idx - 1
            
            # Проверяем, подходит ли этот подраздел по классу
            has_keyword = any(kw.lower() in num.lower() for kw in keywords)
            if has_keyword:
                sub_start = idx
                sub_end = None
        
        # Также проверяем IFC класс в строке
        work_ifc = str(row.get('IFC класс', '')) if pd.notna(row.get('IFC класс', '')) else ''
        if ifc_class.lower() in work_ifc.lower() and sub_start is None:
            # Если нашли строку с matching IFC классом до подразделения
            pass
    
    # Если нашли начало, но не нашли конец - берем до конца диапазона
    if sub_start is not None and sub_end is None:
        sub_end = end_idx
    
    if sub_start is not None:
        return (sub_start, sub_end)
    
    # Если не нашли подраздел, возвращаем весь диапазон
    return (start_idx, end_idx)


def select_candidate_works_with_structure(element_params: Dict, works_df: pd.DataFrame) -> pd.DataFrame:
    """
    Отбирает кандидатуры работ по структуре документа:
    1. Сначала фильтруем по уровню (Подземный/Надземный) - находим нужный раздел
    2. Затем по классу элемента - находим нужный подраздел
    3. В пределах подраздела берем все работы с подходящим IFC классом
    
    Args:
        element_params: Параметры элемента
        works_df: DataFrame с работами
        
    Returns:
        DataFrame с кандидатами работ
    """
    if not element_params['is_target']:
        return pd.DataFrame()
    
    element_ifc = element_params['ifc_class']
    element_level = element_params['level']
    
    print(f"  Поиск работ для {element_ifc} ({element_level})...")
    
    # Шаг 1: Находим раздел по уровню
    section_start, section_end = find_section_range_by_level(works_df, element_level)
    print(f"  Раздел: строки {section_start}-{section_end}")
    
    # Шаг 2: Находим подраздел по классу элемента
    sub_start, sub_end = find_subsection_by_class(works_df, element_ifc, section_start, section_end)
    print(f"  Подраздел: строки {sub_start}-{sub_end}")
    
    # Шаг 3: В пределах подраздела отбираем работы
    candidates = []
    for idx in range(sub_start, min(sub_end + 1, len(works_df))):
        work_row = works_df.iloc[idx]
        work_name = str(work_row.get('Наименование работ', '')) if pd.notna(work_row.get('Наименование работ', '')) else ''
        
        # Пропускаем заголовки подразделов
        if 'Подраздел' in work_name or not work_name.strip():
            continue
        
        # Проверяем, что это не заголовок материала (обычно следует после основной работы)
        # Основная работа обычно имеет номер без точки (1, 2, 3) или с буквой (1а, 1б)
        num = str(work_row.get('№ п/п', '')) if pd.notna(work_row.get('№ п/п', '')) else ''
        
        # Добавляем работу в кандидаты
        candidates.append(work_row)
    
    print(f"  Найдено кандидатов: {len(candidates)}")
    return pd.DataFrame(candidates) if candidates else pd.DataFrame()


def match_works_with_llm(element_params: Dict, candidate_works: pd.DataFrame) -> List[Dict]:
    """
    Использует LLM для подбора работ к элементу из списка кандидатов
    
    Алгоритм LLM:
    1. Фильтрация по параметрам (толщина, ширина, высота, площадь, объем)
    2. Фильтрация по материалу (класс бетона, морозостойкость, водонепроницаемость)
    
    Args:
        element_params: Параметры элемента
        candidate_works: DataFrame с кандидатами работ
        
    Returns:
        Список подобранных работ
    """
    if candidate_works.empty or not element_params['is_target']:
        return []
    
    # Формируем описание элемента
    element_description = format_element_for_prompt(element_params)
    
    # Формируем список работ для анализа
    works_list = []
    for idx, work_row in candidate_works.iterrows():
        work_info = {
            'idx': idx,
            'name': work_row.get('Наименование работ', ''),
            'unit': work_row.get('Ед. изм', ''),
            'tsn_code': work_row.get('Шифр ТСН', ''),
            'description': work_row.get('Наименование расценки/ресурса', ''),
            'formula': work_row.get('Формула расчёта объёмов работ и расхода материалов', ''),
            'parametrization': work_row.get('Параметризация', ''),
        }
        works_list.append(work_info)
    
    # Ограничиваем количество работ для одного запроса
    max_works_per_request = 30
    
    matched_works = []
    
    for i in range(0, len(works_list), max_works_per_request):
        batch = works_list[i:i + max_works_per_request]
        
        # Формируем промпт для LLM с акцентом на параметры и материалы
        works_text = ""
        for j, work in enumerate(batch):
            works_text += f"{j + 1}. [{work['idx']}] {work['name']}\n"
            if work['parametrization'] and pd.notna(work['parametrization']) and work['parametrization'] not in ['-', 'nan', '']:
                works_text += f"   Параметры: {work['parametrization']}\n"
            if work['description'] and pd.notna(work['description']):
                works_text += f"   Описание: {work['description']}\n"
            works_text += "\n"
        
        prompt = f"""Ты эксперт по строительным работам и сметному нормированию.
Задача: Подобрать подходящие виды работ для строительного элемента.

ЭЛЕМЕНТ:
{element_description}

СПИСОК РАБОТ-КАНДИДАТОВ (уже отобраны по уровню и типу конструкции):
{works_text}

ИНСТРУКЦИЯ ПО ПОДБОРУ (строго по порядку):

ШАГ 1 - ФИЛЬТРАЦИЯ ПО ПАРАМЕТРАМ:
Проверь параметры элемента против ограничений в графе "Параметризация":
- Если указано t<XXX мм - сравни с толщиной элемента
- Если указано a<XXX мм - сравни с размером сечения (ширина/высота)  
- Если указано S<XXX м2 - сравни с площадью элемента
- Если указано V>XXX м3 - сравни с объемом элемента
Отбери работы где параметры элемента соответствуют ограничениям.

ШАГ 2 - ФИЛЬТРАЦИЯ ПО МАТЕРИАЛАМ:
Из оставшихся работ выбери те где требования к материалу совпадают с элементом:
- Класс бетона (В25, В30, В35 и т.д.)
- Морозостойкость (F100, F150, F200 и т.д.)
- Водонепроницаемость (W4, W6, W8 и т.д.)
Если у работы нет требований к материалу - она подходит любому элементу своего типа.

ВАЖНО:
- Не выбирай работы которые явно не подходят по смыслу
- Учитывай параметрические ограничения строго
- Если параметров нет в работе - она может подойти
- Исключи дубликаты (работы с одинаковым названием)

Ответь ТОЛЬКО списком номеров работ из списка выше которые подходят, в формате:
[1, 3, 5]

Если ни одна работа не подходит, ответь: []

Номера подходящих работ:"""

        try:
            response = requests.post(
                f"{OLLAMA_BASE_URL}/api/generate",
                json={
                    "model": LLM_MODEL,
                    "prompt": prompt,
                    "stream": False,
                    "options": {
                        "temperature": 0.1,
                        "top_p": 0.9,
                        "num_predict": 200
                    }
                },
                timeout=180
            )
            
            if response.status_code == 200:
                result = response.json()
                answer = result.get("response", "").strip()
                
                # Парсим ответ - извлекаем номера работ
                match = re.search(r'\[([^\]]*)\]', answer)
                if match:
                    indices_str = match.group(1)
                    selected_indices = []
                    
                    for idx_str in indices_str.split(','):
                        idx_str = idx_str.strip()
                        if idx_str.isdigit():
                            idx = int(idx_str) - 1  # Преобразуем из 1-based в 0-based
                            if 0 <= idx < len(batch):
                                selected_indices.append(idx)
                    
                    # Добавляем выбранные работы в результат
                    for idx in selected_indices:
                        work = batch[idx]
                        matched_works.append({
                            'work_idx': work['idx'],
                            'work_name': work['name'],
                            'unit': work['unit'],
                            'tsn_code': work['tsn_code'],
                            'description': work['description'],
                            'formula': work['formula'],
                            'parametrization': work['parametrization'],
                        })
                else:
                    print(f"   ⚠️  LLM вернул неожиданный формат ответа: {answer[:100]}...")
            else:
                print(f"   ❌ Ошибка Ollama API: {response.status_code}")
                
        except requests.exceptions.Timeout:
            print(f"   ❌ Таймаут запроса к LLM")
        except Exception as e:
            print(f"   ❌ Ошибка при запросе к LLM: {e}")
    
    # Удаляем дубликаты по названию работы
    seen = set()
    unique_works = []
    for work in matched_works:
        key = work['work_name']
        if key not in seen:
            seen.add(key)
            unique_works.append(work)
    
    return unique_works


def load_and_prepare_works(works_file: str) -> pd.DataFrame:
    """Загружает таблицу работ без фильтрации - нужна полная структура"""
    print(f"Загрузка таблицы работ из {works_file}...")
    
    try:
        df_works = pd.read_excel(works_file, sheet_name='ВОР КР+расценки')
    except:
        try:
            df_works = pd.read_excel(works_file, sheet_name='Перечень работ КР_new')
        except:
            df_works = pd.read_excel(works_file)
    
    print(f"Виды работ: {len(df_works)} строк")
    
    # Не фильтруем - нужны все строки включая заголовки разделов и подразделов
    return df_works


def check_ollama_availability() -> bool:
    """Проверяет доступность Ollama сервера"""
    try:
        response = requests.get(f"{OLLAMA_BASE_URL}/api/tags", timeout=5)
        if response.status_code == 200:
            print("✅ Ollama сервер доступен")
            return True
        else:
            print(f"⚠️  Ollama вернул статус {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ Не удалось подключиться к Ollama: {e}")
        return False


def map_elements_to_works_llm(session_folder=None):
    """Основная функция маппинга элементов к работам с использованием LLM"""
    try:
        result = main(session_folder)
        return result
    except Exception as e:
        print(f"Ошибка при маппинге элементов к работам: {e}")
        return {'success': False, 'error': str(e)}


def main(session_folder=None):
    # Пути к файлам
    if session_folder is None and len(sys.argv) >= 4:
        elements_file = sys.argv[1]
        works_file = sys.argv[2]
        output_file = sys.argv[3]
    elif session_folder:
        elements_file = os.path.join(session_folder, 'full_elements.xlsx')
        works_file = os.path.join(session_folder, 'Перечень работ КР_new.xlsx')
        output_file = os.path.join(session_folder, 'mapped_elements_works_llm.xlsx')
    else:
        elements_file = 'uploads/b43c911d-191d-42e5-a953-810feb9bf2de/full_elements.xlsx'
        works_file = 'uploads/b43c911d-191d-42e5-a953-810feb9bf2de/Перечень работ КР_new.xlsx'
        output_file = 'uploads/b43c911d-191d-42e5-a953-810feb9bf2de/mapped_elements_works_llm.xlsx'
    
    # Проверка файлов
    if not os.path.exists(elements_file):
        error_msg = f"Ошибка: файл элементов не найден: {elements_file}"
        print(error_msg)
        return {'success': False, 'error': error_msg}
    
    if not os.path.exists(works_file):
        error_msg = f"Ошибка: файл работ не найден: {works_file}"
        print(error_msg)
        return {'success': False, 'error': error_msg}
    
    print("=" * 60)
    print("ПОДБОР ВИДОВ РАБОТ К ЭЛЕМЕНТАМ С ИСПОЛЬЗОВАНИЕМ LLM")
    print("=" * 60)
    print(f"Модель: {LLM_MODEL}")
    print(f"Ollama URL: {OLLAMA_BASE_URL}")
    print("=" * 60)
    print("\nАЛГОРИТМ РАБОТЫ:")
    print("1. Определение уровня элемента (Подземный/Надземный)")
    print("2. Выбор раздела работ по уровню")
    print("3. Выбор подраздела работ по классу элемента")
    print("4. Фильтрация работ по параметрам (LLM)")
    print("5. Фильтрация работ по материалам (LLM)")
    print("=" * 60)
    
    # Проверяем доступность Ollama
    ollama_available = check_ollama_availability()
    if not ollama_available:
        print("⚠️  LLM недоступен. Работа скрипта невозможна.")
        return {'success': False, 'error': 'Ollama сервер недоступен'}
    
    # Загружаем элементы
    print(f"\nЗагрузка таблицы элементов из {elements_file}...")
    df_elements = pd.read_excel(elements_file)
    print(f"Элементы: {len(df_elements)} строк")
    
    # Загружаем работы
    df_works_valid = load_and_prepare_works(works_file)
    
    # Фильтрация элементов по целевым классам
    print("\nФильтрация элементов по целевым классам...")
    print(f"Целевые классы: {', '.join(TARGET_IFC_CLASSES)}")
    valid_element_indices = []
    for idx in range(len(df_elements)):
        element_row = df_elements.iloc[idx]
        params = get_element_parameters(element_row)
        if params['is_target'] and params['ifc_class']:
            valid_element_indices.append(idx)
    
    print(f"Элементов для обработки: {len(valid_element_indices)}")
    
    # Обрабатываем элементы
    results = []
    elements_with_works = set()
    elements_without_works = set()
    
    print("\nПодбор видов работ для элементов с помощью LLM...")
    total = len(valid_element_indices)
    
    for batch_idx, idx in enumerate(valid_element_indices):
        if batch_idx % 50 == 0:
            print(f"Обработано {batch_idx} из {total} элементов...")
        
        element_row = df_elements.iloc[idx]
        element_params = get_element_parameters(element_row)
        
        # Предварительный отбор кандидатов по структуре документа
        candidate_works = select_candidate_works_with_structure(element_params, df_works_valid)
        
        if candidate_works.empty:
            elements_without_works.add(idx)
            continue
        
        # LLM подбирает работы из кандидатов
        print(f"  [{batch_idx}/{total}] {element_params['name'][:60]}...")
        matched_works = match_works_with_llm(element_params, candidate_works)
        
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
            elements_without_works.add(idx)
    
    # Добавляем исключенные элементы
    for idx in range(len(df_elements)):
        if idx not in valid_element_indices:
            elements_without_works.add(idx)
    
    # Создаем DataFrame с результатами
    df_results = pd.DataFrame(results)
    
    # Удаляем дубликаты
    if len(df_results) > 0:
        initial_count = len(df_results)
        df_results = df_results.drop_duplicates(subset=['Element_Index', 'Work_Name'], keep='first')
        duplicates_removed = initial_count - len(df_results)
        if duplicates_removed > 0:
            print(f"Удалено дубликатов работ: {duplicates_removed}")
    
    print(f"\n{'=' * 60}")
    print(f"Результаты подбора видов работ (LLM)")
    print(f"{'=' * 60}")
    print(f"Найдено {len(df_results)} соответствий элемент-работа")
    print(f"Уникальных элементов с работами: {len(elements_with_works)}")
    print(f"Элементов без работ (исключены): {len(elements_without_works)}")
    
    # Сохраняем результат
    df_results.to_excel(output_file, index=False)
    print(f"\nРезультаты сохранены в файл: {output_file}")
    
    # Статистика
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
    if len(sys.argv) >= 2:
        session_folder = sys.argv[1]
        result = main(session_folder)
        if not result.get('success'):
            sys.exit(1)
    else:
        print("Использование: python map_elements_to_works_llm.py <session_folder>")
        print("Пример: python map_elements_to_works_llm.py /workspace/uploads/session_id")
        result = main()
        if result and not result.get('success'):
            sys.exit(1)
