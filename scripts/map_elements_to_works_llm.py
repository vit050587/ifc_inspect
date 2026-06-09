#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для подбора видов работ к элементам из таблицы full_elements.xlsx
на основе таблицы Перечень работ КР_new.xlsx с использованием LLM (Ollama + Gemma2:27b)

Алгоритм с применением LLM:
1. Загружаются элементы и виды работ из Excel файлов
2. Для каждого элемента формируется контекст (параметры, материал, характеристики)
3. LLM анализирует элемент и подбирает подходящие работы из справочника
4. Результат сохраняется в mapped_elements_works.xlsx

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
    
    # Элементы без работ
    if 'Термовкладыш' in name or 'термовкладыш' in name:
        return 'IfcThermalInsert'
    
    if 'Отверст' in name or 'отверст' in name or 'Проем' in name:
        return 'IfcOpeningElement'
    
    if 'Труб' in name and ('Гильз' in name or 'гильз' in name):
        return 'IfcBuildingElementProxy'
    
    # Сопоставление по ключевым словам
    if 'Колонн' in name or 'НесКол' in name:
        return 'IfcColumn'
    elif 'Балк' in name or 'Ригел' in name:
        return 'IfcBeam'
    elif 'Стен' in name or 'Диафрагм' in name:
        return 'IfcWall'
    elif 'Плит' in name or 'Перекрыт' in name or 'Покрыт' in name:
        if 'Фундамент' in name or 'ФП' in name:
            return 'IfcSlabFoundation'
        return 'IfcSlab'
    elif 'Лестничн' in name and 'Марш' in name:
        return 'IfcStairFlight'
    elif 'Фундамент' in name and 'плит' not in name.lower():
        return 'IfcFooting'
    elif 'Ростверк' in name:
        return 'IfcFooting'
    elif 'Арматур' in name:
        return 'IfcReinforcingBar'
    elif 'Закладн' in name or 'Пластина' in name:
        return 'IfcPlate'
    elif 'Гидрошпонк' in name:
        return 'IfcDiscreteAccessory'
    elif 'Приям' in name:
        return 'IfcSlab'
    elif 'Рен' in name or 'стяжк' in name.lower() or 'Подготовка' in name:
        return 'IfcCovering'
    elif 'Пандус' in name:
        return 'IfcRamp'
    
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


def is_element_without_works(ifc_class: str, name: str) -> bool:
    """Проверяет, является ли элемент элементом без работ"""
    no_works_classes = ['IfcOpeningElement', 'IfcBuildingElementProxy', 'IfcThermalInsert']
    
    if ifc_class in no_works_classes:
        return True
    
    if 'термовкладыш' in name.lower():
        return True
    
    return False


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
        'no_works': is_element_without_works(ifc_class, name)
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


def select_candidate_works(element_params: Dict, works_df: pd.DataFrame, max_candidates: int = 50) -> pd.DataFrame:
    """
    Предварительно отбирает кандидатуры работ на основе базовых фильтров
    перед отправкой к LLM
    """
    if element_params['no_works']:
        return pd.DataFrame()
    
    element_ifc = element_params['ifc_class'].lower()
    element_level = element_params['level']
    
    candidates = []
    
    for idx, work_row in works_df.iterrows():
        work_ifc = str(work_row.get('IFC класс', ''))
        
        # Базовая фильтрация по IFC классу
        if pd.notna(work_ifc) and work_ifc not in ['не моделируется', 'чаще всего не моделируется', '-', '']:
            work_ifcs = [x.strip().lower() for x in str(work_ifc).split('\n')]
            
            ifc_match = False
            for w_ifc in work_ifcs:
                if w_ifc in ['не моделируется', 'чаще всего не моделируется', '-', '']:
                    continue
                
                if element_ifc == w_ifc:
                    ifc_match = True
                    break
                
                # Особые случаи
                if element_ifc == 'ifcslabfoundation' and w_ifc in ['ifcslab', 'ifcfooting']:
                    ifc_match = True
                    break
            
            if not ifc_match:
                continue
        
        # Фильтрация по уровню (если возможно определить)
        work_name = str(work_row.get('Наименование работ', '')).lower()
        
        if element_level == 'Подземный':
            if 'надземн' in work_name and 'подземн' not in work_name:
                continue
        elif element_level == 'Надземный':
            if 'подземн' in work_name and 'надземн' not in work_name:
                continue
        
        candidates.append(work_row)
    
    if len(candidates) > max_candidates:
        # Если кандидатов слишком много, берем первые max_candidates
        candidates = candidates[:max_candidates]
    
    return pd.DataFrame(candidates)


def match_works_with_llm(element_params: Dict, candidate_works: pd.DataFrame) -> List[Dict]:
    """
    Использует LLM для подбора работ к элементу из списка кандидатов
    
    Args:
        element_params: Параметры элемента
        candidate_works: DataFrame с кандидатами работ
        
    Returns:
        Список подобранных работ
    """
    if candidate_works.empty or element_params['no_works']:
        return []
    
    # Формируем описание элемента
    element_description = format_element_for_prompt(element_params)
    
    # Формируем список работ для анализа (ограничиваем размер)
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
        
        # Формируем промпт для LLM
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

СПИСОК РАБОТ-КАНДИДАТОВ:
{works_text}

ИНСТРУКЦИЯ:
Проанализируй элемент и выбери из списка ТОЛЬКО те работы, которые действительно подходят этому элементу.
Учитывай:
1. Тип элемента (IFC класс) - работа должна соответствовать типу конструкции
2. Уровень расположения (подземный/надземный) - если указано в названии работы
3. Параметры элемента (размеры, объем, площадь) - сверяй с ограничениями в графе "Параметры"
4. Материал элемента (класс бетона, морозостойкость, водонепроницаемость) - сверяй с требованиями в графе "Параметры"
5. Семантику работы - работа должна иметь смысл для данного типа элемента

ВАЖНО:
- Не выбирай работы которые явно не подходят (например, опалубка для немонолитных элементов)
- Учитывай параметрические ограничения (t<XXX, V>XXX, S<XXX и т.д.)
- Если параметров нет в работе - она может подойти любому элементу своего типа
- Исключи дубликаты

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
    
    # Удаляем дубликаты
    seen = set()
    unique_works = []
    for work in matched_works:
        key = work['work_name']
        if key not in seen:
            seen.add(key)
            unique_works.append(work)
    
    return unique_works


def load_and_prepare_works(works_file: str) -> pd.DataFrame:
    """Загружает таблицу работ"""
    print(f"Загрузка таблицы работ из {works_file}...")
    
    try:
        df_works = pd.read_excel(works_file, sheet_name='ВОР КР+расценки')
    except:
        try:
            df_works = pd.read_excel(works_file, sheet_name='Перечень работ КР_new')
        except:
            df_works = pd.read_excel(works_file)
    
    print(f"Виды работ: {len(df_works)} строк")
    
    # Фильтруем заголовки разделов
    valid_works_mask = df_works['IFC класс'].notna() | df_works['Шифр ТСН'].notna()
    df_works_valid = df_works[valid_works_mask].copy()
    print(f"Валидных видов работ: {len(df_works_valid)}")
    
    return df_works_valid


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
    
    # Предварительная фильтрация элементов без работ
    print("\nПредварительная фильтрация элементов...")
    valid_element_indices = []
    for idx in range(len(df_elements)):
        element_row = df_elements.iloc[idx]
        params = get_element_parameters(element_row)
        if not params['no_works'] and params['ifc_class']:
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
        
        # Предварительный отбор кандидатов
        candidate_works = select_candidate_works(element_params, df_works_valid)
        
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
