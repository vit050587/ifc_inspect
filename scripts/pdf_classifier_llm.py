#!/usr/bin/env python3
"""
Скрипт классификации PDF документов с помощью LLM (Ollama + gemma3:27b).
Анализирует первые 3 страницы каждого PDF файла и определяет тип документа:
- Ведомость объема работ и материалов -> volume_statement/
- Текстовая и графическая часть (пояснительная записка) -> explanatory_note/

Использует Ollama API для подключения к локальной модели.
"""

import os
import json
import requests
from pathlib import Path
from typing import Dict, List, Any, Optional
import fitz  # PyMuPDF


# Настройки из переменных окружения
OLLAMA_BASE_URL = os.getenv("OLLAMA_BASE_URL", "http://localhost:11434")
LLM_MODEL = os.getenv("DRAWING_VALIDATION_MODEL", "gemma3:27b")

# Категории документов (обновленные названия папок)
DOCUMENT_CATEGORIES = {
    "volume_statement": {
        "folder": "volume_statement",
        "description": "Ведомость объемов работ и материалов"
    },
    "explanatory_note": {
        "folder": "explanatory_note",
        "description": "Текстовая и графическая часть (пояснительная записка)"
    }
}


def extract_text_from_pdf_pages(pdf_path: str, max_pages: int = 3) -> str:
    """
    Извлекает текст из первых max_pages страниц PDF файла.
    
    Args:
        pdf_path: Путь к PDF файлу
        max_pages: Максимальное количество страниц для анализа
        
    Returns:
        Извлеченный текст
    """
    try:
        doc = fitz.open(pdf_path)
        text_parts = []
        
        pages_to_read = min(max_pages, len(doc))
        print(f"   📖 Чтение первых {pages_to_read} страниц из {len(doc)}...")
        
        for page_num in range(pages_to_read):
            page = doc[page_num]
            page_text = page.get_text()
            
            # Добавляем маркер номера страницы
            text_parts.append(f"[Страница {page_num + 1}]")
            text_parts.append(page_text[:3000])  # Ограничиваем текст с каждой страницы
            
        doc.close()
        
        full_text = "\n\n".join(text_parts)
        return full_text
    
    except Exception as e:
        print(f"   ❌ Ошибка при чтении PDF: {e}")
        return ""


def classify_document_with_llm(text: str, filename: str) -> Optional[str]:
    """
    Классифицирует документ с помощью LLM через Ollama API.
    
    Args:
        text: Текст из документа
        filename: Имя файла для контекста
        
    Returns:
        Категория документа или None если не удалось классифицировать
    """
    
    prompt = f"""
Ты эксперт по строительной документации. Проанализируй текст из документа и определи его тип.

Имя файла: {filename}

Текст документа (первые 3 страницы):
{text[:8000]}  # Ограничиваем общий размер текста

Категории для классификации:
1. "volume_statement" - ведомость объемов работ, таблицы с количествами материалов, единицами измерения (м3, м2, шт.), спецификация элементов
2. "explanatory_note" - текстовое описание проекта, пояснительная записка, архитектурные и конструктивные решения, характеристики здания, описание проекта

ВАЖНО: Ответь ТОЛЬКО одним словом из категорий выше (volume_statement или explanatory_note). Без пояснений.

Категория:"""

    try:
        response = requests.post(
            f"{OLLAMA_BASE_URL}/api/generate",
            json={
                "model": LLM_MODEL,
                "prompt": prompt,
                "stream": False,
                "options": {
                    "temperature": 0.1,  # Низкая температура для детерминированного ответа
                    "top_p": 0.9,
                    "num_predict": 50  # Достаточно для одного слова
                }
            },
            timeout=120  # 2 минуты на генерацию
        )
        
        if response.status_code == 200:
            result = response.json()
            category = result.get("response", "").strip().lower()
            
            # Очищаем ответ от лишних символов
            category = category.replace('"', '').replace('.', '').replace('\n', '').strip()
            
            # Проверяем что категория валидна
            if category in DOCUMENT_CATEGORIES:
                return category
            else:
                # Пытаемся найти похожую категорию
                for cat_key in DOCUMENT_CATEGORIES.keys():
                    if cat_key in category or category in cat_key:
                        return cat_key
                
                print(f"   ⚠️  Неизвестная категория от LLM: '{category}', используем 'explanatory_note' по умолчанию")
                return "explanatory_note"
        else:
            print(f"   ❌ Ошибка Ollama API: {response.status_code} - {response.text}")
            return None
            
    except requests.exceptions.Timeout:
        print(f"   ❌ Таймаут запроса к Ollama (>120 сек)")
        return None
    except requests.exceptions.ConnectionError:
        print(f"   ❌ Не удалось подключиться к Ollama по адресу {OLLAMA_BASE_URL}")
        return None
    except Exception as e:
        print(f"   ❌ Ошибка при классификации: {e}")
        return None


def heuristic_fallback_classification(filename: str, text: str) -> str:
    """
    Эвристическая классификация если LLM недоступна.
    Использует ключевые слова в тексте и имени файла.
    """
    filename_lower = filename.lower()
    text_lower = text.lower()
    
    # Проверка по имени файла
    if "ведомост" in filename_lower or "volum" in filename_lower or "volume" in filename_lower:
        return "volume_statement"
    
    if "спецификац" in filename_lower or "specif" in filename_lower:
        return "volume_statement"
    
    if "пояснит" in filename_lower or "explanatory" in filename_lower or "description" in filename_lower:
        return "explanatory_note"
    
    if "записк" in filename_lower or "note" in filename_lower:
        return "explanatory_note"
    
    # Проверка по содержанию текста
    text_sample = text_lower[:2000]
    
    # Ключевые слова для ведомостей
    volume_keywords = ["ведомость объемов", "объем работ", "единица измерения", "кол-во", 
                       "количество", "м3", "м2", "пог.м", "шт.", "таблица", "спецификация"]
    if sum(1 for kw in volume_keywords if kw in text_sample) >= 3:
        return "volume_statement"
    
    # Ключевые слова для пояснительной записки
    explanatory_keywords = ["пояснительная записка", "описание проекта", "характеристики здания",
                            "архитектурные решения", "конструктивные решения", "раздел"]
    if sum(1 for kw in explanatory_keywords if kw in text_sample) >= 2:
        return "explanatory_note"
    
    # По умолчанию - пояснительная записка
    return "explanatory_note"


def classify_pdf_files(session_folder: str) -> Dict[str, List[Dict[str, Any]]]:
    """
    Классифицирует все PDF файлы в папке pdf_documents и распределяет по папкам.
    
    Args:
        session_folder: Путь к папке сессии
        
    Returns:
        Словарь с результатами классификации по категориям
    """
    
    results = {cat: [] for cat in DOCUMENT_CATEGORIES.keys()}
    
    # Создаем папки для категорий
    for category, info in DOCUMENT_CATEGORIES.items():
        category_folder = Path(session_folder) / info["folder"]
        category_folder.mkdir(parents=True, exist_ok=True)
        print(f"📁 Создана папка: {info['folder']}")
    
    # Находим все PDF файлы в папке pdf_documents
    pdf_folder = Path(session_folder) / 'pdf_documents'
    if not pdf_folder.exists():
        print(f"⚠️  Папка pdf_documents не найдена: {pdf_folder}")
        return results
    
    pdf_files = [f for f in pdf_folder.glob("*.pdf") if f.is_file()]
    
    print(f"\n🔍 Найдено PDF файлов для классификации: {len(pdf_files)}")
    
    for pdf_path in pdf_files:
        filename = pdf_path.name
        print(f"\n📄 Обработка файла: {filename}")
        
        # Извлекаем текст из первых 3 страниц
        text = extract_text_from_pdf_pages(str(pdf_path), max_pages=3)
        
        if not text:
            print(f"   ⚠️  Не удалось извлечь текст, используем эвристику")
            category = heuristic_fallback_classification(filename, "")
        else:
            # Пробуем классифицировать через LLM
            print(f"   🤖 Запрос к LLM ({LLM_MODEL})...")
            category = classify_document_with_llm(text, filename)
            
            # Если LLM не ответил, используем эвристику
            if category is None:
                print(f"   ⚠️  LLM не ответил, используем эвристическую классификацию")
                category = heuristic_fallback_classification(filename, text)
        
        print(f"   ✅ Категория: {category}")
        
        # Перемещаем файл в соответствующую папку
        category_info = DOCUMENT_CATEGORIES[category]
        dest_folder = Path(session_folder) / category_info["folder"]
        dest_path = dest_folder / filename
        
        try:
            # Копируем файл вместо перемещения (чтобы сохранить оригиналы)
            import shutil
            shutil.copy2(str(pdf_path), str(dest_path))
            print(f"   📋 Скопирован в: {category_info['folder']}/{filename}")
        except Exception as e:
            print(f"   ❌ Ошибка копирования: {e}")
            dest_path = pdf_path
        
        # Добавляем в результаты
        file_info = {
            "filename": filename,
            "source_path": str(pdf_path),
            "dest_path": str(dest_path),
            "category": category,
            "pages_analyzed": min(3, len(fitz.open(str(pdf_path)))),
            "file_size": pdf_path.stat().st_size
        }
        
        results[category].append(file_info)
    
    # Сохраняем результаты классификации
    classification_results_path = Path(session_folder) / "classification_results.json"
    with open(classification_results_path, 'w', encoding='utf-8') as f:
        json.dump({
            "categories": {cat: info["description"] for cat, info in DOCUMENT_CATEGORIES.items()},
            "results": results,
            "summary": {cat: len(files) for cat, files in results.items()}
        }, f, ensure_ascii=False, indent=2)
    
    print(f"\n💾 Результаты сохранены в: {classification_results_path}")
    
    # Выводим сводку
    print("\n" + "="*60)
    print("📊 СВОДКА ПО КАТЕГОРИЯМ:")
    print("="*60)
    for category, files in results.items():
        emoji = {"volume_statement": "📊", "explanatory_note": "📘"}.get(category, "📁")
        print(f"{emoji} {category}: {len(files)} файл(ов)")
        for file_info in files:
            print(f"   • {file_info['filename']} ({file_info['file_size'] / 1024 / 1024:.1f} MB)")
    
    return results


def main(session_folder: str):
    """
    Основная функция классификации PDF файлов.
    
    Args:
        session_folder: Путь к папке сессии
    """
    print("="*60)
    print("🔍 КЛАССИФИКАЦИЯ PDF ДОКУМЕНТОВ С ПОМОЩЬЮ LLM")
    print("="*60)
    print(f"Папка сессии: {session_folder}")
    print(f"Ollama URL: {OLLAMA_BASE_URL}")
    print(f"Модель: {LLM_MODEL}")
    print("="*60)
    
    # Проверяем доступность Ollama
    try:
        response = requests.get(f"{OLLAMA_BASE_URL}/api/tags", timeout=5)
        if response.status_code == 200:
            print("✅ Ollama сервер доступен")
        else:
            print(f"⚠️  Ollama вернул статус {response.status_code}")
    except Exception as e:
        print(f"❌ Не удалось подключиться к Ollama: {e}")
        print("Будет использована эвристическая классификация")
    
    # Запускаем классификацию
    results = classify_pdf_files(session_folder)
    
    return results


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Использование: python pdf_classifier_llm.py <session_folder>")
        print("Пример: python pdf_classifier_llm.py /path/to/uploads/session_id")
        sys.exit(1)
    
    session_folder = sys.argv[1]
    main(session_folder)
