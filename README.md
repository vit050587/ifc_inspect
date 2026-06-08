# IFC Inspect

Сервис автоматизированного анализа проектной документации для строительства: обработка PDF (пояснительные записки, чертежи, ведомости), IFC моделей (BIM) и Excel спецификаций с генерацией отчетов.

## 📖 О проекте

IFC Inspect — это веб-сервис, который автоматизирует анализ проектной документации в 5 шагов:

1. **Организация файлов**: автоматическое распределение по папкам (IFC/Excel/PDF)
2. **Классификация PDF через LLM**: определение типа документа (ведомости vs пояснительная записка) через Ollama + gemma3:27b
3. **Извлечение чертежей**: разделение PDF на текстовые страницы (≤30см) и чертежи (>30см) с коррекцией ориентации
4. **Парсинг IFC**: извлечение элементов, материалов, свойств → детальный Excel отчет
5. **Сводная таблица материалов**: агрегация данных по типам и материалам с подсчетом объема и количества

**Веб-интерфейс**: drag-and-drop загрузка, запуск обработки, просмотр результатов, скачивание Excel отчетов.

## 🏗️ Структура проекта

```
/workspace/
├── backend/
│   └── app.py              # Flask API сервер (6 endpoints, порт 6003)
├── frontend/
│   └── index.html          # Single-page приложение (Vanilla JS)
├── scripts/
│   ├── files_classifier.py       # Классификация файлов по типам (IFC/Excel/PDF)
│   ├── pdf_classifier_llm.py     # LLM-классификация PDF (gemma3:27b via Ollama)
│   ├── draw_detector.py          # Извлечение чертежей из PDF (>30см)
│   └── ifc_parser.py             # Парсинг IFC → Excel отчет + materials_summary.xlsx
├── uploads/                # Сессионные данные (session_id/)
├── requirements.txt        # Зависимости Python
├── start-dev.sh           # Скрипт запуска (проверка Ollama + Flask)
└── README.md              # Документация
```

**Не используются** (можно удалить): `scripts/page_splitter.py`, `scripts/pdf_classifier.py`, `scripts/materials_summary.py`, `ifc_inspect.code-workspace`

## ✨ Возможности

### Обработка PDF
- **LLM-классификация**: анализ первых 3 страниц через Ollama API (gemma3:27b, temperature=0.1)
- **Категории**: `volume_statement` (ведомости объемов работ) или `explanatory_note` (пояснительная записка)
- **Fallback**: эвристическая классификация по ключевым словам при ошибке LLM
- **Извлечение чертежей**: страницы >30см → отдельные PDF с коррекцией ориентации (альбомная)

### Парсинг IFC
- **Элементы**: IfcWall, IfcSlab, IfcColumn, IfcBeam, IfcStair, IfcRamp
- **Материалы**: извлечение из Pset_*, ExpCheck_* свойств
- **Объем**: расчет через bounding box
- **Отчеты**: 
  - Детальный Excel по каждому элементу
  - Сводная таблица `materials_summary.xlsx` (агрегация по типу и материалу)

### Веб-интерфейс
- Drag-and-drop загрузка множественных файлов
- Уникальный session_id для каждой сессии
- Просмотр результатов: количество страниц, чертежей, элементов IFC
- Скачивание: `materials_summary.xlsx`, детальные отчеты, извлеченные PDF

## 🚀 Установка и запуск

### Требования
- Python 3.12+
- Ollama с моделями: `gemma3:27b` (классификация), `gemma4:31b` (опционально для VLM)
- pip

### Установка зависимостей

```bash
pip install -r requirements.txt
```

**Примечание**: если `ifcopenshell==0.8.5` не устанавливается, используйте версию `≥0.7.0`:
```bash
pip install ifcopenshell>=0.7.0
```

### Запуск сервера

```bash
chmod +x start-dev.sh
./start-dev.sh
```

Сервер запустится на порту **6003**: http://localhost:6003

**start-dev.sh** автоматически:
- Устанавливает переменные окружения
- Проверяет доступность Ollama API
- Запускает Flask в режиме development

## 📡 API Endpoints

| Метод | URL | Описание | Параметры |
|-------|-----|----------|-----------|
| GET | `/` | Веб-интерфейс | — |
| POST | `/api/upload` | Загрузка файлов | `files[]` (multipart/form-data) |
| POST | `/api/process` | Запуск обработки | `{session_id: string}` |
| GET | `/api/results/<session_id>` | Получение результатов | — |
| GET | `/api/download/<session_id>/<filename>` | Скачивание файлов | — |
| GET | `/api/materials-summary/<session_id>` | Данные сводной таблицы (JSON) | — |

### Примеры запросов

**Загрузка файлов:**
```bash
curl -X POST http://localhost:6003/api/upload \
  -F "files[]=project.pdf" \
  -F "files[]=building.ifc"
```

**Ответ:**
```json
{
  "session_id": "550e8400-e29b-41d4-a716-446655440000",
  "pdf_count": 1,
  "ifc_count": 1,
  "message": "Uploaded 1 PDF(s) and 1 IFC file(s)"
}
```

**Запуск обработки:**
```bash
curl -X POST http://localhost:6003/api/process \
  -H "Content-Type: application/json" \
  -d '{"session_id": "550e8400-e29b-41d4-a716-446655440000"}'
```

**Получение результатов:**
```bash
curl http://localhost:6003/api/results/550e8400-e29b-41d4-a716-446655440000
```

## 🔄 Алгоритм работы

1. **Загрузка**: Пользователь загружает файлы → создание `uploads/<session_id>/`
2. **Организация** (`files_classifier.py`):
   - IFC → `ifc_models/`
   - Excel → `specification/`
   - PDF → `pdf_documents/`

3. **Классификация PDF** (`pdf_classifier_llm.py`):
   - Подключение к Ollama API (http://localhost:11434/api/generate)
   - Анализ первых 3 страниц каждого PDF
   - Промпт: "Определи тип документа: ведомость объемов работ или пояснительная записка"
   - Результат → `classification_results.json`

4. **Извлечение чертежей** (`draw_detector.py`):
   - Обработка PDF из `explanatory_note/`
   - PyMuPDF: измерение размера каждой страницы
   - Страницы >30см → `drawings/` (поворот в альбомную ориентацию)
   - Страницы ≤30см → `explanatory_note/` (текстовые)
   - Результат → `drawings_results.json`

5. **Парсинг IFC** (`ifc_parser.py`):
   - ifcopenshell: чтение `.ifc` файла
   - Фильтрация элементов: Wall, Slab, Column, Beam, Stair, Ramp
   - Извлечение свойств: Material, Volume, Area, Dimensions
   - Создание `materials_summary.xlsx` (агрегация: Тип, Материал, Кол-во, Объем)

6. **Сохранение результатов**:
   - `results.json` — детали обработки
   - `session_info.json` — статус сессии
   - `materials_summary.xlsx` + `materials_summary.json` — сводная таблица
   - Папки: `drawings/`, `explanatory_note/`, `volume_statement/`

7. **Отображение**: Фронтенд показывает статистику и кнопки скачивания

## 📊 Формат данных

### materials_summary.json
```json
{
  "success": true,
  "source_file": "ifc модель КР.ifc",
  "total_items": 14,
  "items": [
    {
      "Тип (RU)": "Стены",
      "Тип элемента": "IfcWall",
      "Материал": "Бетон В30 F150 W6",
      "Количество, шт": 984,
      "Объем, м³": 1731.541
    }
  ]
}
```

### results.json
```json
{
  "text_pages": [
    {"page_num": 1, "source_file": "pz.pdf", "size": "21.0x29.7cm", "type": "text"}
  ],
  "drawing_pages": [
    {"page_num": 5, "source_file": "pz.pdf", "size": "59.4x84.1cm", "type": "drawing"}
  ],
  "pdf_classification": {
    "pz.pdf": {"category": "explanatory_note", "confidence": 0.95}
  },
  "ifc_results": {
    "success": true,
    "excel_filename": "materials_summary.xlsx",
    "total_elements": 3209,
    "materials_count": 47
  }
}
```

## 🛠️ Технологии

| Компонент | Технология | Версия |
|-----------|------------|--------|
| Backend | Flask | 3.0.0 |
| CORS | flask-cors | 4.0.0 |
| PDF | PyMuPDF (fitz) | 1.23.8 |
| IFC | ifcopenshell | ≥0.7.0 |
| Excel | openpyxl, pandas | ≥3.1.5, ≥2.1.4 |
| LLM | Ollama API | gemma3:27b |
| Frontend | HTML5, CSS3, Vanilla JS | — |

## ⚙️ Конфигурация

Переменные окружения задаются в `start-dev.sh`:

| Параметр | Значение | Описание |
|----------|----------|----------|
| `FLASK_APP` | backend/app.py | Приложение Flask |
| `FLASK_ENV` | development | Режим среды |
| `UPLOAD_FOLDER` | ./uploads | Папка загрузок |
| `OLLAMA_BASE_URL` | http://localhost:11434 | URL Ollama API |
| `DRAWING_VALIDATION_MODEL` | gemma3:27b | Модель для классификации PDF |
| `DRAWING_MIN_SIZE_CM` | 42.0 | Порог размера чертежа (см) |

**В скриптах:**
- `pdf_classifier_llm.py`: `temperature=0.1`, `top_p=0.9`, `num_predict=50`
- `draw_detector.py`: `DRAWING_MIN_SIZE_CM = 30.0` (локальная константа)

## 🔐 Безопасность

- Нет аутентификации (предназначен для локального использования)
- Сессионные данные хранятся в JSON файлах в `uploads/<session_id>/`
- Для продакшена рекомендуется nginx + HTTPS

## 📝 Лицензия

MIT
