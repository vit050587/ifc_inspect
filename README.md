# IFC Inspect

Сервис анализа проектной документации для строительства: обработка PDF файлов (проектная документация), IFC моделей (информационное моделирование зданий) и Excel спецификаций с генерацией отчетов и 3D визуализацией.

## 📖 О проекте

IFC Inspect — это веб-сервис, который автоматизирует анализ проектной документации:
- **PDF файлы**: классификация через LLM (Ollama + gemma3:27b), извлечение чертежей из пояснительной записки
- **IFC модели**: парсинг элементов, материалов, свойств → детальный Excel отчет + 3D вьюер в браузере
- **Excel спецификации**: парсинг ведомостей материалов → сводная таблица по объемам/площадям/количеству
- **Веб-интерфейс**: загрузка файлов, запуск анализа, просмотр результатов, скачивание отчетов, 3D визуализация IFC

## 🏗️ Структура проекта

```
/workspace/
├── backend/
│   └── app.py              # Flask API сервер (6 endpoints)
├── frontend/
│   └── index.html          # Single-page приложение (Vanilla JS + Three.js для IFC)
├── scripts/
│   ├── files_classifier.py       # Классификация файлов по типам (IFC/Excel/PDF)
│   ├── pdf_classifier_llm.py     # Классификация PDF через LLM (gemma3:27b)
│   ├── draw_detector.py          # Извлечение чертежей из PDF (>30см)
│   ├── ifc_viewer.py             # Подготовка IFC для веб-вьюера (геометрия → JSON)
│   ├── ifc_parser.py             # Парсинг IFC → Excel отчет
│   ├── xlsx_parser.py            # Парсинг Excel спецификаций → сводная таблица
│   │
│   ├── page_splitter.py          # ⚠️ НЕ ИСПОЛЬЗУЕТСЯ (заменен на draw_detector.py)
│   ├── pdf_classifier.py         # ⚠️ НЕ ИСПОЛЬЗУЕТСЯ (заменен на pdf_classifier_llm.py)
│   └── materials_summary.py      # ⚠️ НЕ ИСПОЛЬЗУЕТСЯ (логика в xlsx_parser.py)
├── uploads/                # Сессионные данные (создается автоматически)
├── requirements.txt        # Зависимости Python
├── start-dev.sh           # Скрипт запуска (проверка Ollama + Flask порт 6003)
├── tunnel.sh              # SSH туннель для удаленного доступа
├── ifc_inspect.code-workspace  # ⚠️ НЕ ИСПОЛЬЗУЕТСЯ (настройки VS Code)
└── README.md              # Документация
```

## ✨ Возможности

### 1. Загрузка файлов
- Drag-and-drop интерфейс для загрузки PDF и IFC файлов
- Поддержка множественной загрузки
- Автоматическая категоризация по типу файла
- Уникальные session_id для каждой сессии

### 2. Обработка PDF
- **Классификация страниц** по размеру:
  - Текстовые страницы: ≤ 30 см (формат A4 и меньше)
  - Чертежи: > 30 см (большие форматы)
- **Коррекция ориентации**: автоматический поворот страниц
- **Сохранение результатов**: отдельные PDF файлы для каждой страницы

### 3. Парсинг IFC
- **Информация о проекте**: название, дата создания, схема IFC (IFC2X3/IFC4)
- **Подсчет элементов по типам**:
  - Стены (IfcWall)
  - Перекрытия (IfcSlab)
  - Колонны (IfcColumn)
  - Балки (IfcBeam)
  - Лестницы (IfcStair)
  - Ограждения (IfcRailing)
  - Крыши (IfcRoof)
  - Проемы (IfcOpeningElement)
  - Окна (IfcWindow)
  - Двери (IfcDoor)
  - Прочие элементы (IfcBuildingElementProxy)
- **Материалы**: извлечение всех материалов и их свойств
- **Этажи**: подсчет количества уровней здания
- **Excel отчет**: детальная таблица с параметрами каждого элемента

## 🚀 Установка и запуск

### Требования
- Python 3.8+
- pip

### Установка зависимостей

```bash
pip install -r requirements.txt
```

**Примечание**: если `ifcopenshell==0.8.5` не устанавливается, используйте версию `0.7.0` или новее:
```bash
pip install ifcopenshell>=0.7.0
```

### Запуск сервера разработки

```bash
cd backend
python app.py
```

Сервер запустится на порту **6003**: http://localhost:6003

### Запуск в продакшене

Используйте скрипт `start-dev.sh`:

```bash
chmod +x start-dev.sh
./start-dev.sh
```

Сервер запустится на порту **6003**.

### Удаленный доступ через SSH туннель

Если сервер находится на удаленной машине:

```bash
chmod +x tunnel.sh
./tunnel.sh <remote_host> <remote_port>
```

Затем откройте http://localhost:8082 в браузере.

## 📡 API Endpoints

| Метод | URL | Описание | Параметры |
|-------|-----|----------|-----------|
| GET | `/` | Веб-интерфейс | — |
| POST | `/api/upload` | Загрузка файлов | `files` (multipart/form-data) |
| POST | `/api/process` | Обработка файлов | `{session_id: string}` |
| GET | `/api/results/<session_id>` | Получение результатов | — |
| GET | `/api/viewer/<session_id>/<filename>` | Файлы для IFC вьюера | — |
| GET | `/api/download/<session_id>/<filename>` | Скачивание файлов | — |

### Примеры запросов

#### Загрузка файлов

```bash
curl -X POST http://localhost:5001/api/upload \
  -F "files=@project.pdf" \
  -F "files=@building.ifc"
```

Ответ:
```json
{
  "session_id": "uuid-string",
  "pdf_count": 1,
  "ifc_count": 1,
  "other_count": 0,
  "message": "Uploaded 1 PDF(s) and 1 IFC file(s)"
}
```

#### Обработка файлов

```bash
curl -X POST http://localhost:6003/api/process \
  -H "Content-Type: application/json" \
  -d '{"session_id": "uuid-string"}'
```

#### Получение результатов

```bash
curl http://localhost:6003/api/results/uuid-string
```

## 🔄 Алгоритм работы

1. **Загрузка**: Пользователь загружает PDF, IFC и Excel файлы через веб-интерфейс
2. **Создание сессии**: Сервер генерирует unique session_id и сохраняет файлы в `uploads/<session_id>/`
3. **Обработка** (по кнопке "Запустить анализ"):
   - **ШАГ 1: Организация файлов** (`files_classifier.py`):
     - Распределение по папкам: `ifc_models/`, `specification/`, `pdf_documents/`
   
   - **ШАГ 2: Классификация PDF** (`pdf_classifier_llm.py`):
     - Подключение к Ollama API (gemma3:27b)
     - Анализ первых 3 страниц каждого PDF
     - Категоризация: `volume_statement` (ведомости) или `explanatory_note` (пояснительная записка)
   
   - **ШАГ 3: Извлечение чертежей** (`draw_detector.py`):
     - Обработка PDF из папки `explanatory_note/`
     - Определение размера каждой страницы через PyMuPDF
     - Страницы > 30см → чертежи (сохраняются в `drawings/`)
     - Страницы ≤ 30см → текст (остаются в `text_pages/`)
     - Коррекция ориентации (альбомная для чертежей)
   
   - **ШАГ 4: Обработка IFC** (`ifc_viewer.py` + `ifc_parser.py`):
     - **4a. Подготовка вьюера**: извлечение геометрии через ifcopenshell → JSON для Three.js
     - **4b. Парсинг**: элементы, материалы, свойства, этажи → Excel отчет
   
   - **ШАГ 5: Парсинг Excel спецификации** (`xlsx_parser.py`):
     - Чтение файлов из `specification/`
     - Извлечение: марка материала, объем, площадь, количество
     - Агрегация данных по материалам
     - Создание сводной таблицы `materials_summary.xlsx`

4. **Сохранение результатов**:
   - `results.json` — детали обработки (страницы, IFC результаты, классификация)
   - `session_info.json` — статус сессии и сводка
   - `viewer/` — файлы для 3D вьюера (IFC + геометрия JSON)
   - `ifc_report.xlsx` — детальный отчет по элементам IFC
   - `materials_summary.xlsx` — сводная таблица материалов

5. **Отображение**: Фронтенд получает результаты и показывает:
   - Количество текстовых страниц и чертежей
   - Статистику по элементам IFC
   - 3D вьюер модели (Three.js + BIMSurfer)
   - Кнопки для скачивания Excel отчетов

## 📊 Формат данных

### session_info.json

```json
{
  "session_id": "uuid",
  "status": "completed",
  "pdf_files": ["project.pdf"],
  "ifc_files": ["building.ifc"],
  "other_files": [],
  "ifc_excel_file": "ifc_report.xlsx",
  "results_summary": {
    "text_pages_count": 15,
    "drawing_pages_count": 42,
    "ifc_processed": true
  }
}
```

### results.json

```json
{
  "text_pages": [
    {"page_num": 1, "source_file": "project.pdf", "size": "21.0x29.7cm", "type": "text"}
  ],
  "drawing_pages": [
    {"page_num": 2, "source_file": "project.pdf", "size": "59.4x84.1cm", "type": "drawing"}
  ],
  "ifc_results": {
    "success": true,
    "excel_filename": "ifc_report.xlsx",
    "project_info": {
      "name": "Жилой комплекс",
      "creation_date": "2024-01-15",
      "schema": "IFC4"
    },
    "elements_count": {
      "IfcWall": 1221,
      "IfcSlab": 125,
      "IfcColumn": 89
    },
    "total_elements": 3209,
    "materials_count": 47,
    "stories_count": 23
  }
}
```

### Excel отчет (ifc_report.xlsx)

**Лист "Сводка"**:
- Название проекта
- Дата создания
- Схема IFC
- Общее количество элементов
- Количество материалов
- Количество этажей

**Листы по типам элементов** (Стены, Перекрытия, Колонны, и т.д.):
- GlobalId (уникальный ID)
- Name (имя элемента)
- Tag (марка)
- Description (описание)
- Material (материал)
- Volume (объем, м³)
- Area (площадь, м²)
- Dimensions (размеры)
- Properties (дополнительные свойства)

## 🛠️ Технологии

| Компонент | Технология |
|-----------|------------|
| Backend | Flask 3.0.0 |
| CORS | flask-cors 4.0.0 |
| PDF обработка | PyMuPDF 1.23.8 (fitz) |
| IFC парсинг | ifcopenshell >=0.7.0 |
| Excel генерация | openpyxl ≥3.1.5, pandas 2.1.4 |
| LLM классификация | Ollama API (gemma3:27b, gemma4:31b) |
| Frontend | HTML5, CSS3, Vanilla JavaScript |
| 3D визуализация | Three.js + BIMSurfer |

## ⚙️ Конфигурация

Настройки задаются через переменные окружения в `start-dev.sh`:

| Параметр | Значение по умолчанию | Описание |
|----------|----------|----------|
| `FLASK_APP` | backend/app.py | Приложение Flask |
| `FLASK_ENV` | development | Режим среды |
| `UPLOAD_FOLDER` | ./uploads | Папка загрузок |
| `OLLAMA_BASE_URL` | http://localhost:11434 | URL Ollama API |
| `DRAWING_VLM_MODEL` | gemma4:31b | Модель для анализа чертежей (VLM) |
| `DRAWING_VALIDATION_MODEL` | gemma3:27b | Модель для классификации PDF |
| `DRAWING_MIN_SIZE_CM` | 42.0 | Порог размера для чертежа (см) |

## 🔐 Безопасность

- Нет аутентификации/авторизации (локальное использование)
- Сессионные данные хранятся в JSON файлах
- Рекомендуется использовать за обратным прокси (nginx) в продакшене

## 📝 Лицензия

MIT

## 🗑️ Файлы для удаления (не используются)

Следующие файлы **не участвуют** в работе сервиса и могут быть удалены:

| Файл | Причина |
|------|---------|
| `scripts/page_splitter.py` | Заменен на `draw_detector.py` (более полная логика) |
| `scripts/pdf_classifier.py` | Заменен на `pdf_classifier_llm.py` (обновленная версия) |
| `scripts/materials_summary.py` | Логика интегрирована в `xlsx_parser.py` |
| `frontend/index.html.styles` | Дубликат стилей (все стили в index.html) |
| `ifc_inspect.code-workspace` | Настройки VS Code, не нужны для работы сервиса |
