# IFC Inspect

Сервис анализа проектной документации для строительства: обработка PDF файлов (проектная документация) и IFC моделей (информационное моделирование зданий) с генерацией Excel отчетов.

## 📖 О проекте

IFC Inspect — это веб-сервис, который автоматизирует анализ проектной документации:
- **PDF файлы**: автоматическое разделение на текстовые страницы и чертежи по размеру
- **IFC модели**: извлечение информации о здании, элементах конструкции, материалах и свойствах
- **Excel отчеты**: детализированные таблицы по всем элементам модели

## 🏗️ Структура проекта

```
/workspace/
├── backend/
│   └── app.py              # Flask API сервер (5 endpoints)
├── frontend/
│   └── index.html          # Single-page приложение (Vanilla JS)
├── scripts/
│   ├── page_splitter.py    # Классификация PDF страниц (текст vs чертежи)
│   └── ifc_parser.py       # Парсинг IFC → Excel отчет
├── uploads/                # Сессионные данные (создается автоматически)
├── requirements.txt        # Зависимости Python
├── start-dev.sh           # Скрипт запуска (продакшен, порт 6003)
├── tunnel.sh              # SSH туннель для удаленного доступа
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

Сервер запустится на порту **5001**: http://localhost:5001

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

Затем откройте http://localhost:8080 в браузере.

## 📡 API Endpoints

| Метод | URL | Описание | Параметры |
|-------|-----|----------|-----------|
| GET | `/` | Веб-интерфейс | — |
| POST | `/api/upload` | Загрузка файлов | `files` (multipart/form-data) |
| POST | `/api/process` | Обработка файлов | `{session_id: string}` |
| GET | `/api/results/<session_id>` | Получение результатов | — |
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
curl -X POST http://localhost:5001/api/process \
  -H "Content-Type: application/json" \
  -d '{"session_id": "uuid-string"}'
```

#### Получение результатов

```bash
curl http://localhost:5001/api/results/uuid-string
```

## 🔄 Алгоритм работы

1. **Загрузка**: Пользователь загружает PDF и IFC файлы через веб-интерфейс
2. **Создание сессии**: Сервер генерирует unique session_id и сохраняет файлы в `uploads/<session_id>/`
3. **Обработка** (по кнопке "Запустить анализ"):
   - **PDF**: 
     - Открытие через PyMuPDF (fitz)
     - Измерение размера каждой страницы
     - Классификация: текст (≤30см) или чертеж (>30см)
     - Поворот страниц при необходимости
     - Сохранение отдельных PDF файлов в папки `text_pages/` и `drawing_pages/`
   - **IFC**:
     - Открытие через ifcopenshell
     - Извлечение информации о проекте
     - Группировка элементов по типам
     - Сбор данных о материалах и этажах
     - Генерация Excel отчета через openpyxl
4. **Сохранение результатов**:
   - `results.json` — детали обработки (списки страниц, результаты IFC)
   - `session_info.json` — статус сессии и сводка
5. **Отображение**: Фронтенд получает результаты и показывает:
   - Количество текстовых страниц и чертежей
   - Статистику по элементам IFC
   - Кнопку для скачивания Excel отчета

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
| IFC парсинг | ifcopenshell 0.8.5 |
| Excel генерация | openpyxl ≥3.1.5, pandas 2.1.4 |
| Frontend | HTML5, CSS3, Vanilla JavaScript |

## ⚙️ Конфигурация

Настройки заданы напрямую в коде (без .env файлов):

| Параметр | Значение | Файл |
|----------|----------|------|
| Порт разработки | 5001 | backend/app.py |
| Порт продакшена | 6003 | start-dev.sh |
| Папка загрузок | uploads/ | backend/app.py |
| Debug режим | True | backend/app.py |

## 🔐 Безопасность

- Нет аутентификации/авторизации (локальное использование)
- Сессионные данные хранятся в JSON файлах
- Рекомендуется использовать за обратным прокси (nginx) в продакшене

## 📝 Лицензия

MIT
