#!/bin/bash
PROJECT_DIR="/home/kosinov.va/Projects/processing_schema/ifc_inspect"
VENV_DIR="/home/kosinov.va/Projects/processing_schema/ifc_inspect.venv"

cd "$PROJECT_DIR" || exit

# Проверка и активация venv (если существует)
if [ -d "$VENV_DIR" ]; then
    source "$VENV_DIR/bin/activate"
fi

# Основные настройки Flask
export FLASK_APP=backend/app.py
export FLASK_ENV=development

# Пути к данным
export UPLOAD_FOLDER="$PROJECT_DIR/uploads"
export OUTPUT_FOLDER="$PROJECT_DIR/outputs"

# Создание необходимых директорий
mkdir -p "$UPLOAD_FOLDER"
mkdir -p "$OUTPUT_FOLDER"

echo "✅ Запуск сервера IFC Inspect..."
echo ""

# Запуск сервера на порту 6003
flask run --host=0.0.0.0 --port=6003
