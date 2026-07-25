# Базовый образ с Python
FROM python:3.9-slim

# Устанавливаем рабочую директорию внутри контейнера
WORKDIR /app

# Копируем файлы проекта в контейнер
COPY . /app

# Устанавливаем зависимости (если есть requirements.txt)
RUN pip install --no-cache-dir -r requirements.txt

# Открываем порт (обычно 5000 для Flask)
EXPOSE 5000

# Команда для запуска сайта
CMD ["python", "app.py"]