# Textonator
flowchart TD
    A[Пользователь] -->|Отправляет файл| B[Telegram Bot]
    B --> C{Выбор действия}
    C -->|Полный анализ| D[Извлечение аудио]
    C -->|Обрезать| E[FFmpeg: обрезка]
    E --> D
    D --> F[Faster-Whisper: транскрибация]
    F --> G[YandexGPT: анализ]
    G --> H[Mind Map]
    G --> I[Статистика]
    G --> J[Тайм-коды]
    G --> K[Протокол .docx]
    H --> L[Пользователь]
    I --> L
    J --> L
    K --> L
    F --> M[Google Translate]
    M --> L
