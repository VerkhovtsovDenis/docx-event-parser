# Настройки парсера
DEFAULT_OUTPUT_FORMATS = ["xlsx", "json"]
ALLOWED_FILE_TYPES = [".pdf", ".docx"]
TABLE_FIELDS_MAPPING = {
    "Event name": ["Название мероприятия"],
    "Department": ["Организатор", "Подразделение"],
    "Date of event": ["Даты проведения мероприятия"],
    "Date of installation": ["Даты монтажа", "подготовки площадки"],
    "Order": ["Приказ об организации"],
    "Participants": ["Количество участников", "контингент"],
    "Responsible": ["Ответственный за проведение"],
    "Event format": ["Формат мероприятия"],
    "Guests of honor": ["Почетные гости", "ведущие мероприятия"],
    "Event level": ["Уровень мероприятия"],
    "Schedule": ["Расписание", "разбивка по времени"],
    "Necessary technical equipment": ["Необходимое техническое оснащение"],
    "Training on working with audio equipment": [
        "Обучение работе",
        "звуковом оборудовании",
    ],
}
