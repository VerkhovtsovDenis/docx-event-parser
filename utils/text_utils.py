def normalize_text(text: str) -> str:
    """Нормализация текста"""
    return " ".join(text.strip().split())


def extract_phone(text: str) -> str:
    """Извлечение телефона из текста"""
    return "".join(filter(str.isdigit, text))
