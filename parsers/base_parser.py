from abc import ABC, abstractmethod
from typing import Any, Dict


class BaseParser(ABC):
    @abstractmethod
    def parse(self, file_path: str) -> Dict[str, Any]:
        """Базовый метод для парсинга файла"""
        pass

    @staticmethod
    def clean_text(text: str) -> str:
        """Очистка текста от лишних пробелов и переносов"""
        return ' '.join(text.replace('\n', ' ').split()) if text else ''
