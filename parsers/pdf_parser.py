from typing import Any, Dict, Tuple

import pdfplumber

from config.settings import TABLE_FIELDS_MAPPING

from .base_parser import BaseParser


def parse_old_pdf_format(text: str) -> Dict[str, Any]:
    """Парсинг старых PDF-файлов с нумерованными пунктами"""
    result = {
        "Event name": "",
        "Department": "Молодежный коворкинг А11",  # По умолчанию
        "Date of event": "",
        "Date of installation": "",
        "Order": "",
        "Participants": "",
        "Responsible": "",
        "Event format": "",
        "Guests of honor": "",
        "Event level": "",
        "Schedule": "",
        "Necessary technical equipment": "",
        "Training on working with audio equipment": "",
    }

    # Полный маппинг всех полей старого формата
    field_mapping = {
        "1": ("Responsible", "Заявитель (ФИО)"),
        "2": ("Date of event", "Дата и время бронирования"),
        "3": ("Event format", "Формат проведения мероприятия"),
        "4": ("Participants", "Контингент (кол-во, состав)"),
        "5": ("Event name", "Повестка/программа"),
        "6.1": ("Necessary technical equipment", "Телевизоры/проектор"),
        "6.2": ("Necessary technical equipment", "Звуковая аппаратура"),
        "6.3": (
            "Training on working with audio equipment",
            "Обучение работе с техникой",
        ),
        "8": (
            "Event level",
            "Требования к посадке участников",
        ),  # Используем для уровня
        "9.1": ("Responsible", "Ответственный организатор (ФИО)"),
        "9.2": ("Responsible_phone", "Номер телефона"),
        "9.3": ("Additional_requirements", "Дополнительные требования"),
    }

    # Разбиваем текст на строки
    lines = [line.strip() for line in text.split("\n") if line.strip()]

    current_field = None
    collected_data = {}

    for line in lines:
        # Ищем строки с номером пункта (|1|, |2| и т.д.)
        if line.startswith("|") and "|" in line[1:]:
            parts = [p.strip() for p in line.split("|") if p.strip()]
            if len(parts) >= 3:
                point_num = parts[0]
                if point_num in field_mapping:
                    field, _ = field_mapping[point_num]
                    collected_data[field] = parts[2]
                    current_field = field

    # Специальная обработка для технического оборудования
    tech_equipment = []
    if "6.1" in field_mapping and field_mapping["6.1"][0] in collected_data:
        tech_equipment.append(collected_data[field_mapping["6.1"][0]])
    if "6.2" in field_mapping and field_mapping["6.2"][0] in collected_data:
        tech_equipment.append(collected_data[field_mapping["6.2"][0]])
    if tech_equipment:
        result["Necessary technical equipment"] = ", ".join(tech_equipment)

    # Переносим все собранные данные в результат
    for field in result:
        if field in collected_data:
            result[field] = collected_data[field]

    # Особые случаи:
    if "Responsible_phone" in collected_data:
        result[
            "Responsible"
        ] += f" (тел.: {collected_data['Responsible_phone']})"

    if "Additional_requirements" in collected_data:
        if collected_data["Additional_requirements"] not in ["-", ""]:
            result[
                "Event format"
            ] += f". Доп.требования: {collected_data['Additional_requirements']}"

    return result


class PDFParser(BaseParser):
    def parse(self, file_path: str) -> Tuple[Dict[str, Any], str]:
        try:
            with pdfplumber.open(file_path) as pdf:
                if not pdf.pages:
                    return {}, "empty_pdf"

                first_page = pdf.pages[0]
                text = first_page.extract_text()

                # Улучшенное определение старого формата
                if self._is_old_format(text):
                    return parse_old_pdf_format(text), "old_pdf_format"

                # Пробуем распарсить как таблицы
                tables = first_page.extract_tables()
                if tables and self._validate_tables(tables):
                    return self._parse_tables(tables), "pdf_table"

                # Если не распознано как таблицы, пробуем текст
                return self._parse_text(text), "pdf_text"

        except Exception as e:
            print(f"PDF parsing error: {str(e)}")
            return {}, "error"

    def _is_old_format(self, text: str) -> bool:
        """Определяет, является ли PDF старым форматом"""
        if not text:
            return False

        # Проверяем несколько характерных признаков старого формата
        lines = text.split("\n")
        
        old_format_indicators = [
            any(line.strip().startswith("9.3") for line in lines)
            for i in range(1, 6)
        ]

        # Должны быть хотя бы 3 совпадения из 5 возможных
        return sum(old_format_indicators) >= 3

    def _validate_tables(self, tables: list) -> bool:
        """Проверяет, что таблицы соответствуют новому формату"""
        if not tables or len(tables) < 1:
            return False

        # Проверяем первую таблицу на наличие характерных заголовков
        first_table = tables[0]
        if len(first_table) < 1 or len(first_table[0]) < 2:
            return False

        # Ищем хотя бы один известный заголовок
        known_headers = [
            "Название мероприятия",
            "Организатор",
            "Даты проведения",
        ]
        return any(
            any(header in str(cell) for cell in first_table[0])
            for header in known_headers
        )

    def _parse_tables(self, tables: list) -> Dict[str, Any]:
        """
        Парсит данные из таблиц PDF с использованием маппинга из settings.py
        """
        extracted_data = {key: "" for key in TABLE_FIELDS_MAPPING.keys()}

        try:
            # Обработка основной таблицы
            if tables and len(tables[0]) > 0:
                for row in tables[0]:
                    if (
                        len(row) >= 4
                    ):  # Предполагаем формат [key1, key2, None, value]
                        key_part1 = self.clean_text(row[0] or "")
                        key_part2 = self.clean_text(row[1] or "")
                        value = self.clean_text(row[3] or "")

                        if not value:
                            continue

                        # Ищем соответствие в маппинге полей
                        for field, keywords in TABLE_FIELDS_MAPPING.items():
                            if any(
                                keyword in key_part1 or keyword in key_part2
                                for keyword in keywords
                            ):
                                extracted_data[field] = value
                                break

            # Обработка дополнительных таблиц (если есть)
            for table in tables[1:]:
                for row in table:
                    if len(row) >= 4:
                        key_part1 = self.clean_text(row[0] or "")
                        key_part2 = self.clean_text(row[1] or "")
                        value = self.clean_text(row[3] or "")

                        if not value:
                            continue

                        for field, keywords in TABLE_FIELDS_MAPPING.items():
                            if not extracted_data[field] and any(
                                keyword in key_part1 or keyword in key_part2
                                for keyword in keywords
                            ):
                                extracted_data[field] = value
                                break

            return extracted_data

        except Exception as e:
            print(f"Error parsing PDF tables: {str(e)}")
            return extracted_data

    def _parse_text(self, text: str) -> Dict[str, Any]:
        """
        Парсит текст PDF, когда не удалось извлечь таблицы
        """
        extracted_data = {key: "" for key in TABLE_FIELDS_MAPPING.keys()}

        if not text:
            return extracted_data

        lines = [line.strip() for line in text.split("\n") if line.strip()]
        current_field = None

        for line in lines:
            # Проверяем, начинается ли строка с известного ключевого слова
            found_field = None
            for field, keywords in TABLE_FIELDS_MAPPING.items():
                if any(line.startswith(keyword) for keyword in keywords):
                    found_field = field
                    # Разделяем строку на ключ и значение
                    for keyword in keywords:
                        if line.startswith(keyword):
                            value = line[len(keyword)].strip(":")
                            extracted_data[field] = value
                            current_field = field
                            break
                    break

            if not found_field and current_field:
                # Продолжение предыдущего поля
                extracted_data[current_field] += " " + line

        return extracted_data
