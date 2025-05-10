from typing import Any, Dict, Tuple

import pdfplumber

from config.settings import TABLE_FIELDS_MAPPING

from .base_parser import BaseParser


def parse_old_pdf_format(text: str) -> Dict[str, Any]:
    """Парсинг старых PDF-файлов с нумерованными пунктами"""
    result = {
        "Event name": "",
        "Department": "",
        "Date of event": "",
        "Participants": "",
        "Responsible": "",
        "Event format": "",
        "Necessary technical equipment": "",
        "Training on working with audio equipment": "",
    }

    # Словарь соответствия номеров пунктов полям
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
        "9.1": ("Responsible", "Ответственный организатор (ФИО)"),
    }

    # Разбиваем текст на строки
    lines = [line.strip() for line in text.split("\n") if line.strip()]

    current_field = None
    collected_data = {}

    for line in lines:
        # Проверяем, начинается ли строка с номера пункта
        parts = line.split(maxsplit=2)
        if len(parts) >= 2 and parts[0].replace(".", "").isdigit():
            point_num = parts[0]
            if point_num in field_mapping:
                field, field_name = field_mapping[point_num]
                value = parts[2] if len(parts) > 2 else ""
                collected_data[field] = value.strip()
                current_field = field
        elif current_field:
            # Продолжение предыдущего пункта
            if collected_data.get(current_field):
                collected_data[current_field] += " " + line
            else:
                collected_data[current_field] = line

    # Объединяем данные из нескольких связанных полей
    if "Necessary technical equipment" in collected_data:
        tech_equip = []
        for subpoint in ["6.1", "6.2"]:
            if subpoint in field_mapping:
                field = field_mapping[subpoint][0]
                if collected_data.get(field):
                    tech_equip.append(collected_data[field])
        result["Necessary technical equipment"] = ", ".join(
            filter(None, tech_equip)
        )

    if "6.3" in field_mapping:
        field = field_mapping["6.3"][0]
        result[field] = collected_data.get(field, "")

    # Заполняем основные поля
    for field in result:
        if field in collected_data:
            if field == "Responsible" and result[field]:
                # Объединяем данные из пунктов 1 и 9.1
                result[field] += " / " + collected_data[field]
            else:
                result[field] = collected_data.get(field, "")

    return result


class PDFParser(BaseParser):
    def parse(self, file_path: str) -> Tuple[Dict[str, Any], str]:
        try:
            with pdfplumber.open(file_path) as pdf:
                if not pdf.pages:
                    return {}, "empty_pdf"

                first_page = pdf.pages[0]
                text = first_page.extract_text()

                # Проверяем, является ли это старым форматом
                if any(
                    line.strip().startswith("|1|") for line in text.split("\n")
                ):
                    return parse_old_pdf_format(text), "old_pdf_format"

                tables = first_page.extract_tables()
                if tables:
                    return self._parse_tables(tables), "pdf_table"
                return self._parse_text(text), "pdf_text"

        except Exception as e:
            print(f"PDF parsing error: {str(e)}")
            return {}, "error"

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
