import json
import os
import pandas as pd
from typing import Dict, Any

from parsers.docx_parser import DocxParser
from parsers.pdf_parser import PDFParser
from utils.file_utils import clear_output_files, find_files


def save_to_excel(data: Dict[str, Dict[str, Any]], filename: str) -> None:
    """Сохраняет данные в Excel файл"""
    if not data:
        print("Нет данных для сохранения в Excel")
        return

    # Убедимся, что filename имеет правильное расширение
    if not filename.endswith(".xlsx"):
        filename = os.path.splitext(filename)[0] + ".xlsx"

    # Преобразуем данные в DataFrame
    df_data = []
    for filename_key, fields in data.items():
        row = {"Filename": filename_key}
        row.update(fields)
        df_data.append(row)

    df = pd.DataFrame(df_data)

    # Сохраняем с указанием движка (engine='openpyxl')
    df.to_excel(filename, index=False, engine="openpyxl")


def save_to_json(data: Dict[str, Dict[str, Any]], filename: str) -> None:
    """Сохраняет данные в JSON файл"""
    if not filename.endswith(".json"):
        filename = os.path.splitext(filename)[0] + ".json"

    with open(filename, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


def main():
    # Конфигурация путей
    input_dir = "inputs"
    output_xlsx = "output.xlsx"  # Важно: должно быть .xlsx
    output_json = "output.json"  # Важно: должно быть .json

    # Очистка предыдущих результатов
    clear_output_files(output_xlsx, output_json)

    # Инициализация парсеров
    pdf_parser = PDFParser()
    docx_parser = DocxParser()

    # Поиск файлов
    docx_files, pdf_files = find_files(input_dir)
    results = {}

    # Обработка PDF файлов
    for pdf_file in pdf_files:
        try:
            data, doc_type = pdf_parser.parse(pdf_file)
            if data:  # Добавляем только если есть данные
                results[os.path.basename(pdf_file)] = data
                print(f"Обработан PDF: {pdf_file} ({doc_type})")
        except Exception as e:
            print(f"Ошибка при обработке PDF {pdf_file}: {str(e)}")

    # Обработка DOCX файлов
    for docx_file in docx_files:
        try:
            data, doc_type = docx_parser.parse(docx_file)
            if data:  # Добавляем только если есть данные
                results[os.path.basename(docx_file)] = data
                print(f"Обработан DOCX: {docx_file} ({doc_type})")
        except Exception as e:
            print(f"Ошибка при обработке DOCX {docx_file}: {str(e)}")

    # Сохранение результатов
    if results:
        save_to_json(results, output_json)
        save_to_excel(results, output_xlsx)
        print(f"Результаты сохранены в {output_json} и {output_xlsx}")
    else:
        print("Нет данных для сохранения")


if __name__ == "__main__":
    main()
