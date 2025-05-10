import json
import os
import pandas as pd
from typing import Dict, Any

from parsers.docx_parser import DocxParser
from parsers.pdf_parser import PDFParser
from utils.file_utils import clear_output_files, find_files


def process_files(
    pdf_parser: PDFParser, docx_parser: DocxParser, files: list, file_type: str
) -> Dict[str, Any]:
    """Обрабатывает файлы указанного типа и возвращает результаты"""
    results = {}
    for file in files:
        try:
            if file_type == "pdf":
                data, doc_type = pdf_parser.parse(file)
            else:
                data, doc_type = docx_parser.parse(file)

            if data:
                results[os.path.basename(file)] = data
                print(f"Обработан {file_type.upper()}: {file} ({doc_type})")
        except Exception as e:
            print(f"Ошибка при обработке {file_type.upper()} {file}: {str(e)}")
    return results


def save_results(
    results: Dict[str, Any], output_xlsx: str, output_json: str
) -> None:
    """Сохраняет результаты в файлы"""
    if not results:
        print("Нет данных для сохранения")
        return

    # Подготовка данных для Excel
    excel_data = []
    for filename, fields in results.items():
        row = {
            "Filename": filename,
            "File Type": fields.get("doc_type", "unknown"),
        }
        row.update({k: v for k, v in fields.items() if k != "doc_type"})
        excel_data.append(row)

    # Сохранение в Excel
    df = pd.DataFrame(excel_data)
    df.to_excel(output_xlsx, index=False, engine="openpyxl")

    # Сохранение в JSON
    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=4)

    print(f"Результаты сохранены в {output_xlsx} и {output_json}")


def main():
    # Конфигурация путей
    input_dir = "inputs"
    output_xlsx = "output.xlsx"
    output_json = "output.json"

    # Очистка предыдущих результатов
    clear_output_files(output_xlsx, output_json)

    # Инициализация парсеров
    pdf_parser = PDFParser()
    docx_parser = DocxParser()

    # Поиск файлов
    docx_files, pdf_files = find_files(input_dir)

    # Обработка файлов
    pdf_results = process_files(pdf_parser, docx_parser, pdf_files, "pdf")
    docx_results = process_files(pdf_parser, docx_parser, docx_files, "docx")

    # Объединение результатов
    all_results = {**pdf_results, **docx_results}

    # Сохранение результатов
    save_results(all_results, output_xlsx, output_json)


if __name__ == "__main__":
    main()
