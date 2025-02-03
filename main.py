import os
import json
import pandas as pd
from typing import Dict, List, Tuple, Any
from docx import Document


def clear(output_xlsx: str, output_json: str) -> None:
    """
    Очищает файлы output.xlsx и output.json, если они существуют.
    """
    if os.path.exists(output_xlsx):
        os.remove(output_xlsx)
    if os.path.exists(output_json):
        os.remove(output_json)


def create(output_xlsx: str, output_json: str) -> None:
    """
    Создает пустые файлы output.xlsx и output.json.
    """
    df = pd.DataFrame(columns=['Filename', 'Status'])
    df.to_excel(output_xlsx, index=False)
    with open(output_json, 'w', encoding='utf-8') as json_file:
        json.dump({}, json_file, ensure_ascii=False, indent=4)


def find_docx_files(directory: str) -> List[str]:
    """
    Рекурсивно ищет все .docx файлы в указанной папке и подпапках.

    :param directory: Путь к корневой папке.
    :return: Список путей к файлам .docx.
    """
    docx_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith(".docx"):
                docx_files.append(os.path.join(root, file))
    return docx_files


def parse_first_page_tables(doc_path: str) -> Dict[str, Any]:
    """
    Извлекает данные из первых трех таблиц первой страницы документа .docx.

    :param doc_path: Путь к .docx файлу.
    :return: Словарь с ключами, соответствующими полям json-результата.
    """
    doc = Document(doc_path)
    tables = doc.tables[:3]  # Берем только первые три таблицы

    extracted_data = {
        'Event name': '',
        'Department': '',
        'Date of event': '',
        'Date of installation': '',
        'Order': '',
        'Participants': '',
        'Responsible': '',
        'Event format': '',
        'Guests of honor': '',
        'Event level': '',
        'Schedule': '',
        'Necessary technical equipment': '',
        'Training on working with audio equipment': ''
    }

    try:
        if len(tables) >= 3:
            extracted_data['Event name'] = tables[0].rows[0].cells[1].text.strip()
            extracted_data['Department'] = tables[0].rows[1].cells[1].text.strip()
            extracted_data['Date of event'] = tables[0].rows[2].cells[1].text.strip()
            extracted_data['Date of installation'] = tables[0].rows[3].cells[1].text.strip()
            extracted_data['Order'] = tables[0].rows[4].cells[1].text.strip()
            extracted_data['Participants'] = tables[0].rows[5].cells[1].text.strip()
            extracted_data['Responsible'] = tables[0].rows[6].cells[1].text.strip()
            extracted_data['Event format'] = tables[1].rows[0].cells[1].text.strip()
            extracted_data['Guests of honor'] = tables[1].rows[1].cells[1].text.strip()
            extracted_data['Event level'] = tables[1].rows[2].cells[1].text.strip()
            extracted_data['Schedule'] = tables[1].rows[3].cells[1].text.strip()
            extracted_data['Necessary technical equipment'] = tables[2].rows[0].cells[1].text.strip()
            extracted_data['Training on working with audio equipment'] = tables[2].rows[1].cells[1].text.strip()
    except IndexError:
        pass  # В случае проблем с доступом к ячейкам оставляем пустые значения

    return extracted_data


def process_files(input_dir: str, output_xlsx: str, output_json: str) -> None:
    """
    Обрабатывает все .docx файлы в папке, создавая итоговые xlsx и json файлы.

    :param input_dir: Папка, где находятся .docx файлы.
    :param output_xlsx: Путь к выходному .xlsx файлу.
    :param output_json: Путь к выходному .json файлу.
    """
    docx_files = find_docx_files(input_dir)
    results = {}
    processing_status = []

    for file in docx_files:
        try:
            parsed_data = parse_first_page_tables(file)
            results[os.path.basename(file)] = parsed_data
            processing_status.append({'Filename': os.path.basename(file), 'Status': 'Processed'})
            print(f"{os.path.basename(file)} - OK")
        except Exception as e:
            processing_status.append({'Filename': os.path.basename(file), 'Status': f'Error: {str(e)}'})
            print(f"{os.path.basename(file)} - ERROR: {str(e)}")

    # Сохранение данных в JSON
    with open(output_json, 'w', encoding='utf-8') as json_file:
        json.dump(results, json_file, ensure_ascii=False, indent=4)

    # Сохранение статусов в XLSX
    df = pd.DataFrame(processing_status)
    df.to_excel(output_xlsx, index=False)


def main() -> None:
    """
    Главная функция, запускающая процесс обработки файлов .docx.
    """
    input_directory = "inputs"
    output_xlsx_path = "output.xlsx"
    output_json_path = "output.json"

    clear(output_xlsx_path, output_json_path)
    create(output_xlsx_path, output_json_path)

    process_files(input_directory, output_xlsx_path, output_json_path)
    print("Обработка завершена!")


if __name__ == "__main__":
    main()
