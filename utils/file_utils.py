import os
from typing import List, Tuple


def find_files(directory: str) -> Tuple[List[str], List[str]]:
    """Поиск файлов в директории"""
    docx_files, pdf_files = [], []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith(".docx"):
                docx_files.append(os.path.join(root, file))
            elif file.endswith(".pdf"):
                pdf_files.append(os.path.join(root, file))
    return docx_files, pdf_files


def clear_output_files(*files):
    """Очистка выходных файлов"""
    for file in files:
        if os.path.exists(file):
            os.remove(file)
