import shutil
import os
from zipfile import ZipFile
from pathlib import Path
import logging


def remake_file(file_path: Path):

    # Убедитесь, что файл существует и имеет правильное расширение
    if not file_path.is_file() or not file_path.suffix.lower() == ".xlsx":
        return

    # Создание временной папки для извлечения
    tmp_folder = file_path.parent / "tmp"
    os.makedirs(tmp_folder, exist_ok=True)

    # Извлечение содержимого Excel файла
    with ZipFile(file_path) as excel_container:
        excel_container.extractall(tmp_folder)

    wrong_file_path = tmp_folder / "xl" / "SharedStrings.xml"
    correct_file_path = tmp_folder / "xl" / "sharedStrings.xml"

    # Переименование файла
    try:
        os.rename(wrong_file_path, correct_file_path)
    except FileNotFoundError:
        pass

    # Удаление оригинального файла
    os.remove(file_path)

    # Создание нового файла Excel
    shutil.make_archive(file_path.stem, "zip", tmp_folder)
    os.rename(file_path.stem + ".zip", file_path)

    # Очистка временной папки
    shutil.rmtree(tmp_folder)
