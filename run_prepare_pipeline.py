import os
import logging
from utils.doc_converter import convert_doc_to_docx
from utils.media_cleaner import WordImageCleanerDocx
from services.approval_processor import process_docx, generate_years
from services.signature_processor import process_docx_signatures
from utils.filename_cleaner import FilenameCleaner
from utils.file_cleaner import FileCleaner

logging.basicConfig(level=logging.INFO, format="%(message)s")

def main():
    folder_path = input("Введите путь к папке: ").strip().strip('"')
    if not os.path.isdir(folder_path):
        print("Ошибка: Путь не найден.")
        return

    try:
        start_y = int(input("Введите начальный год (для согласования): "))
        end_y = int(input("Введите конечный год (для согласования): "))
    except ValueError:
        print("Ошибка: Введите корректные числа для годов.")
        return

    old_fio = input("Введите ФИО, которое нужно заменить (подписи): ").strip()
    new_fio = input("Введите новое ФИО: ").strip()

    print("\n[1/6] Очистка папки от PDF и изображений...")
    deleted_count = FileCleaner.cleanup_folder(folder_path)
    print(f"Удалено файлов: {deleted_count}")

    print("\n[2/6] Конвертация .doc в .docx...")
    convert_doc_to_docx(folder_path)

    print("\n[3/6] Оптимизация медиа-объектов...")
    cleaner = WordImageCleanerDocx(folder_path)
    cleaner.process_all()

    docx_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.docx') and not f.startswith('~$')]

    print("\n[4/6] Обновление учебных лет в Листах согласования...")
    years_list = generate_years(start_y, end_y)
    for filename in docx_files:
        process_docx(os.path.join(folder_path, filename), years_list)

    print("\n[5/6] Замена ФИО в зонах подписей...")
    for filename in docx_files:
        process_docx_signatures(os.path.join(folder_path, filename), old_fio, new_fio)

    print("\n[6/6] Очистка имен файлов...")
    fn_cleaner = FilenameCleaner(folder_path)
    fn_cleaner.run()

    print("\n=== Все операции успешно завершены ===")

if __name__ == "__main__":
    main()