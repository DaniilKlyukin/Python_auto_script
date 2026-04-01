import os
from utils.pdf_generator import PDFGenerator


def main():
    print("=== Конвертация DOCX и PPTX в PDF ===")
    path = input("Введите путь к папке: ").strip().strip('"')

    if not os.path.isdir(path):
        print("Путь не найден.")
        return

    files = [f for f in os.listdir(path) if not f.startswith('~$')]
    docx_files = [f for f in files if f.lower().endswith('.docx')]
    pptx_files = [f for f in files if f.lower().endswith('.pptx')]

    if not docx_files and not pptx_files:
        print("Подходящие файлы не найдены.")
        return

    print(f"Найдено: DOCX - {len(docx_files)}, PPTX - {len(pptx_files)}")

    generator = PDFGenerator()
    success_count = 0

    try:
        for filename in docx_files:
            full_path = os.path.join(path, filename)
            if generator.convert_docx(full_path):
                print(f"[OK] {filename} -> PDF")
                success_count += 1
            else:
                print(f"[FAIL] {filename}")

        for filename in pptx_files:
            full_path = os.path.join(path, filename)
            if generator.convert_pptx(full_path):
                print(f"[OK] {filename} -> PDF")
                success_count += 1
            else:
                print(f"[FAIL] {filename}")
    finally:
        generator.quit()

    total = len(docx_files) + len(pptx_files)
    print(f"\n[Готово] Успешно: {success_count} из {total}")


if __name__ == "__main__":
    main()