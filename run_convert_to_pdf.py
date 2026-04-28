import os
from utils.pdf_generator import PDFGenerator

def main():
    print("=== Рекурсивная конвертация DOCX/PPTX в PDF ===")
    path = input("Введите путь к папке (просканирую и все подпапки): ").strip().strip('"')

    if not os.path.isdir(path):
        print("Ошибка: Указанный путь не является папкой.")
        return

    generator = PDFGenerator()
    try:
        generator.process_folder(path)
    except KeyboardInterrupt:
        print("\nПроцесс прерван пользователем.")
    finally:
        generator.quit()
        input("\nНажмите Enter, чтобы выйти...")

if __name__ == "__main__":
    main()