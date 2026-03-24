import logging
from utils.doc_converter import convert_doc_to_docx

logging.basicConfig(level=logging.INFO, format="%(message)s")

def main():
    folder_path = input("Путь к папке с .doc: ").strip().strip('"')
    if not folder_path:
        return
    convert_doc_to_docx(folder_path)

if __name__ == "__main__":
    main()