import win32com.client
import os
from pathlib import Path

class PDFGenerator:
    @staticmethod
    def convert(docx_path: str, pdf_path: str = None):
        pdf_path = pdf_path or str(Path(docx_path).with_suffix('.pdf'))
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        try:
            doc = word.Documents.Open(os.path.abspath(docx_path))
            doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
            doc.Close()
            return True
        except Exception: return False
        finally: word.Quit()