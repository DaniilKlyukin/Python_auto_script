import win32com.client
import os
from pathlib import Path

class PDFGenerator:
    def __init__(self):
        self.word = None
        self.ppt = None

    def _get_word(self):
        if self.word is None:
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = False
        return self.word

    def _get_ppt(self):
        if self.ppt is None:
            self.ppt = win32com.client.Dispatch("PowerPoint.Application")
        return self.ppt

    def convert_docx(self, docx_path):
        pdf_path = str(Path(docx_path).with_suffix('.pdf'))
        try:
            word = self._get_word()
            doc = word.Documents.Open(os.path.abspath(docx_path))
            doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
            doc.Close()
            return True
        except Exception:
            return False

    def convert_pptx(self, pptx_path):
        pdf_path = str(Path(pptx_path).with_suffix('.pdf'))
        try:
            ppt = self._get_ppt()
            pres = ppt.Presentations.Open(os.path.abspath(pptx_path), WithWindow=False)
            pres.SaveAs(os.path.abspath(pdf_path), FileFormat=32)
            pres.Close()
            return True
        except Exception:
            return False

    def quit(self):
        if self.word:
            self.word.Quit()
        if self.ppt:
            self.ppt.Quit()