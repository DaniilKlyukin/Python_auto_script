import logging
from pathlib import Path
from core.docx_editor import DocxEditor
from utils.scan_finder import ScanFinder

logger = logging.getLogger(__name__)

class ScanInsertionManager:
    def __init__(self, scan_finder: ScanFinder):
        self.scan_finder = scan_finder

    def process_documents(self, doc_paths: list):
        for path in doc_paths:
            self._process_single(path)

    def _process_single(self, doc_path: str):
        file_name = Path(doc_path).name
        scans = self.scan_finder.find_scans_for_program(file_name)
        if not scans or len(scans) < 3: return

        try:
            with DocxEditor(doc_path) as editor:
                editor.add_image_at_beginning(scans[0])
                if not editor.insert_image_after_text("АННОТАЦИЯ", scans[1]):
                    if len(editor.doc.paragraphs) > 5:
                        editor.add_floating_scan(editor.doc.paragraphs[5], scans[1])
                if not editor.insert_image_after_text("Лист согласования", scans[2]):
                    editor.add_image_at_end(scans[2])
        except Exception as e:
            logger.error(f"Ошибка {file_name}: {e}")