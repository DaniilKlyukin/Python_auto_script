import os
import re
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor

class FilenameCleaner:
    def __init__(self, root_dir: str, max_workers: int = 4):
        self.root_dir = Path(root_dir)
        self.max_workers = max_workers

    def _sanitize(self, filename: str) -> str:
        name, ext = os.path.splitext(filename)
        new_name = re.sub(r'^[\+\-\s]+', '', name).replace('+', '').replace('.', ' ')
        return f"{re.sub(r'\s+', ' ', new_name).strip()}{ext}"

    def _process_file(self, file_path: Path):
        try:
            new_name = self._sanitize(file_path.name)
            if new_name == file_path.name: return
            new_path = file_path.with_name(new_name)
            c = 1
            while new_path.exists():
                new_path = file_path.with_name(f"{Path(new_name).stem}_{c}{new_path.suffix}")
                c += 1
            file_path.rename(new_path)
        except Exception: pass

    def run(self):
        files = [Path(root) / f for root, _, fs in os.walk(self.root_dir) for f in fs]
        with ThreadPoolExecutor(max_workers=self.max_workers) as e:
            e.map(self._process_file, files)