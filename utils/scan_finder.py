import os
import re
from typing import List, Tuple, Optional
from pathlib import Path

class ScanFinder:
    def __init__(self, scans_folder: str, extensions: List[str] = None):
        self.scans_folder = scans_folder
        self.extensions = extensions or ['.jpg', '.jpeg', '.png', '.pdf']

    def find_scans_for_program(self, program_name: str) -> Optional[Tuple[str, str, str]]:
        base_name = self._extract_base(program_name)
        if not base_name: return None
        files = self._find_matching(base_name)
        return self._sort_scans(files)

    def _extract_base(self, program_name: str) -> Optional[str]:
        match = re.search(r'РП\s+(\S+)(?=\s|$)', Path(program_name).stem)
        return match[0].split(' ')[1].strip() if match else None

    def _find_matching(self, program_name: str) -> List[str]:
        files = []
        for f in os.listdir(self.scans_folder):
            if not any(f.lower().endswith(ext) for ext in self.extensions): continue
            match = re.match(r'^(.+?)([123])\.(?:png|jpg|jpeg|pdf)$', f.lower())
            if match and program_name.lower() in match.group(1):
                files.append(os.path.join(self.scans_folder, f))
        return files

    def _sort_scans(self, files: List[str]) -> Optional[Tuple[str, str, str]]:
        if len(files) != 3: return None
        return tuple(sorted(files, key=lambda p: int(re.search(r'(\d+)', Path(p).stem).group(1) if re.search(r'(\d+)', Path(p).stem) else 0)))