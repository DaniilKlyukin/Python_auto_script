import os
from pathlib import Path
from fpdf import FPDF


class ImageToPDFService:
    def __init__(self, images_per_pdf: int = 3):
        self.images_per_pdf = images_per_pdf
        self.supported_ext = ('.jpg', '.jpeg', '.png')

    def generate_pdfs(self, input_path: str, output_path: str = None):
        src_dir = Path(input_path)
        if not src_dir.exists(): return
        dst_dir = Path(output_path) if output_path else src_dir / "PDF_Output"
        dst_dir.mkdir(parents=True, exist_ok=True)
        files = sorted(
            [f for f in os.listdir(src_dir) if f.lower().endswith(self.supported_ext) and not f.startswith('~$')])

        groups = [files[i:i + self.images_per_pdf] for i in range(0, len(files), self.images_per_pdf)]
        for group in groups:
            pdf_name = Path(group[0]).stem.rstrip('0123456789_ ')
            try:
                self._create_pdf(src_dir, group, dst_dir / f"{pdf_name}.pdf")
            except Exception:
                pass

    def _create_pdf(self, root_dir: Path, image_names: list, output_path: Path):
        pdf = FPDF()
        for img_name in image_names:
            pdf.add_page()
            pdf.image(str(root_dir / img_name), 0, 0, 210, 297)
        pdf.output(str(output_path), "F")