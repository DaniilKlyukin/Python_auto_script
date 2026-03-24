import os

class FileCleaner:
    @staticmethod
    def delete(file_path: str) -> bool:
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                return True
            return False
        except Exception:
            return False

    @staticmethod
    def cleanup_folder(folder_path: str, extensions=('.pdf', '.jpg', '.jpeg', '.png')):
        count = 0
        for root, _, files in os.walk(folder_path):
            for filename in files:
                if filename.lower().endswith(extensions):
                    if FileCleaner.delete(os.path.join(root, filename)):
                        count += 1
        return count