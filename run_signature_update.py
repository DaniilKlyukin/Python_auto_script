
import os
from services.signature_processor import process_docx_signatures


def main():
    print("=== Утилита замены ФИО в зонах подписей (РПД) ===")

    dir_path = input("Введите путь к папке с файлами .docx: ").strip().strip('"')
    if not os.path.isdir(dir_path):
        print("Ошибка: Путь не найден.")
        return

    old_fio = input("Введите ФИО, которое нужно заменить: ").strip()
    new_fio = input("Введите новое ФИО: ").strip()

    if not old_fio or not new_fio:
        print("Ошибка: ФИО не могут быть пустыми.")
        return

    files = [f for f in os.listdir(dir_path) if f.lower().endswith('.docx') and not f.startswith('~$')]

    if not files:
        print("В папке не найдено файлов .docx")
        return

    stats = {"ok": 0, "skip": 0, "err": 0}

    for filename in files:
        full_path = os.path.join(dir_path, filename)
        success, message = process_docx_signatures(full_path, old_fio, new_fio)

        if success:
            print(f"[OK] {filename}")
            stats["ok"] += 1
        else:
            if "не найдено" in message:
                print(f"[SKIP] {filename} (ФИО не обнаружено в зонах подписи)")
                stats["skip"] += 1
            else:
                print(f"[ERR] {filename}: {message}")
                stats["err"] += 1

    print(f"\nЗавершено!")
    print(f"Обновлено: {stats['ok']}")
    print(f"Пропущено: {stats['skip']}")
    print(f"Ошибок: {stats['err']}")


if __name__ == "__main__":
    main()