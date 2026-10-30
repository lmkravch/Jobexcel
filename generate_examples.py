"""Создание пяти небольших файлов с заведомо известными данными."""
from openpyxl import Workbook
from main import BASE_DIR, COLUMNS


def generate():
    folder = BASE_DIR / 'input'
    folder.mkdir(exist_ok=True)
    paths = [folder / f'example_{i}.xlsx' for i in range(1, 6)]
    if any(path.exists() for path in paths):
        raise SystemExit('Примеры уже существуют. Генератор не перезаписывает файлы.')
    anna = ['2026-10-01', 'Анна', '+7 999 123-45-67', 'Книга', 100.10]
    boris = ['2026-10-02', 'Борис', '89161234567', 'Ручка', 200]
    batches = [
        [anna, anna, boris, ['2026-10-03', 'Вера', None, 'Блокнот', 50]],
        [anna, ['2026-10-04', 'Анна', '89991234567', 'Тетрадь', 200.20],
         ['2026-10-05', 'Глеб', '9992223344', 'Книга', None]],
        [boris, ['2026-10-06', 'Анна', '9991234567', 'Альбом', 300.30],
         ['2026-10-07', 'Вера', '', 'Папка', 70]],
        [['2026-10-08', 'Борис', '+7 (916) 123-45-67', 'Альбом', 300],
         ['2026-10-09', 'Глеб', '89992223344', 'Папка', None],
         ['2026-10-10', 'Дарья', '9995556677', 'Карандаш', 0]],
        [['2026-10-11', 'Вера', '9994445566', 'Бумага', 75.50],
         ['2026-10-12', 'Елена', None, 'Ластик', None],
         ['2026-10-13', 'Дарья', '+79995556677', 'Пенал', 25]],
    ]
    for path, rows in zip(paths, batches):
        book = Workbook()
        book.active.append(COLUMNS)
        for row in rows:
            book.active.append(row)
        book.save(path)
        book.close()
    print('Создано 5 примеров в input.')


if __name__ == '__main__':
    generate()
