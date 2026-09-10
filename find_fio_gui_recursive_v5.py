#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
find_fio_gui_recursive_v5.py

Версия: оптимизированная для случая, когда ФИО в целевых файлах
встречается только в одной строке (горизонтально). Убрано вертикальное
сканирование — стало быстрее и однозначнее.

Новые опции:
  --max-seq-length N   # максимальная длина последовательности соседних ячеек (по умолчанию 3)
  --min-parts M        # минимальное число слов в объединённой строке, чтобы считать возможным ФИО (по умолчанию 2)

Остальное поведение как в предыдущих версиях (detekciya header, GUI выбор файла/папки,
рекурсивный поиск по подпапкам и т.д.).
"""

import argparse
import re
import sys
from pathlib import Path

import pandas as pd

try:
    from openpyxl import load_workbook
except Exception:
    load_workbook = None


def normalize_text(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    s = re.sub(r'\s+', ' ', s)
    s = re.sub(r'[^0-9A-Za-zА-Яа-яЁё\s\-]', '', s)
    return s.lower()


def looks_like_fio_text(s: str) -> bool:
    if s is None:
        return False
    s = str(s).strip()
    if not s:
        return False
    parts = s.split()
    if len(parts) < 2 or len(parts) > 4:
        return False
    for p in parts:
        if not re.match(r"^[A-Za-zА-Яа-яЁё\-]+$", p):
            return False
    return True


# GUI helpers
def choose_file_via_gui(title: str = 'Выберите исходный xlsx файл') -> str:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        return ''
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    file = filedialog.askopenfilename(title=title, filetypes=[('Excel files', '*.xlsx')])
    root.destroy()
    return file


def choose_folder_via_gui(title: str = 'Выберите корневую папку для поиска') -> str:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        return ''
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    d = filedialog.askdirectory(title=title)
    root.destroy()
    return d


# Console helpers
def choose_file_console(prompt: str = 'Введите путь к исходному xlsx файлу: ') -> str:
    s = input(prompt).strip()
    return s


def choose_folder_console(prompt: str = 'Введите путь к корневой папке для поиска: ') -> str:
    s = input(prompt).strip()
    return s


def collect_xlsx_files(folder: Path, recursive: bool, exclude_paths=None):
    exclude_paths = set([str(p) for p in (exclude_paths or [])])
    files = []
    if recursive:
        it = folder.rglob('*.xlsx')
    else:
        it = folder.glob('*.xlsx')
    for p in it:
        try:
            rp = str(p.resolve())
        except Exception:
            rp = str(p)
        if rp in exclude_paths:
            continue
        files.append(p)
    return files


def scan_file_openpyxl_horizontal(path: Path, target_set, max_seq_length=3, min_parts=2):
    """Сканирует .xlsx файл, ищет target_set только в горизонтальных последовательностях
    соседних ячеек в одной строке. Возвращает set найденных нормализованных ключей.

    Правила:
    - рассматриваем последовательности длиной 1..max_seq_length
    - объединяем непустые части через пробел
    - перед сравнением проверяем, что объединённая строка содержит >= min_parts слов
    """
    found = set()
    if load_workbook is None:
        raise RuntimeError('openpyxl не установлен. Установите openpyxl: pip install openpyxl')
    try:
        wb = load_workbook(filename=str(path), read_only=True, data_only=True)
    except Exception:
        return found
    try:
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue
            nrows = len(rows)
            ncols = max((len(r) for r in rows), default=0)
            for r_idx, row in enumerate(rows):
                # normalize row length
                row_vals = [None] * ncols
                for c_idx, v in enumerate(row):
                    row_vals[c_idx] = v
                for c_start in range(ncols):
                    parts = []
                    for length in range(1, max_seq_length+1):
                        c = c_start + length - 1
                        if c >= ncols:
                            break
                        val = row_vals[c]
                        if val is None:
                            parts.append('')
                        else:
                            parts.append(str(val).strip())
                        comb = ' '.join([p for p in parts if p])
                        if not comb:
                            continue
                        # require minimum words (parts by whitespace) to reduce false positives
                        if len(comb.split()) < min_parts:
                            continue
                        n = normalize_text(comb)
                        if n in target_set:
                            found.add(n)
    finally:
        try:
            wb.close()
        except Exception:
            pass
    return found


def read_source_with_better_header_detection(path: Path):
    try:
        df = pd.read_excel(path, dtype=str)
    except Exception as e:
        raise

    need_reread = False
    if df.shape[0] == 0:
        need_reread = True

    if not need_reread:
        for col in df.columns:
            if looks_like_fio_text(str(col)):
                need_reread = True
                break

    if need_reread:
        try:
            df2 = pd.read_excel(path, header=None, dtype=str)
        except Exception:
            return df, True
        ncols = df2.shape[1]
        cols = [f'col{i+1}' for i in range(ncols)]
        df2.columns = cols
        print('Внимание: файл был прочитан повторно с header=None — первая строка будет считаться данными.')
        return df2, False

    return df, True


def main():
    parser = argparse.ArgumentParser(description='Ищет ФИО среди xlsx-файлов в выбранной корневой папке и подпапках (по умолчанию).')
    parser.add_argument('--mode', choices=['auto', 'three', 'single'], default='auto')

    parser.add_argument('--source', '-s', default=None, help='Путь к исходному xlsx с ФИО')
    parser.add_argument('--folder', '-d', default=None, help='Корневая папка для поиска (будет выполнен рекурсивный поиск по подпапкам по умолчанию)')

    parser.add_argument('--cols', '-c', nargs=3, metavar=('Фамилия', 'Имя', 'Отчество'), help='Имена трёх колонок (для three)')
    parser.add_argument('--col-name', default=None, help='Имя колонки с ФИО (для single)')
    parser.add_argument('--col-index', type=int, default=None, help='Индекс колонки (0-based) с ФИО (для single)')

    parser.add_argument('--no-recursive', dest='recursive', action='store_false', help='Не искать рекурсивно (только в корневой папке)')
    parser.add_argument('--output', '-o', default='результат_фио.xlsx', help='Имя выходного xlsx файла')

    parser.add_argument('--gui', action='store_true', help='Открыть GUI-диалоги для выбора исходного файла и корневой папки')
    parser.add_argument('--max-seq-length', type=int, default=3, help='Максимальная длина последовательности соседних ячеек в строке (по умолчанию 3)')
    parser.add_argument('--min-parts', type=int, default=2, help='Минимальное число слов в объединённой строке для считания возможным ФИО (по умолчанию 2)')

    args = parser.parse_args()

    use_gui = args.gui or (not any([args.source, args.folder, args.cols, args.col_name, args.col_index]))

    # choose source
    if use_gui:
        src = choose_file_via_gui()
        if not src:
            print('Исходный файл не выбран через GUI. Выход.')
            return
        source_path = Path(src)
    else:
        if args.source:
            source_path = Path(args.source)
        else:
            sp = choose_file_console()
            if not sp:
                print('Исходный файл не задан. Выход.')
                return
            source_path = Path(sp)

    if not source_path.exists():
        print(f'Исходный файл не найден: {source_path}', file=sys.stderr)
        return

    # choose root folder
    if use_gui:
        folder_selected = choose_folder_via_gui()
        if not folder_selected:
            print('Папка не выбрана через GUI. Выход.')
            return
        search_folder = Path(folder_selected)
    else:
        if args.folder:
            search_folder = Path(args.folder)
        else:
            sf = choose_folder_console()
            if not sf:
                print('Папка для поиска не задана. Выход.')
                return
            search_folder = Path(sf)

    if not search_folder.exists() or not search_folder.is_dir():
        print(f'Папка для поиска не найдена: {search_folder}', file=sys.stderr)
        return

    print('Читаю исходный файл:', source_path)
    try:
        df_src, header_present = read_source_with_better_header_detection(source_path)
    except Exception as e:
        print('Ошибка при чтении исходного файла через pandas:', e, file=sys.stderr)
        return

    mode = args.mode
    if mode == 'auto':
        hdrs = [str(h).lower() for h in df_src.columns]
        if any('фио' in h or 'fio' in h for h in hdrs):
            mode = 'single'
        elif df_src.shape[1] >= 3:
            mode = 'three'
        else:
            mode = 'single'
        print('Auto mode detected:', mode)
    else:
        print('Mode:', mode)

    entries = []
    if mode == 'three':
        if args.cols:
            cols = args.cols
        else:
            cols = list(df_src.columns[:3])
        print('Используем колонки для ФИО:', cols)
        for idx, row in df_src.iterrows():
            vals = []
            for c in cols:
                v = row.get(c, '') if c in df_src.columns else ''
                if pd.isna(v):
                    v = ''
                vals.append(str(v).strip())
            full = ' '.join([p for p in vals if p])
            norm = normalize_text(full)
            entries.append((vals, norm))
        if header_present:
            out_columns = list(cols) + ['Найдено в файлах']
        else:
            out_columns = ['Фамилия', 'Имя', 'Отчество', 'Найдено в файлах']
    else:
        chosen_idx = None
        chosen_col_name = None
        if args.col_index is not None:
            chosen_idx = args.col_index
            if chosen_idx < 0 or chosen_idx >= df_src.shape[1]:
                print('col_index вне диапазона', file=sys.stderr)
                return
            chosen_col_name = df_src.columns[chosen_idx]
        elif args.col_name:
            if args.col_name not in df_src.columns:
                print(f"Колонки с именем '{args.col_name}' нет в исходном файле. Доступные: {list(df_src.columns)}", file=sys.stderr)
                return
            chosen_col_name = args.col_name
        else:
            hdrs = [str(h).lower() for h in df_src.columns]
            found = None
            for i, h in enumerate(hdrs):
                if 'фио' in h or 'fio' in h:
                    found = i
                    break
            if found is not None:
                chosen_idx = found
                chosen_col_name = df_src.columns[chosen_idx]
            else:
                chosen_idx = 0
                chosen_col_name = df_src.columns[0]
        print('Используем колонку для ФИО:', chosen_col_name)
        for idx, row in df_src.iterrows():
            v = row.get(chosen_col_name, '')
            if pd.isna(v):
                v = ''
            s = str(v).strip()
            norm = normalize_text(s)
            entries.append((s, norm))
        if header_present:
            out_columns = [chosen_col_name, 'Найдено в файлах']
        else:
            out_columns = ['ФИО', 'Найдено в файлах']

    target_set = set([e[1] for e in entries if e[1]])
    if not target_set:
        print('Нет нормализованных ФИО для поиска — проверьте исходный файл.', file=sys.stderr)
        return

    print(f'Итого уникальных для поиска: {len(target_set)}')

    # collect files
    exclude = [source_path.resolve(), Path(args.output).resolve()]
    files = collect_xlsx_files(search_folder, args.recursive, exclude_paths=exclude)
    print(f'Будет проверено {len(files)} xlsx файлов в {search_folder} (recursive={args.recursive})')

    found_map = {k: set() for k in target_set}
    for i, p in enumerate(files, 1):
        print(f'[{i}/{len(files)}] Проверяю: {p}')
        try:
            matches = scan_file_openpyxl_horizontal(p, target_set, max_seq_length=args.max_seq_length, min_parts=args.min_parts)
        except RuntimeError as e:
            print(e, file=sys.stderr)
            return
        if matches:
            for m in matches:
                found_map[m].add(str(p))

    # prepare output
    if mode == 'three':
        rows = []
        for vals, norm in entries:
            paths = sorted(found_map.get(norm, []))
            rows.append(vals + ['; '.join(paths)])
        out_df = pd.DataFrame(rows, columns=out_columns)
    else:
        rows = []
        for s, norm in entries:
            paths = sorted(found_map.get(norm, []))
            rows.append([s, '; '.join(paths)])
        out_df = pd.DataFrame(rows, columns=out_columns)

    out_path = Path(args.output)
    try:
        out_df.to_excel(out_path, index=False)
    except Exception as e:
        print('Ошибка при записи выходного файла через pandas:', e, file=sys.stderr)
        return

    # If tkinter available, show a messagebox with path
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        messagebox.showinfo('Готово', f'Результат записан в:\n{out_path.resolve()}')
        root.destroy()
    except Exception:
        print('Готово. Результат записан в', out_path)


if __name__ == '__main__':
    main()
