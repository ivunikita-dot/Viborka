#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Упрощённая копия скрипта для сборки в репозитории. Убедитесь, что это та же версия,
которую вы используете. В противном случае замените на свою последнюю версию.
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


def scan_file_openpyxl_horizontal(path: Path, target_set, max_seq_length=3, min_parts=2):
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
    parser = argparse.ArgumentParser(description='Ищет ФИО среди xlsx-файлов')
    parser.add_argument('--mode', choices=['auto', 'three', 'single'], default='auto')
    parser.add_argument('--source', '-s', default=None, help='Путь к исходному xlsx с ФИО')
    parser.add_argument('--folder', '-d', default=None, help='Корневая папка для поиска')
    parser.add_argument('--cols', '-c', nargs=3, metavar=('Фамилия', 'Имя', 'Отчество'), help='Имена трёх колонок (для three)')
    parser.add_argument('--col-name', default=None, help='Имя колонки с ФИО (для single)')
    parser.add_argument('--col-index', type=int, default=None, help='Индекс колонки (0-based) с ФИО (для single)')
    parser.add_argument('--no-recursive', dest='recursive', action='store_false', help='Не искать рекурсивно')
    parser.add_argument('--output', '-o', default='результат_фио.xlsx', help='Имя выходного xlsx файла')
    parser.add_argument('--gui', action='store_true', help='Открыть GUI-диалоги для выбора исходного файла и корневой папки')
    parser.add_argument('--max-seq-length', type=int, default=3, help='Максимальная длина последовательности соседних ячеек')
    parser.add_argument('--min-parts', type=int, default=2, help='Минимальное число слов в объединённой строке')

    args = parser.parse_args()

    use_gui = args.gui or (not any([args.source, args.folder, args.cols, args.col_name, args.col_index]))

    if use_gui:
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
            src = filedialog.askopenfilename(title='Выберите исходный xlsx файл', filetypes=[('Excel files','*.xlsx')])
            if not src:
                print('Исходный файл не выбран. Выход.')
                return
            source_path = Path(src)
            d = filedialog.askdirectory(title='Выберите корневую папку для поиска')
            if not d:
                print('Папка не выбрана. Выход.')
                return
            search_folder = Path(d)
        except Exception:
            print('GUI недоступно. Запустите без --gui или используйте CLI.')
            return
    else:
        if args.source:
            source_path = Path(args.source)
        else:
            source_path = Path(input('Введите путь к исходному xlsx: ').strip())
        if args.folder:
            search_folder = Path(args.folder)
        else:
            search_folder = Path(input('Введите путь к папке для поиска: ').strip())

    print('Читаю исходный файл:', source_path)
    df_src, header_present = read_source_with_better_header_detection(source_path)

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

    entries = []
    if mode == 'three':
        cols = args.cols if args.cols else list(df_src.columns[:3])
        for idx, row in df_src.iterrows():
            vals = []
            for c in cols:
                v = row.get(c, '') if c in df_src.columns else ''
                if pd.isna(v): v = ''
                vals.append(str(v).strip())
            full = ' '.join([p for p in vals if p])
            norm = normalize_text(full)
            entries.append((vals, norm))
        out_columns = list(cols) + ['Найдено в файлах'] if header_present else ['Фамилия','Имя','Отчество','Найдено в файлах']
    else:
        chosen_col_name = None
        if args.col_index is not None:
            chosen_col_name = df_src.columns[args.col_index]
        elif args.col_name:
            chosen_col_name = args.col_name
        else:
            hdrs = [str(h).lower() for h in df_src.columns]
            found = None
            for i,h in enumerate(hdrs):
                if 'фио' in h or 'fio' in h:
                    found = i; break
            chosen_col_name = df_src.columns[found] if found is not None else df_src.columns[0]
        for idx, row in df_src.iterrows():
            v = row.get(chosen_col_name, '')
            if pd.isna(v): v = ''
            s = str(v).strip()
            norm = normalize_text(s)
            entries.append((s, norm))
        out_columns = [chosen_col_name, 'Найдено в файлах'] if header_present else ['ФИО','Найдено в файлах']

    target_set = set([e[1] for e in entries if e[1]])
    if not target_set:
        print('Нет нормализованных ФИО для поиска — проверьте исходный файл.')
        return

    exclude = [source_path.resolve(), Path(args.output).resolve()]
    files = []
    for p in search_folder.rglob('*.xlsx'):
        try: rp = str(p.resolve())
        except: rp = str(p)
        if rp in [str(x) for x in exclude]: continue
        files.append(p)

    found_map = {k:set() for k in target_set}
    for p in files:
        matches = scan_file_openpyxl_horizontal(p, target_set, max_seq_length=args.max_seq_length, min_parts=args.min_parts)
        if matches:
            for m in matches:
                found_map[m].add(str(p))

    rows = []
    if mode == 'three':
        for vals,norm in entries:
            paths = sorted(found_map.get(norm, []))
            display = '; '.join(paths) if paths else 'ФИО не найдено'
            rows.append(vals + [display])
        out_df = pd.DataFrame(rows, columns=out_columns)
    else:
        for s,norm in entries:
            paths = sorted(found_map.get(norm, []))
            display = '; '.join(paths) if paths else 'ФИО не найдено'
            rows.append([s, display])
        out_df = pd.DataFrame(rows, columns=out_columns)

    try:
        out_df.to_excel(Path(args.output), index=False)
    except Exception as e:
        print('Ошибка при записи выходного файла:', e)
        return

    print('Готово. Результат записан в', Path(args.output))

if __name__ == '__main__':
    main()
