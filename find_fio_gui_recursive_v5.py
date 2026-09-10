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

   def column_is_name_token(v):
    if v is None:
        return False
    s = str(v).strip()
    if not s:
        return False
    # одно «слово» (только буквы/дефис)
    return bool(re.match(r'^[A-Za-zА-Яа-яЁё\-]+$', s))

def column_looks_like_fullname(v):
    if v is None:
        return False
    s = str(v).strip()
    if not s:
        return False
    # 2..4 слова, каждое — буквы/дефис
    parts = s.split()
    if len(parts) < 2 or len(parts) > 4:
        return False
    for p in parts:
        if not re.match(r'^[A-Za-zА-Яа-яЁё\-]+$', p):
            return False
    return True

def detect_mode_and_columns(df, args):
    # Если пользователь явно указал режим — используем его
    if args.mode != 'auto':
        return args.mode, None

    nrows = max(1, len(df))
    # 1) Проверим: есть ли колонка с большим количеством полных ФИО (single)
    single_counts = []
    for col in df.columns:
        cnt = 0
        for v in df[col].head(200):  # ограничиваемся первыми 200 строк для скорости
            if column_looks_like_fullname(v):
                cnt += 1
        single_counts.append(cnt)
    best_single = max(single_counts) if single_counts else 0
    # доля строк где колонка выглядит как ФИО
    single_frac = best_single / min(nrows, 200)

    # 2) Проверим: есть ли три соседних колонки, где каждое поле похоже на отдельное имя
    three_best = (None, 0)  # (start_col_index, count)
    cols = list(df.columns)
    for i in range(max(0, len(cols) - 3 + 1)):
        cnt = 0
        for r in range(min(nrows, 200)):
            try:
                v1 = df.iloc[r, i]
                v2 = df.iloc[r, i+1]
                v3 = df.iloc[r, i+2]
            except Exception:
                continue
            if column_is_name_token(v1) and column_is_name_token(v2) and column_is_name_token(v3):
                cnt += 1
        if cnt > three_best[1]:
            three_best = (i, cnt)
    three_frac = three_best[1] / min(nrows, 200)

    # Правила выбора (порог можно менять)
    # Если доля full-name в одной колонке > 0.25 -> single
    # Иначе если доля трёх колонок > 0.25 -> three
    # Иначе fallback: если df.columns >=3 -> three, иначе single
    if single_frac >= 0.25:
        return 'single', None
    if three_frac >= 0.25:
        start = three_best[0]
        # вернуть имена колонок (если имена существуют) или индексы
        if start is not None:
            return 'three', cols[start:start+3]
    if df.shape[1] >= 3:
        return 'three', list(df.columns[:3])
    return 'single', None

# Применяем детектор
mode_detected, detected_cols = detect_mode_and_columns(df_src, args)
mode = mode_detected
print('Detected mode:', mode, 'detected_cols:', detected_cols)

# Формируем entries одинаково как раньше, но с поддержкой detected_cols
entries = []
if mode == 'three':
    if args.cols:
        cols = args.cols
    elif detected_cols:
        cols = detected_cols
    else:
        cols = list(df_src.columns[:3])
    for idx, row in df_src.iterrows():
        vals = []
        for c in cols:
            # поддерживаем как именованные колонки, так и индексы
            if isinstance(c, int):
                try:
                    v = row.iloc[c]
                except Exception:
                    v = ''
            else:
                v = row.get(c, '') if c in df_src.columns else ''
            if pd.isna(v): v = ''
            vals.append(str(v).strip())
        full = ' '.join([p for p in vals if p])
        norm = normalize_text(full)
        entries.append((vals, norm))
    out_columns = list(cols) + ['Найдено в файлах'] if df_src.shape[0] and any(isinstance(h, str) for h in df_src.columns) else ['Фамилия','Имя','Отчество','Найдено в файлах']
else:
    # single
    chosen_col_name = None
    if args.col_index is not None:
        chosen_col_name = df_src.columns[args.col_index]
    elif args.col_name:
        chosen_col_name = args.col_name
    else:
        chosen_col_name = None  # пусть детектор выберет
    if chosen_col_name is None:
        # найдём наиболее вероятную колонку с ФИО (по column_looks_like_fullname)
        best_col = None
        best_cnt = -1
        for col in df_src.columns:
            cnt = sum(1 for v in df_src[col].head(200) if column_looks_like_fullname(v))
            if cnt > best_cnt:
                best_cnt = cnt; best_col = col
        chosen_col_name = best_col if best_col is not None else df_src.columns[0]
    for idx, row in df_src.iterrows():
        v = row.get(chosen_col_name, '') if chosen_col_name in df_src.columns else ''
        if pd.isna(v): v = ''
        s = str(v).strip()
        norm = normalize_text(s)
        entries.append((s, norm))
    out_columns = [chosen_col_name, 'Найдено в файлах'] if df_src.shape[0] and any(isinstance(h, str) for h in df_src.columns) else ['ФИО','Найдено в файлах']

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
