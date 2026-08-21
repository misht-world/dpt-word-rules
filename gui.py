# -*- coding: utf-8 -*-
"""
GUI-обёртка над пайплайнами нормализации .docx (ПЗ от разных Исполнителей).

Задача: привести каждый Word-документ к единому стилю. Ничего не склеивает —
каждый файл обрабатывается независимо.

Ничего не переизобретает: вызывает уже проверенные скрипты как есть, через
subprocess (их CLI оттестирован на реальных документах):
    - normalize_structure.py — стили, списки, таблицы, поля, единая нумерация заголовков
    - apply_docx.py          — типографика (неразрывные пробелы, дефисы, единицы)

Порядок при обоих включённых тумблерах: СТРУКТУРА -> ТИПОГРАФИКА.
Причина: структура сначала "съедает" литеральные маркеры списков ("- ", "N) ")
и номера заголовков, а типографика уже потом чистит текст. Если наоборот —
типографика может превратить ведущий "- " в тире и сломать распознавание списков.

Зависимости: только стандартная библиотека (tkinter входит в Python).
Запуск:  python gui.py
"""
import os
import sys
import queue
import threading
import subprocess
import tempfile
import traceback

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
STRUCT_SCRIPT = os.path.join(SCRIPT_DIR, 'normalize_structure.py')
TYPO_SCRIPT = os.path.join(SCRIPT_DIR, 'apply_docx.py')

SUFFIX = '_normalized'          # суффикс выходного файла в режиме "копия рядом"
DOCX_EXT = '.docx'


# ---------------------------------------------------------------------------
# Сбор файлов
# ---------------------------------------------------------------------------

def is_processable_docx(path):
    """.docx, не временный файл Word (~$...), не наш собственный результат."""
    name = os.path.basename(path)
    if not name.lower().endswith(DOCX_EXT):
        return False
    if name.startswith('~$'):
        return False
    stem = name[:-len(DOCX_EXT)]
    if stem.endswith(SUFFIX):     # чтобы повторный прогон не каскадировал
        return False
    return True


def collect_docx(paths):
    """Разворачивает список путей (файлы и папки) в плоский список .docx.
    Папки обходятся рекурсивно (вместе со вложенными). Дубликаты убираются."""
    found = []
    seen = set()

    def add(p):
        ap = os.path.abspath(p)
        key = os.path.normcase(ap)
        if key not in seen and is_processable_docx(ap):
            seen.add(key)
            found.append(ap)

    for p in paths:
        if os.path.isdir(p):
            for dirpath, dirnames, filenames in os.walk(p):
                for fn in filenames:
                    add(os.path.join(dirpath, fn))
        elif os.path.isfile(p):
            add(p)
    return found


def output_path_for(input_path, in_place):
    if in_place:
        return input_path
    root, ext = os.path.splitext(input_path)
    return root + SUFFIX + ext


# ---------------------------------------------------------------------------
# Запуск дочернего скрипта
# ---------------------------------------------------------------------------

def _child_env():
    """UTF-8 у дочернего процесса, чтобы русский текст в stdout не бился на Windows."""
    env = dict(os.environ)
    env['PYTHONUTF8'] = '1'
    env['PYTHONIOENCODING'] = 'utf-8'
    return env


def run_script(script, args):
    """Возвращает (returncode, output_text). Ничего не бросает."""
    cmd = [sys.executable, script] + args
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            env=_child_env(),
            encoding='utf-8',
            errors='replace',
        )
        out = (proc.stdout or '') + (proc.stderr or '')
        return proc.returncode, out.strip()
    except Exception:
        return -1, traceback.format_exc()


def process_one(input_path, do_struct, do_typo, in_place, make_report, log):
    """Обрабатывает один файл выбранными этапами. Возвращает True при успехе.
    log(str) — колбэк для вывода в интерфейс."""
    final_out = output_path_for(input_path, in_place)
    out_root, _ = os.path.splitext(final_out)

    # Промежуточный файл нужен только когда включены ОБА этапа
    tmp_dir = None
    stages = []
    if do_struct and do_typo:
        tmp_dir = tempfile.mkdtemp(prefix='dpt_gui_')
        mid = os.path.join(tmp_dir, 'stage_struct.docx')
        stages.append(('Структура', STRUCT_SCRIPT, input_path, mid,
                       out_root + '.structure.report.txt'))
        stages.append(('Типографика', TYPO_SCRIPT, mid, final_out,
                       out_root + '.typo.report.txt'))
    elif do_struct:
        stages.append(('Структура', STRUCT_SCRIPT, input_path, final_out,
                       out_root + '.structure.report.txt'))
    elif do_typo:
        stages.append(('Типографика', TYPO_SCRIPT, input_path, final_out,
                       out_root + '.typo.report.txt'))

    ok = True
    try:
        for label, script, src, dst, report in stages:
            args = [src, dst]
            if make_report:
                args += ['--report', report]
            code, out = run_script(script, args)
            for line in out.splitlines():
                log(f'      {line}')
            if code != 0:
                log(f'   [ОШИБКА] этап "{label}", код {code}')
                ok = False
                break
    finally:
        if tmp_dir:
            import shutil
            shutil.rmtree(tmp_dir, ignore_errors=True)
    return ok


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class App:
    def __init__(self, root):
        self.root = root
        root.title('ДПТ — нормализация Word (ПЗ)')
        root.geometry('760x620')
        root.minsize(640, 520)

        self.paths = []                 # выбранные файлы/папки (как выбрал пользователь)
        self.msg_queue = queue.Queue()  # сообщения из рабочего потока в UI
        self.worker = None

        self._build_ui()
        self.root.after(100, self._drain_queue)

    def _build_ui(self):
        pad = {'padx': 8, 'pady': 4}

        # --- Источник ---
        src = ttk.LabelFrame(self.root, text='1. Что обрабатываем')
        src.pack(fill='x', **pad)

        btns = ttk.Frame(src)
        btns.pack(fill='x', padx=6, pady=6)
        ttk.Button(btns, text='Добавить файлы…', command=self.add_files).pack(side='left')
        ttk.Button(btns, text='Добавить папку…', command=self.add_folder).pack(side='left', padx=6)
        ttk.Button(btns, text='Убрать выбранное', command=self.remove_selected).pack(side='left')
        ttk.Button(btns, text='Очистить всё', command=self.clear_paths).pack(side='left', padx=6)

        self.path_list = tk.Listbox(src, height=6, selectmode='extended')
        self.path_list.pack(fill='x', padx=6, pady=(0, 4))
        ttk.Label(src, text='Папки обходятся рекурсивно (вместе со вложенными). '
                            'Временные ~$-файлы и уже готовые *_normalized пропускаются.',
                  foreground='#666').pack(anchor='w', padx=6, pady=(0, 6))

        # --- Что применять ---
        what = ttk.LabelFrame(self.root, text='2. Что применять')
        what.pack(fill='x', **pad)
        self.var_struct = tk.BooleanVar(value=True)
        self.var_typo = tk.BooleanVar(value=True)
        ttk.Checkbutton(what, text='Структура  (стили, списки, таблицы, поля, единая нумерация заголовков)',
                        variable=self.var_struct).pack(anchor='w', padx=8, pady=2)
        ttk.Checkbutton(what, text='Типографика  (неразрывные пробелы, дефисы, единицы измерения)',
                        variable=self.var_typo).pack(anchor='w', padx=8, pady=2)
        ttk.Label(what, text='При обоих включённых порядок: структура → типографика.',
                  foreground='#666').pack(anchor='w', padx=8, pady=(0, 6))

        # --- Куда ---
        out = ttk.LabelFrame(self.root, text='3. Результат')
        out.pack(fill='x', **pad)
        self.var_inplace = tk.BooleanVar(value=False)
        ttk.Radiobutton(out, text=f'Копия рядом  (файл{SUFFIX}.docx) — оригинал не трогаем',
                        variable=self.var_inplace, value=False).pack(anchor='w', padx=8, pady=2)
        ttk.Radiobutton(out, text='На месте (перезапись оригинала) — только с бэкапом!',
                        variable=self.var_inplace, value=True).pack(anchor='w', padx=8, pady=2)
        self.var_report = tk.BooleanVar(value=True)
        ttk.Checkbutton(out, text='Писать текстовый отчёт (report.txt рядом с результатом)',
                        variable=self.var_report).pack(anchor='w', padx=8, pady=(6, 6))

        # --- Запуск ---
        run = ttk.Frame(self.root)
        run.pack(fill='x', **pad)
        self.run_btn = ttk.Button(run, text='▶  Запустить', command=self.start)
        self.run_btn.pack(side='left')
        self.progress = ttk.Progressbar(run, mode='determinate')
        self.progress.pack(side='left', fill='x', expand=True, padx=8)

        # --- Лог ---
        logf = ttk.LabelFrame(self.root, text='Журнал')
        logf.pack(fill='both', expand=True, **pad)
        self.log_text = tk.Text(logf, height=10, wrap='word', state='disabled')
        self.log_text.pack(side='left', fill='both', expand=True, padx=(6, 0), pady=6)
        sb = ttk.Scrollbar(logf, command=self.log_text.yview)
        sb.pack(side='right', fill='y', pady=6)
        self.log_text['yscrollcommand'] = sb.set

    # --- работа со списком путей ---
    def _refresh_list(self):
        self.path_list.delete(0, 'end')
        for p in self.paths:
            tag = '[папка] ' if os.path.isdir(p) else ''
            self.path_list.insert('end', tag + p)

    def add_files(self):
        files = filedialog.askopenfilenames(
            title='Выберите .docx',
            filetypes=[('Word', '*.docx'), ('Все файлы', '*.*')])
        for f in files:
            if f not in self.paths:
                self.paths.append(f)
        self._refresh_list()

    def add_folder(self):
        d = filedialog.askdirectory(title='Выберите папку (обрабатываются все .docx внутри)')
        if d and d not in self.paths:
            self.paths.append(d)
        self._refresh_list()

    def remove_selected(self):
        for i in reversed(self.path_list.curselection()):
            del self.paths[i]
        self._refresh_list()

    def clear_paths(self):
        self.paths = []
        self._refresh_list()

    # --- лог ---
    def log(self, msg):
        self.msg_queue.put(('log', msg))

    def _drain_queue(self):
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                if kind == 'log':
                    self.log_text['state'] = 'normal'
                    self.log_text.insert('end', payload + '\n')
                    self.log_text.see('end')
                    self.log_text['state'] = 'disabled'
                elif kind == 'progress':
                    done, total = payload
                    self.progress['maximum'] = total
                    self.progress['value'] = done
                elif kind == 'done':
                    self.run_btn['state'] = 'normal'
        except queue.Empty:
            pass
        self.root.after(100, self._drain_queue)

    # --- запуск ---
    def start(self):
        if self.worker and self.worker.is_alive():
            return
        do_struct = self.var_struct.get()
        do_typo = self.var_typo.get()
        if not (do_struct or do_typo):
            messagebox.showwarning('Нечего делать', 'Включите хотя бы один тумблер: структура или типографика.')
            return
        if not self.paths:
            messagebox.showwarning('Нет входных данных', 'Добавьте файлы или папку.')
            return

        files = collect_docx(self.paths)
        if not files:
            messagebox.showinfo('Пусто', 'В выбранных путях не найдено подходящих .docx.')
            return

        if self.var_inplace.get():
            if not messagebox.askyesno(
                    'Перезапись оригиналов',
                    f'Будут ПЕРЕЗАПИСАНЫ {len(files)} исходных файлов.\n'
                    'Убедитесь, что есть бэкап. Продолжить?'):
                return

        self.run_btn['state'] = 'disabled'
        self.log_text['state'] = 'normal'
        self.log_text.delete('1.0', 'end')
        self.log_text['state'] = 'disabled'

        args = (files, do_struct, do_typo, self.var_inplace.get(), self.var_report.get())
        self.worker = threading.Thread(target=self._run_batch, args=args, daemon=True)
        self.worker.start()

    def _run_batch(self, files, do_struct, do_typo, in_place, make_report):
        total = len(files)
        stages = ' + '.join([s for s, on in (('структура', do_struct), ('типографика', do_typo)) if on])
        mode = 'перезапись на месте' if in_place else f'копия ({SUFFIX})'
        self.log(f'Файлов к обработке: {total}   |   этапы: {stages}   |   выход: {mode}')
        self.log('=' * 70)
        self.msg_queue.put(('progress', (0, total)))

        ok_count = 0
        err_count = 0
        for i, f in enumerate(files, 1):
            self.log(f'[{i}/{total}] {f}')
            try:
                ok = process_one(f, do_struct, do_typo, in_place, make_report, self.log)
            except Exception:
                self.log('      [ИСКЛЮЧЕНИЕ] ' + traceback.format_exc())
                ok = False
            if ok:
                ok_count += 1
                self.log('      ✓ готово: ' + output_path_for(f, in_place))
            else:
                err_count += 1
            self.msg_queue.put(('progress', (i, total)))

        self.log('=' * 70)
        self.log(f'Готово. Успешно: {ok_count}, с ошибками: {err_count}.')
        self.msg_queue.put(('done', None))


def main():
    for script in (STRUCT_SCRIPT, TYPO_SCRIPT):
        if not os.path.exists(script):
            print(f'Не найден скрипт: {script}', file=sys.stderr)
            sys.exit(1)
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()
