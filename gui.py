# -*- coding: utf-8 -*-
"""
GUI для работы с .docx (ПЗ от разных Исполнителей). Две вкладки:

1. «Типографика» — приводит каждый документ к единому оформлению. Обёртка над
   проверенным apply_docx.py (вызов через subprocess, CLI оттестирован).
   «Структура» (normalize_structure.py) временно отключена — ломает документы,
   дорабатывается отдельно; плечо do_struct в process_one сохранено.

2. «Поиск и замена» — поиск слова/словосочетания во всех выбранных .docx (по
   каталогам рекурсивно), просмотр найденного с контекстом, выбор отдельных
   вхождений и замена. Логика — в find_replace.py (уровень текста абзаца, чтобы
   находить фразы, разорванные на run'ы; замена через redistribute без потери
   форматирования).

Выбор источника (файлы/папки) и режим сохранения (копия рядом / на месте) —
общие для обеих вкладок.

Зависимости: только стандартная библиотека (tkinter входит в Python).
Запуск:  python gui.py
"""
import os
import sys
import json
import queue
import threading
import subprocess
import tempfile
import traceback

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

import find_replace as FR
import doc_convert as DC

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
STRUCT_SCRIPT = os.path.join(SCRIPT_DIR, 'normalize_structure.py')
TYPO_SCRIPT = os.path.join(SCRIPT_DIR, 'apply_docx.py')

# Сохранённые наборы путей и последний выбор — вне репозитория (в профиле
# пользователя), чтобы не коммитить и не терять между запусками.
CONFIG_DIR = os.path.join(os.environ.get('APPDATA') or os.path.expanduser('~'), 'dpt-word-rules')
PRESETS_PATH = os.path.join(CONFIG_DIR, 'presets.json')

SUFFIX = '_normalized'          # суффикс выходного файла в режиме "копия рядом"
DOCX_EXT = '.docx'

CHECKED = '☑'
UNCHECKED = '☐'
PARTIAL = '◪'


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


def collect_doc(paths):
    """Список файлов .doc (старый формат) в выбранных путях, рекурсивно.
    Их нельзя обрабатывать напрямую — сначала конвертируются в .docx (Word)."""
    found = []
    seen = set()

    def add(p):
        ap = os.path.abspath(p)
        key = os.path.normcase(ap)
        name = os.path.basename(ap)
        if key in seen or name.startswith('~$'):
            return
        if not name.lower().endswith('.doc'):
            return
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
# Запуск дочернего скрипта (типографика)
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
        root.title('ДПТ — обработка Word (ПЗ)')
        root.geometry('860x720')
        root.minsize(720, 600)

        self.paths = []                 # выбранные файлы/папки
        self.msg_queue = queue.Queue()  # сообщения из рабочих потоков в UI
        self.busy = False
        self.store = self._load_store()  # наборы путей + последний выбор

        # состояние вкладки поиска/замены
        self.fr_results = {}            # path -> [match dict]
        self.fr_item_meta = {}          # tree item id -> meta
        self.fr_query = ''
        self.fr_match_case = False
        self.fr_whole_word = True

        self._build_ui()
        # восстановить последний выбор путей
        self.paths = [p for p in self.store.get('last', []) if isinstance(p, str)]
        self._refresh_list()
        # Ctrl+C/V/X/A на русской раскладке (Ctrl шлёт кириллические клавиши,
        # для которых нет стандартных биндов копирования/вставки).
        self.root.bind_all('<Control-KeyPress>', self._ctrl_key)
        self.root.after(100, self._drain_queue)

    # ------------------------------------------------------------ хранилище
    def _load_store(self):
        try:
            with open(PRESETS_PATH, encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, dict):
                data.setdefault('presets', {})
                data.setdefault('last', [])
                return data
        except Exception:
            pass
        return {'presets': {}, 'last': []}

    def _save_store(self):
        try:
            os.makedirs(CONFIG_DIR, exist_ok=True)
            with open(PRESETS_PATH, 'w', encoding='utf-8') as f:
                json.dump(self.store, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _persist_last(self):
        self.store['last'] = list(self.paths)
        self._save_store()

    def _ctrl_key(self, event):
        """Копирование/вставка/вырезание/выделение при русской раскладке.
        По латинским c/v/x/a не вмешиваемся — там работает стандартный бинд Tk;
        реагируем только когда keysym не латинский (кириллица), опираясь на
        keycode физической клавиши (он от раскладки не зависит)."""
        if not (event.state & 0x4):        # не Control
            return
        if event.keysym.lower() in ('c', 'v', 'x', 'a'):
            return                          # латиница — стандартный бинд справится
        kc = event.keycode
        vev = {67: '<<Copy>>', 86: '<<Paste>>', 88: '<<Cut>>'}.get(kc)  # C/V/X
        if vev:
            try:
                event.widget.event_generate(vev)
            except Exception:
                pass
            return 'break'
        if kc == 65:                        # A — выделить всё
            w = event.widget
            try:
                if isinstance(w, (tk.Entry, ttk.Entry)):
                    w.select_range(0, 'end'); w.icursor('end')
                elif isinstance(w, tk.Text):
                    w.tag_add('sel', '1.0', 'end-1c')
            except Exception:
                pass
            return 'break'

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        pad = {'padx': 8, 'pady': 4}

        # --- Общий источник ---
        src = ttk.LabelFrame(self.root, text='Источник (общий для обеих вкладок)')
        src.pack(fill='x', **pad)
        btns = ttk.Frame(src)
        btns.pack(fill='x', padx=6, pady=6)
        ttk.Button(btns, text='Добавить файлы…', command=self.add_files).pack(side='left')
        ttk.Button(btns, text='Добавить папку…', command=self.add_folder).pack(side='left', padx=6)
        ttk.Button(btns, text='Убрать выбранное', command=self.remove_selected).pack(side='left')
        ttk.Button(btns, text='Очистить всё', command=self.clear_paths).pack(side='left', padx=6)
        self.path_list = tk.Listbox(src, height=5, selectmode='extended')
        self.path_list.pack(fill='x', padx=6, pady=(0, 4))
        ttk.Label(src, text='Папки обходятся рекурсивно (вместе со вложенными). '
                            'Временные ~$-файлы и уже готовые *_normalized пропускаются.',
                  foreground='#666').pack(anchor='w', padx=6, pady=(0, 4))

        prow = ttk.Frame(src)
        prow.pack(fill='x', padx=6, pady=(0, 6))
        ttk.Label(prow, text='Наборы:').pack(side='left')
        self.preset_combo = ttk.Combobox(prow, state='readonly', width=26)
        self.preset_combo.pack(side='left', padx=6)
        ttk.Button(prow, text='Загрузить', command=self.load_preset).pack(side='left')
        ttk.Button(prow, text='Сохранить как…', command=self.save_preset).pack(side='left', padx=6)
        ttk.Button(prow, text='Удалить', command=self.delete_preset).pack(side='left')
        self._refresh_presets()

        # --- Общий режим сохранения ---
        out = ttk.LabelFrame(self.root, text='Сохранение результата')
        out.pack(fill='x', **pad)
        self.var_inplace = tk.BooleanVar(value=True)
        ttk.Radiobutton(out, text=f'Копия рядом  (файл{SUFFIX}.docx) — оригинал не трогаем',
                        variable=self.var_inplace, value=False).pack(side='left', padx=8, pady=4)
        ttk.Radiobutton(out, text='На месте (перезапись) — только с бэкапом!',
                        variable=self.var_inplace, value=True).pack(side='left', padx=8, pady=4)

        # --- Вкладки ---
        nb = ttk.Notebook(self.root)
        nb.pack(fill='both', expand=True, **pad)
        self._build_typo_tab(nb)
        self._build_fr_tab(nb)

        # --- Общий прогресс ---
        self.progress = ttk.Progressbar(self.root, mode='determinate')
        self.progress.pack(fill='x', padx=8, pady=(0, 8))

    def _build_typo_tab(self, nb):
        tab = ttk.Frame(nb)
        nb.add(tab, text='  Типографика  ')

        what = ttk.LabelFrame(tab, text='Что применять')
        what.pack(fill='x', padx=6, pady=6)
        self.var_struct = tk.BooleanVar(value=False)   # структура пока выключена
        self.var_typo = tk.BooleanVar(value=True)
        ttk.Checkbutton(what, text='Типографика  (неразрывные пробелы, дефисы, единицы измерения)',
                        variable=self.var_typo).pack(anchor='w', padx=8, pady=2)
        ttk.Label(what, text='Структура (стили, списки, таблицы, заголовки) временно отключена — '
                            'ломает документы, дорабатывается отдельно.',
                  foreground='#a00').pack(anchor='w', padx=8, pady=(0, 4))
        self.var_report = tk.BooleanVar(value=False)
        ttk.Checkbutton(what, text='Писать текстовый отчёт (report.txt рядом с результатом)',
                        variable=self.var_report).pack(anchor='w', padx=8, pady=(0, 6))

        row = ttk.Frame(tab)
        row.pack(fill='x', padx=6, pady=4)
        self.typo_btn = ttk.Button(row, text='▶  Запустить типографику', command=self.start_typo)
        self.typo_btn.pack(side='left')

        logf = ttk.LabelFrame(tab, text='Журнал')
        logf.pack(fill='both', expand=True, padx=6, pady=6)
        self.log_text = tk.Text(logf, height=10, wrap='word', state='disabled')
        self.log_text.pack(side='left', fill='both', expand=True, padx=(6, 0), pady=6)
        sb = ttk.Scrollbar(logf, command=self.log_text.yview)
        sb.pack(side='right', fill='y', pady=6)
        self.log_text['yscrollcommand'] = sb.set

    def _build_fr_tab(self, nb):
        tab = ttk.Frame(nb)
        nb.add(tab, text='  Поиск и замена  ')

        # строка поиска
        top = ttk.Frame(tab)
        top.pack(fill='x', padx=6, pady=6)
        ttk.Label(top, text='Найти:').pack(side='left')
        self.fr_find = ttk.Entry(top, width=34)
        self.fr_find.pack(side='left', padx=6)
        self.fr_case = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text='Учитывать регистр', variable=self.fr_case).pack(side='left', padx=6)
        self.fr_partial = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text='Часть слова', variable=self.fr_partial).pack(side='left', padx=6)
        self.fr_find_btn = ttk.Button(top, text='🔍  Найти', command=self.start_scan)
        self.fr_find_btn.pack(side='left', padx=6)

        # результаты
        mid = ttk.LabelFrame(tab, text='Результаты (отметьте, что заменять; клик по флажку)')
        mid.pack(fill='both', expand=True, padx=6, pady=4)
        selrow = ttk.Frame(mid)
        selrow.pack(fill='x', padx=6, pady=(6, 0))
        ttk.Button(selrow, text='Отметить все', command=lambda: self._fr_set_all(True)).pack(side='left')
        ttk.Button(selrow, text='Снять все', command=lambda: self._fr_set_all(False)).pack(side='left', padx=6)
        self.fr_count = ttk.Label(selrow, text='Ничего не искали', foreground='#666')
        self.fr_count.pack(side='left', padx=12)

        treef = ttk.Frame(mid)
        treef.pack(fill='both', expand=True, padx=6, pady=6)
        self.fr_tree = ttk.Treeview(treef, columns=('chk',), show='tree headings', selectmode='none')
        self.fr_tree.heading('#0', text='Документ / контекст (⟦…⟧ — найденное)')
        self.fr_tree.heading('chk', text='Заменять')
        self.fr_tree.column('#0', width=560, stretch=True)
        self.fr_tree.column('chk', width=80, anchor='center', stretch=False)
        self.fr_tree.pack(side='left', fill='both', expand=True)
        tsb = ttk.Scrollbar(treef, command=self.fr_tree.yview)
        tsb.pack(side='right', fill='y')
        self.fr_tree['yscrollcommand'] = tsb.set
        self.fr_tree.bind('<Button-1>', self._fr_click)

        # строка замены
        bot = ttk.Frame(tab)
        bot.pack(fill='x', padx=6, pady=6)
        ttk.Label(bot, text='Заменить на:').pack(side='left')
        self.fr_repl = ttk.Entry(bot, width=34)
        self.fr_repl.pack(side='left', padx=6)
        self.fr_repl_btn = ttk.Button(bot, text='Заменить отмеченное', command=self.start_replace)
        self.fr_repl_btn.pack(side='left', padx=6)
        self.fr_repl_btn['state'] = 'disabled'

        logf = ttk.LabelFrame(tab, text='Журнал')
        logf.pack(fill='x', padx=6, pady=(0, 6))
        self.fr_log = tk.Text(logf, height=5, wrap='word', state='disabled')
        self.fr_log.pack(side='left', fill='both', expand=True, padx=(6, 0), pady=6)
        fsb = ttk.Scrollbar(logf, command=self.fr_log.yview)
        fsb.pack(side='right', fill='y', pady=6)
        self.fr_log['yscrollcommand'] = fsb.set

    # ------------------------------------------------------- источник (общий)
    def _refresh_list(self):
        self.path_list.delete(0, 'end')
        for p in self.paths:
            tag = '[папка] ' if os.path.isdir(p) else ''
            self.path_list.insert('end', tag + p)

    def add_files(self):
        files = filedialog.askopenfilenames(
            title='Выберите .docx / .doc',
            filetypes=[('Word', '*.docx *.doc'), ('Все файлы', '*.*')])
        for f in files:
            if f not in self.paths:
                self.paths.append(f)
        self._refresh_list()
        self._persist_last()

    def add_folder(self):
        d = filedialog.askdirectory(title='Выберите папку (обрабатываются все .docx/.doc внутри)')
        if d and d not in self.paths:
            self.paths.append(d)
        self._refresh_list()
        self._persist_last()

    def remove_selected(self):
        for i in reversed(self.path_list.curselection()):
            del self.paths[i]
        self._refresh_list()
        self._persist_last()

    def clear_paths(self):
        self.paths = []
        self._refresh_list()
        self._persist_last()

    # ---- именованные наборы путей ----
    def _refresh_presets(self):
        names = sorted(self.store.get('presets', {}).keys())
        self.preset_combo['values'] = names
        if self.preset_combo.get() not in names:
            self.preset_combo.set('')

    def save_preset(self):
        if not self.paths:
            messagebox.showinfo('Пусто', 'Сначала добавьте файлы или папки.')
            return
        name = simpledialog.askstring('Сохранить набор', 'Название набора:', parent=self.root)
        if not name or not name.strip():
            return
        name = name.strip()
        if name in self.store['presets'] and not messagebox.askyesno(
                'Перезаписать', f'Набор «{name}» уже существует. Перезаписать?'):
            return
        self.store['presets'][name] = list(self.paths)
        self._save_store()
        self._refresh_presets()
        self.preset_combo.set(name)

    def load_preset(self):
        name = self.preset_combo.get()
        if not name or name not in self.store.get('presets', {}):
            messagebox.showinfo('Нет набора', 'Выберите сохранённый набор из списка.')
            return
        self.paths = list(self.store['presets'][name])
        self._refresh_list()
        self._persist_last()

    def delete_preset(self):
        name = self.preset_combo.get()
        if not name or name not in self.store.get('presets', {}):
            return
        if messagebox.askyesno('Удалить набор', f'Удалить набор «{name}»?'):
            del self.store['presets'][name]
            self._save_store()
            self._refresh_presets()

    # ------------------------------------------------------------- очередь/лог
    def _log_to(self, widget, msg):
        widget['state'] = 'normal'
        widget.insert('end', msg + '\n')
        widget.see('end')
        widget['state'] = 'disabled'

    def log(self, msg):
        self.msg_queue.put(('log', msg))

    def fr_log_msg(self, msg):
        self.msg_queue.put(('fr_log', msg))

    def _drain_queue(self):
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                if kind == 'log':
                    self._log_to(self.log_text, payload)
                elif kind == 'fr_log':
                    self._log_to(self.fr_log, payload)
                elif kind == 'progress':
                    done, total = payload
                    self.progress['maximum'] = max(total, 1)
                    self.progress['value'] = done
                elif kind == 'fr_results':
                    self.fr_query, self.fr_match_case, self.fr_whole_word, self.fr_results = payload
                    self._fr_populate()
                    self._set_busy(False)
                elif kind == 'done':
                    self._set_busy(False)
                elif kind == 'fr_done':
                    self._set_busy(False)
        except queue.Empty:
            pass
        self.root.after(100, self._drain_queue)

    def _ask_doc_handling(self, doc_files):
        """Спрашивает, что делать с найденными .doc. Возвращает кортеж
        (decision, delete_originals):
          decision: 'convert' | 'skip' | 'cancel';
          delete_originals: удалять ли исходные .doc после конвертации."""
        if not doc_files:
            return 'skip', False
        n = len(doc_files)
        if not DC.available():
            ok = messagebox.askokcancel(
                'Найдены файлы .doc',
                f'Найдено файлов .doc (старый формат): {n}.\n'
                'Автоконвертация недоступна (нет Word/pywin32) — они будут пропущены.\n\n'
                'Продолжить с .docx?   (Отмена — прервать)')
            return ('skip' if ok else 'cancel'), False
        return self._ask_doc_dialog(n)

    def _ask_doc_dialog(self, n):
        """Модальный диалог с флажком удаления исходников.
        Возвращает (decision, delete_originals)."""
        dlg = tk.Toplevel(self.root)
        dlg.title('Найдены файлы .doc')
        dlg.transient(self.root)
        dlg.resizable(False, False)
        result = {'decision': 'cancel'}
        del_var = tk.BooleanVar(value=True)

        msg = (f'Найдено файлов .doc (старый формат Word): {n}.\n'
               'Напрямую они не обрабатываются, но можно сконвертировать в .docx через Word.\n\n'
               'Да — сконвертировать и обработать\n'
               'Нет — пропустить .doc, работать только с .docx\n'
               'Отмена — ничего не делать')
        ttk.Label(dlg, text=msg, justify='left').pack(padx=16, pady=(14, 8), anchor='w')
        ttk.Checkbutton(dlg, text='Удалить исходные .doc после успешной конвертации',
                        variable=del_var).pack(padx=16, anchor='w')
        btns = ttk.Frame(dlg)
        btns.pack(padx=16, pady=12, anchor='e')

        def choose(d):
            result['decision'] = d
            result['delete'] = del_var.get()
            dlg.destroy()

        ttk.Button(btns, text='Да', command=lambda: choose('convert')).pack(side='left')
        ttk.Button(btns, text='Нет', command=lambda: choose('skip')).pack(side='left', padx=6)
        ttk.Button(btns, text='Отмена', command=lambda: choose('cancel')).pack(side='left')
        dlg.protocol('WM_DELETE_WINDOW', lambda: choose('cancel'))

        dlg.grab_set()
        dlg.update_idletasks()
        x = self.root.winfo_rootx() + max((self.root.winfo_width() - dlg.winfo_width()) // 2, 0)
        y = self.root.winfo_rooty() + max((self.root.winfo_height() - dlg.winfo_height()) // 2, 0)
        dlg.geometry(f'+{x}+{y}')
        self.root.wait_window(dlg)
        return result.get('decision', 'cancel'), result.get('delete', True)

    def _convert_docs(self, doc_files, log, delete_originals=False):
        """Конвертирует .doc → .docx (рядом) через Word. Возвращает список
        путей .docx. log — колбэк вывода (self.log или self.fr_log_msg).
        delete_originals — удалять ли исходный .doc, но ТОЛЬКО если .docx был
        реально создан сейчас (существовавший ранее .docx не считается своим и
        исходник не трогаем). Выполняется в рабочем потоке."""
        out = []
        if not doc_files:
            return out
        if not DC.available():
            log(f'Найдено .doc: {len(doc_files)} — но Word/pywin32 недоступны, файлы пропущены.')
            return out
        log(f'Конвертация .doc → .docx через Word: {len(doc_files)} файл(ов)…')
        conv = None
        try:
            conv = DC.WordConverter()
            for d in doc_files:
                target = os.path.splitext(os.path.abspath(d))[0] + '.docx'
                pre_existed = os.path.exists(target)
                try:
                    t = conv.convert(d)
                    out.append(t)
                    log(f'   ✓ {os.path.basename(d)} → {os.path.basename(t)}')
                    if delete_originals and not pre_existed and os.path.exists(t):
                        try:
                            os.remove(d)
                            log('      удалён исходный .doc')
                        except OSError as e:
                            log(f'      не удалось удалить .doc: {e}')
                except Exception as e:
                    log(f'   [ОШИБКА конвертации] {os.path.basename(d)}: {e}')
        except DC.ConversionUnavailable as e:
            log(f'Не удалось запустить Word: {e}')
        finally:
            if conv:
                conv.close()
        return out

    def _set_busy(self, busy):
        self.busy = busy
        state = 'disabled' if busy else 'normal'
        self.typo_btn['state'] = state
        self.fr_find_btn['state'] = state
        # кнопка замены активна только когда есть результаты и мы не заняты
        if busy:
            self.fr_repl_btn['state'] = 'disabled'
        elif self.fr_results:
            self.fr_repl_btn['state'] = 'normal'

    # --------------------------------------------------------- типографика
    def start_typo(self):
        if self.busy:
            return
        do_typo = self.var_typo.get()
        if not do_typo:
            messagebox.showwarning('Нечего делать', 'Включите тумблер «Типографика».')
            return
        if not self.paths:
            messagebox.showwarning('Нет входных данных', 'Добавьте файлы или папку.')
            return
        docx_files = collect_docx(self.paths)
        doc_files = collect_doc(self.paths)
        if not docx_files and not doc_files:
            messagebox.showinfo('Пусто', 'В выбранных путях не найдено .docx или .doc.')
            return
        decision, delete_doc = self._ask_doc_handling(doc_files)
        if decision == 'cancel':
            return
        use_doc = doc_files if decision == 'convert' else []
        if not docx_files and not use_doc:
            messagebox.showinfo('Пусто', 'Нечего обрабатывать (.doc пропущены).')
            return
        in_place = self.var_inplace.get()
        extra = f' + {len(use_doc)} .doc (конвертация в .docx)' if use_doc else ''
        if in_place and not messagebox.askyesno(
                'Перезапись оригиналов',
                f'Будут ПЕРЕЗАПИСАНЫ {len(docx_files)} .docx{extra}.\n'
                'Убедитесь, что есть бэкап. Продолжить?'):
            return

        self._set_busy(True)
        self._clear_text(self.log_text)
        args = (docx_files, use_doc, delete_doc, do_typo, in_place, self.var_report.get())
        threading.Thread(target=self._run_typo, args=args, daemon=True).start()

    def _run_typo(self, docx_files, doc_files, delete_doc, do_typo, in_place, make_report):
        files = list(docx_files)
        for c in self._convert_docs(doc_files, self.log, delete_doc):
            if c not in files:
                files.append(c)
        total = len(files)
        mode = 'перезапись на месте' if in_place else f'копия ({SUFFIX})'
        self.log(f'Файлов: {total}   |   этап: типографика   |   выход: {mode}')
        self.log('=' * 70)
        self.msg_queue.put(('progress', (0, total)))
        ok_count = err_count = 0
        for i, f in enumerate(files, 1):
            self.log(f'[{i}/{total}] {f}')
            try:
                ok = process_one(f, False, do_typo, in_place, make_report, self.log)
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

    # ------------------------------------------------------- поиск и замена
    def _clear_text(self, widget):
        widget['state'] = 'normal'
        widget.delete('1.0', 'end')
        widget['state'] = 'disabled'

    def start_scan(self):
        if self.busy:
            return
        query = self.fr_find.get()
        if not query:
            messagebox.showwarning('Пустой запрос', 'Введите слово или словосочетание для поиска.')
            return
        if not self.paths:
            messagebox.showwarning('Нет входных данных', 'Добавьте файлы или папку.')
            return
        docx_files = collect_docx(self.paths)
        doc_files = collect_doc(self.paths)
        if not docx_files and not doc_files:
            messagebox.showinfo('Пусто', 'В выбранных путях не найдено .docx или .doc.')
            return
        decision, delete_doc = self._ask_doc_handling(doc_files)
        if decision == 'cancel':
            return
        use_doc = doc_files if decision == 'convert' else []
        if not docx_files and not use_doc:
            messagebox.showinfo('Пусто', 'Нечего искать (.doc пропущены).')
            return

        self._set_busy(True)
        self._clear_text(self.fr_log)
        self.fr_tree.delete(*self.fr_tree.get_children())
        self.fr_results = {}
        self.fr_count['text'] = 'Идёт поиск…'
        whole_word = not self.fr_partial.get()
        self.fr_log_msg(f'Поиск «{query}» '
                        f'({"с учётом" if self.fr_case.get() else "без учёта"} регистра, '
                        f'{"часть слова" if not whole_word else "целое слово"})…')
        args = (docx_files, use_doc, delete_doc, query, self.fr_case.get(), whole_word)
        threading.Thread(target=self._run_scan, args=args, daemon=True).start()

    def _run_scan(self, docx_files, doc_files, delete_doc, query, match_case, whole_word):
        files = list(docx_files)
        for c in self._convert_docs(doc_files, self.fr_log_msg, delete_doc):
            if c not in files:
                files.append(c)
        total = len(files)
        self.msg_queue.put(('progress', (0, total)))
        results = {}
        for i, f in enumerate(files, 1):
            try:
                m = FR.scan_file(f, query, match_case, whole_word)
            except Exception:
                m = []
                self.fr_log_msg(f'[ОШИБКА чтения] {f}')
            if m:
                results[f] = m
            self.msg_queue.put(('progress', (i, total)))
        self.msg_queue.put(('fr_results', (query, match_case, whole_word, results)))

    def _fr_populate(self):
        tree = self.fr_tree
        tree.delete(*tree.get_children())
        self.fr_item_meta = {}
        total = 0
        for path, matches in self.fr_results.items():
            doc_id = tree.insert('', 'end', text=f'📄 {os.path.basename(path)}  ({len(matches)})',
                                 values=(CHECKED,), open=True)
            self.fr_item_meta[doc_id] = {'type': 'doc', 'path': path}
            for md in matches:
                ctx = (md['before'] + '⟦' + md['match'] + '⟧' + md['after']).replace('\n', ' ').replace('\r', ' ')
                cid = tree.insert(doc_id, 'end', text='    ' + ctx, values=(CHECKED,))
                self.fr_item_meta[cid] = {'type': 'match', 'path': path,
                                          'para_idx': md['para_idx'], 'occ_idx': md['occ_idx'],
                                          'checked': True}
                total += 1
        if total:
            self.fr_count['text'] = f'Найдено вхождений: {total} в {len(self.fr_results)} документах'
            self.fr_log_msg(f'Найдено: {total} вхождений в {len(self.fr_results)} документах.')
        else:
            self.fr_count['text'] = 'Ничего не найдено'
            self.fr_log_msg('Ничего не найдено.')

    def _fr_click(self, event):
        tree = self.fr_tree
        if tree.identify('region', event.x, event.y) not in ('tree', 'cell'):
            return
        if tree.identify_column(event.x) != '#1':   # реагируем только на колонку-флажок
            return
        row = tree.identify_row(event.y)
        meta = self.fr_item_meta.get(row)
        if not meta:
            return
        if meta['type'] == 'doc':
            kids = tree.get_children(row)
            new_state = not all(self.fr_item_meta[k]['checked'] for k in kids)
            for k in kids:
                self.fr_item_meta[k]['checked'] = new_state
                tree.set(k, 'chk', CHECKED if new_state else UNCHECKED)
            tree.set(row, 'chk', CHECKED if new_state else UNCHECKED)
        else:
            meta['checked'] = not meta['checked']
            tree.set(row, 'chk', CHECKED if meta['checked'] else UNCHECKED)
            self._fr_refresh_doc(tree.parent(row))

    def _fr_refresh_doc(self, doc_id):
        tree = self.fr_tree
        states = [self.fr_item_meta[k]['checked'] for k in tree.get_children(doc_id)]
        if states and all(states):
            tree.set(doc_id, 'chk', CHECKED)
        elif not any(states):
            tree.set(doc_id, 'chk', UNCHECKED)
        else:
            tree.set(doc_id, 'chk', PARTIAL)

    def _fr_set_all(self, state):
        tree = self.fr_tree
        for item_id, meta in self.fr_item_meta.items():
            if meta['type'] == 'match':
                meta['checked'] = state
            tree.set(item_id, 'chk', CHECKED if state else UNCHECKED)

    def _fr_selected_by_file(self):
        """{path: {para_idx: set(occ_idx)}} по отмеченным совпадениям."""
        sel = {}
        for meta in self.fr_item_meta.values():
            if meta['type'] == 'match' and meta['checked']:
                sel.setdefault(meta['path'], {}).setdefault(meta['para_idx'], set()).add(meta['occ_idx'])
        return sel

    def start_replace(self):
        if self.busy or not self.fr_results:
            return
        replacement = self.fr_repl.get()
        selected = self._fr_selected_by_file()
        if not selected:
            messagebox.showinfo('Ничего не отмечено', 'Отметьте хотя бы одно вхождение для замены.')
            return
        total_occ = sum(len(occ) for paras in selected.values() for occ in paras.values())
        in_place = self.var_inplace.get()
        where = 'ПЕРЕЗАПИШЕТ оригиналы' if in_place else f'сохранит копии ({SUFFIX})'
        if not messagebox.askyesno(
                'Подтверждение замены',
                f'Заменить «{self.fr_query}» → «{replacement}»\n'
                f'в {total_occ} вхождениях ({len(selected)} файлов).\n'
                f'Операция {where}. Продолжить?'):
            return

        self._set_busy(True)
        args = (self.fr_query, replacement, self.fr_match_case, self.fr_whole_word, in_place, selected)
        threading.Thread(target=self._run_replace, args=args, daemon=True).start()

    def _run_replace(self, query, replacement, match_case, whole_word, in_place, selected):
        total = len(selected)
        self.fr_log_msg('=' * 60)
        self.fr_log_msg(f'Замена «{query}» → «{replacement}» в {total} файлах…')
        self.msg_queue.put(('progress', (0, total)))
        done = made = errs = 0
        for path, sel in selected.items():
            out = output_path_for(path, in_place)
            try:
                n = FR.replace_file(path, out, query, replacement, match_case, whole_word, sel)
                made += n
                self.fr_log_msg(f'✓ {n} замен: {out}')
            except Exception:
                errs += 1
                self.fr_log_msg(f'[ОШИБКА] {path}: ' + traceback.format_exc().splitlines()[-1])
            done += 1
            self.msg_queue.put(('progress', (done, total)))
        self.fr_log_msg('=' * 60)
        self.fr_log_msg(f'Готово. Заменено вхождений: {made}, файлов: {done}, ошибок: {errs}.')
        self.msg_queue.put(('fr_done', None))


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
