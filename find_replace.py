# -*- coding: utf-8 -*-
"""
Поиск и замена слов/словосочетаний во множестве .docx (в т.ч. по каталогам
рекурсивно).

Работает на уровне ТЕКСТА АБЗАЦА (собирает все <w:t> абзаца), поэтому находит
фразы, которые Word произвольно разбил на несколько run'ов. Замена возвращается
обратно в run'ы через redistribute() из apply_docx — форматирование
(жирный/курсив/цвет) внутри абзаца не теряется.

Индексация: и при поиске, и при замене абзацы перебираются одинаково
(root.findall('.//w:p') в порядке документа), поэтому пара (номер абзаца,
номер вхождения в абзаце) стабильно идентифицирует конкретное совпадение между
этапами «показать результаты» и «заменить выбранное».

Ограничение: обрабатывается только основной текст (word/document.xml).
Колонтитулы, надписи (text boxes) и сноски — отдельные части пакета, здесь не
затрагиваются.
"""
import os
import re
import zipfile
import tempfile
import shutil
from lxml import etree

import apply_docx as A  # collect_paragraph_text, redistribute, XML_SPACE

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W_NS}


# Классы «эквивалентных» символов для гибкого поиска (flex_ws):
# любой пробел и любой дефис/тире считаются одинаковыми при поиске.
_SPACE_CHARS = ' \t '                 # пробел, таб, неразрывный пробел
_HYPHEN_CHARS = '-‐‑‒–—'  # дефис, неразр. дефис, тире en/em
_SPACE_CLASS = '[' + re.escape(_SPACE_CHARS) + ']'
_HYPHEN_CLASS = '[' + re.escape(_HYPHEN_CHARS) + ']'


def _build_pattern(query, whole_word, flex_ws):
    """Регэксп для запроса. При flex_ws каждый пробел/дефис заменяется классом
    «любой пробел»/«любой дефис». Каждый символ запроса матчит РОВНО один символ,
    поэтому длина совпадения всегда равна len(query) — это важно для замены."""
    parts = []
    for ch in query:
        if flex_ws and ch in _SPACE_CHARS:
            parts.append(_SPACE_CLASS)
        elif flex_ws and ch in _HYPHEN_CHARS:
            parts.append(_HYPHEN_CLASS)
        else:
            parts.append(re.escape(ch))
    pat = ''.join(parts)
    if whole_word:
        pat = r'(?<!\w)' + pat + r'(?!\w)'
    return pat


def find_occurrences(text, query, match_case, whole_word=True, flex_ws=True):
    """Список стартовых индексов НЕПЕРЕКРЫВАЮЩИХСЯ вхождений query в text.

    whole_word=True: совпадение только как отдельное слово/фраза (по краям нет
    буквы/цифры/подчёркивания); False — поиск подстроки.
    flex_ws=True: тип пробела (обычный/неразрывный) и дефиса/тире не важен.
    Длина каждого совпадения равна len(query), поэтому индексы совместимы с заменой.
    """
    if not query:
        return []
    flags = 0 if match_case else re.IGNORECASE
    return [m.start() for m in re.finditer(_build_pattern(query, whole_word, flex_ws), text, flags)]


def scan_file(path, query, match_case, whole_word=True, flex_ws=True, ctx=35):
    """Возвращает список совпадений в одном файле:
    dict(para_idx, occ_idx, before, match, after). Только чтение, без распаковки на диск."""
    try:
        with zipfile.ZipFile(path) as z:
            data = z.read('word/document.xml')
    except (KeyError, zipfile.BadZipFile, OSError):
        return []
    try:
        root = etree.fromstring(data)
    except etree.XMLSyntaxError:
        return []
    n = len(query)
    matches = []
    for pi, p in enumerate(root.findall('.//w:p', NS)):
        full, _ = A.collect_paragraph_text(p)
        if not full:
            continue
        for oi, i in enumerate(find_occurrences(full, query, match_case, whole_word, flex_ws)):
            j = i + n
            matches.append({
                'para_idx': pi,
                'occ_idx': oi,
                'before': full[max(0, i - ctx):i],
                'match': full[i:j],
                'after': full[j:j + ctx],
            })
    return matches


def _build_replaced(text, query, replacement, match_case, whole_word, flex_ws, selected):
    """Заменяет в тексте только те вхождения, чей occ_idx входит в selected (set)."""
    n = len(query)
    out = []
    prev = 0
    for oi, i in enumerate(find_occurrences(text, query, match_case, whole_word, flex_ws)):
        out.append(text[prev:i])
        out.append(replacement if oi in selected else text[i:i + n])
        prev = i + n
    out.append(text[prev:])
    return ''.join(out)


def replace_file(path, out_path, query, replacement, match_case, whole_word, flex_ws, selected_by_para):
    """Применяет замену к выбранным вхождениям и сохраняет результат в out_path.

    selected_by_para: dict para_idx -> set(occ_idx). Возвращает число замен.
    """
    workdir = tempfile.mkdtemp(prefix='dpt_fr_')
    made = 0
    try:
        with zipfile.ZipFile(path) as z:
            z.extractall(workdir)
        # убрать symlink-подобные записи (untrusted source)
        for dirpath, dirnames, filenames in os.walk(workdir):
            for fn in filenames:
                full = os.path.join(dirpath, fn)
                if os.path.islink(full):
                    os.remove(full)

        doc_xml = os.path.join(workdir, 'word', 'document.xml')
        tree = etree.parse(doc_xml)
        root = tree.getroot()

        for pi, p in enumerate(root.findall('.//w:p', NS)):
            sel = selected_by_para.get(pi)
            if not sel:
                continue
            full, spans = A.collect_paragraph_text(p)
            new_text = _build_replaced(full, query, replacement, match_case, whole_word, flex_ws, sel)
            if new_text != full:
                made += len(sel)
                new_vals = A.redistribute(full, new_text, spans)
                if new_vals:
                    for node, val in new_vals.items():
                        node.text = val
                        if val != val.strip() or val == '':
                            node.set(A.XML_SPACE, 'preserve')

        tree.write(doc_xml, xml_declaration=True, encoding='UTF-8', standalone=True)

        if os.path.exists(out_path):
            os.remove(out_path)
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for dirpath, dirnames, filenames in os.walk(workdir):
                for fn in filenames:
                    full = os.path.join(dirpath, fn)
                    rel = os.path.relpath(full, workdir)
                    zf.write(full, rel)
    finally:
        shutil.rmtree(workdir, ignore_errors=True)
    return made
