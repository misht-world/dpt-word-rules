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


def find_occurrences(text, query, match_case, whole_word=True):
    """Список стартовых индексов НЕПЕРЕКРЫВАЮЩИХСЯ вхождений query в text.

    whole_word=True (по умолчанию): совпадение только как отдельное слово/фраза —
    по краям не должно быть буквы/цифры/подчёркивания (напр. поиск «точных вод»
    НЕ находит «сточных вод»). whole_word=False: поиск как подстроки.
    Длина каждого совпадения равна len(query) (шаблон — экранированный запрос),
    поэтому индексы совместимы с заменой.
    """
    if not query:
        return []
    flags = 0 if match_case else re.IGNORECASE
    pat = re.escape(query)
    if whole_word:
        pat = r'(?<!\w)' + pat + r'(?!\w)'
    return [m.start() for m in re.finditer(pat, text, flags)]


def scan_file(path, query, match_case, whole_word=True, ctx=35):
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
        for oi, i in enumerate(find_occurrences(full, query, match_case, whole_word)):
            j = i + n
            matches.append({
                'para_idx': pi,
                'occ_idx': oi,
                'before': full[max(0, i - ctx):i],
                'match': full[i:j],
                'after': full[j:j + ctx],
            })
    return matches


def _build_replaced(text, query, replacement, match_case, whole_word, selected):
    """Заменяет в тексте только те вхождения, чей occ_idx входит в selected (set)."""
    n = len(query)
    out = []
    prev = 0
    for oi, i in enumerate(find_occurrences(text, query, match_case, whole_word)):
        out.append(text[prev:i])
        out.append(replacement if oi in selected else text[i:i + n])
        prev = i + n
    out.append(text[prev:])
    return ''.join(out)


def replace_file(path, out_path, query, replacement, match_case, whole_word, selected_by_para):
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
            new_text = _build_replaced(full, query, replacement, match_case, whole_word, sel)
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
