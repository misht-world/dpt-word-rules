# -*- coding: utf-8 -*-
"""
Применяет типографские правила (rules.py) к word/document.xml внутри распакованного .docx.
Работает на уровне абзаца: собирает весь текст абзаца из всех <w:t>, прогоняет через
конвейер правил, а изменения аккуратно распределяет обратно по исходным run'ам
(через difflib), чтобы не потерять форматирование (жирный/курсив/др.) внутри абзаца.

Использование:
    python3 apply_docx.py input.docx output.docx [--dry-run] [--report report.txt]
"""
import sys
import os
import shutil
import zipfile
import difflib
import argparse
from lxml import etree

sys.path.insert(0, os.path.dirname(__file__))
import rules as R

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W_NS}
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


def qn(tag):
    return '{%s}%s' % (W_NS, tag)


def merge_runs_lxml(root):
    """
    Лёгкая самодостаточная замена внешнего merge_runs.py (без зависимости от
    /mnt/skills — чтобы скрипт работал автономно и локально).
    Сливает соседние <w:r> с одинаковым <w:rPr> внутри одного родителя
    (абзаца или гиперссылки), объединяя их текст в один <w:t>.
    Не трогает runs, находящиеся в разных <w:ins>/<w:del> (они просто имеют
    разных родителей, поэтому не сливаются между собой — и это корректно).

    ВАЖНО: сливаем только run'ы, у которых нет ДРУГИХ детей, кроме <w:rPr> и
    одного <w:t>. Иначе при удалении второго run'а потеряются его прочие
    дочерние элементы: <w:noBreakHyphen/> (неразрывный дефис в Санкт-Петербург,
    д. 13-Б и т.п.), <w:tab/>, <w:br/>, <w:cr/>, <w:sym/>, картинки и др.
    Именно так раньше пропадал дефис — он хранится как <w:noBreakHyphen/>, а не
    как символ в тексте.
    """
    def _only_rpr_and_one_t(r):
        allowed = (qn('rPr'), qn('t'))
        for child in r:
            if child.tag not in allowed:
                return False
        return len(r.findall('w:t', NS)) == 1

    merged = 0
    for parent in root.iter():
        children = list(parent)
        run_children = [c for c in children if c.tag == qn('r')]
        if len(run_children) < 2:
            continue
        i = 0
        while i < len(children) - 1:
            a, b = children[i], children[i + 1]
            if a.tag == qn('r') and b.tag == qn('r'):
                rpr_a = a.find('w:rPr', NS)
                rpr_b = b.find('w:rPr', NS)
                xml_a = etree.tostring(rpr_a) if rpr_a is not None else None
                xml_b = etree.tostring(rpr_b) if rpr_b is not None else None
                # сливаем, только если между ними в дереве нет ничего постороннего
                # (list(parent) уже гарантирует непосредственное соседство)
                if xml_a == xml_b:
                    ts_a = a.findall('w:t', NS)
                    ts_b = b.findall('w:t', NS)
                    if (len(ts_a) == 1 and len(ts_b) == 1
                            and _only_rpr_and_one_t(a) and _only_rpr_and_one_t(b)):
                        combined = (ts_a[0].text or '') + (ts_b[0].text or '')
                        ts_a[0].text = combined
                        if combined != combined.strip() or combined == '':
                            ts_a[0].set(XML_SPACE, 'preserve')
                        parent.remove(b)
                        children = list(parent)
                        merged += 1
                        continue
            i += 1
    return merged


def collect_paragraph_text(p):
    """Возвращает (full_text, [(text_node, start, end), ...]) для абзаца."""
    t_nodes = p.findall('.//w:t', NS)
    full = []
    spans = []
    pos = 0
    for t in t_nodes:
        s = t.text or ''
        spans.append((t, pos, pos + len(s)))
        full.append(s)
        pos += len(s)
    return ''.join(full), spans


def redistribute(old_text, new_text, spans):
    """
    Строит новый текст для каждого text_node на основе diff между old_text и new_text,
    используя исходные позиции символов (spans) для привязки к run'ам.
    Возвращает dict {text_node: new_string}.
    """
    if old_text == new_text:
        return None

    # char_owner[i] = индекс text_node, которому принадлежит old_text[i]
    char_owner = [None] * len(old_text)
    for idx, (node, s, e) in enumerate(spans):
        for i in range(s, e):
            char_owner[i] = idx

    sm = difflib.SequenceMatcher(None, old_text, new_text, autojunk=False)
    buckets = {idx: [] for idx in range(len(spans))}

    def owner_for_old_pos(pos):
        # ближайший владелец: сначала сам символ, потом соседи слева/справа
        if 0 <= pos < len(char_owner) and char_owner[pos] is not None:
            return char_owner[pos]
        # ищем слева
        p2 = pos - 1
        while p2 >= 0:
            if char_owner[p2] is not None:
                return char_owner[p2]
            p2 -= 1
        # ищем справа
        p2 = pos
        while p2 < len(char_owner):
            if char_owner[p2] is not None:
                return char_owner[p2]
            p2 += 1
        return 0 if spans else None

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            for k in range(i2 - i1):
                owner = char_owner[i1 + k]
                buckets.setdefault(owner, []).append(new_text[j1 + k])
        elif tag == 'replace' or tag == 'insert':
            # весь новый фрагмент отдаём владельцу первой "старой" позиции этого места
            owner = owner_for_old_pos(i1)
            if owner is None:
                continue
            buckets.setdefault(owner, []).append(new_text[j1:j2])
        elif tag == 'delete':
            pass  # ничего не добавляем

    result = {}
    for idx, (node, s, e) in enumerate(spans):
        parts = buckets.get(idx, [])
        result[node] = ''.join(parts)
    return result


def ends_with_colon(text):
    return text.rstrip().endswith(':') and text.rstrip() != ''


def set_keep_next(p):
    ppr = p.find('w:pPr', NS)
    if ppr is None:
        ppr = etree.SubElement(p, qn('pPr'))
        p.insert(0, ppr)
    if ppr.find('w:keepNext', NS) is None:
        kn = etree.SubElement(ppr, qn('keepNext'))


def process(doc_path, dry_run=False, report_path=None):
    tree = etree.parse(doc_path)
    root = tree.getroot()

    merged = merge_runs_lxml(root)

    paras = root.findall('.//w:p', NS)

    report_lines = []
    total_changed = 0
    keepnext_added = 0

    for p in paras:
        full, spans = collect_paragraph_text(p)
        if not full.strip():
            continue

        new_text = R.apply_all(full, log=True)

        if new_text != full:
            total_changed += 1
            if report_path:
                report_lines.append("ABZAC:")
                report_lines.append("  БЫЛО: " + full)
                report_lines.append("  СТАЛО: " + new_text.replace(R.NBSP, '\u2038').replace(R.NBH, '\u2038h\u2038'))
                for name, start, old, new in R.LOG:
                    report_lines.append(f"    [{name}] {old!r} -> {new!r}")
                report_lines.append("")

            if not dry_run:
                new_vals = redistribute(full, new_text, spans)
                if new_vals:
                    for node, val in new_vals.items():
                        node.text = val
                        # сохраняем пробелы на границах
                        if val != val.strip() or val == '':
                            node.set(XML_SPACE, 'preserve')

        # keepNext — по НОВОМУ тексту (после правок), не по старому
        check_text = new_text if not dry_run else new_text
        if ends_with_colon(check_text):
            keepnext_added += 1
            if not dry_run:
                set_keep_next(p)

    if report_path:
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(f"Изменено абзацев: {total_changed}\n")
            f.write(f"Добавлен keepNext: {keepnext_added}\n\n")
            f.write('\n'.join(report_lines))

    print(f"Изменено абзацев: {total_changed}")
    print(f"Добавлен keepNext (абзац оканчивается на ':'): {keepnext_added}")

    if not dry_run:
        tree.write(doc_path, xml_declaration=True, encoding='UTF-8', standalone=True)


def main():
    ap = argparse.ArgumentParser(description="Типографское оформление .docx (неразрывные пробелы, дефисы, единицы измерения и т.д.)")
    ap.add_argument('input', help='Исходный .docx')
    ap.add_argument('output', nargs='?', help='Куда сохранить результат (.docx). Не нужен при --dry-run.')
    ap.add_argument('--dry-run', action='store_true', help='Только показать список замен, не изменяя файл')
    ap.add_argument('--report', help='Путь к текстовому отчёту со списком всех замен')
    args = ap.parse_args()

    import tempfile
    workdir = tempfile.mkdtemp(prefix='docformat_')

    with zipfile.ZipFile(args.input) as z:
        z.extractall(workdir)

    # убрать symlink-подобные опасные записи (untrusted source)
    for dirpath, dirnames, filenames in os.walk(workdir):
        for fn in filenames:
            full = os.path.join(dirpath, fn)
            if os.path.islink(full):
                os.remove(full)

    doc_xml = os.path.join(workdir, 'word', 'document.xml')

    process(doc_xml, dry_run=args.dry_run, report_path=args.report)

    if not args.dry_run:
        if not args.output:
            print("Укажите путь для сохранения результата (output.docx)")
            sys.exit(1)
        if os.path.exists(args.output):
            os.remove(args.output)
        # запаковать обратно
        base = os.path.abspath(args.output)
        with zipfile.ZipFile(base, 'w', zipfile.ZIP_DEFLATED) as zf:
            for dirpath, dirnames, filenames in os.walk(workdir):
                for fn in filenames:
                    full = os.path.join(dirpath, fn)
                    rel = os.path.relpath(full, workdir)
                    zf.write(full, rel)
        print(f"Сохранено: {args.output}")

    shutil.rmtree(workdir)


if __name__ == '__main__':
    main()
