# -*- coding: utf-8 -*-
"""
Приводит структуру .docx к единому виду по style_spec.json:
- вклеивает провалидированные стили «ДПТ - ...» (через style_injector.py)
- обычный текст -> стиль «ДПТ - Основной»
- маркированные/нумерованные перечисления (вручную напечатанные "- " / "N) ")
  -> настоящие Word-списки «ДПТ - Список (маркер/нумерация)» (текст маркера
  вырезается, Word генерирует номер сам)
- таблицы (реальные, не layout-обёртки) -> единая рамка/шрифт/шапка
- поля страницы (только книжные/portrait секции)

ЗАГОЛОВКИ — единая схема нумерации на весь документ (DPTHeading1..5),
включая уже стилизованные ("родные" outlineLvl-стили документа, часто через
несколько звеньев basedOn) И "сломанные" (номер напечатан вручную в тексте,
без реального стиля). ВАЖНО: родная нумерация существующих документов не
переиспользуется (она обычно устроена сложно и индивидуально для каждого
документа — цепочки numStyleLink и т.п., ненадёжно обобщать) — вместо этого
ВСЕ заголовки документа переводятся на одну нашу схему нумерации, поэтому
итоговые номера всегда взаимно согласованы (никогда не будет смеси "своей" и
"нашей" нумерации). Подробности и история решения — см. FORMATTING_SPEC.md.

Использование:
    python3 normalize_structure.py input.docx output.docx --report report.txt
    python3 normalize_structure.py input.docx --dry-run --report report.txt
"""
import sys
import os
import re
import json
import shutil
import zipfile
import argparse
from lxml import etree

sys.path.insert(0, os.path.dirname(__file__))
import style_injector

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W_NS}

SPEC_PATH = os.path.join(os.path.dirname(__file__), 'style_spec.json')


def qn(tag):
    return '{%s}%s' % (W_NS, tag)


BULLET_RE = re.compile(r'^-[ \t\u00A0]+(\S.*)$', re.DOTALL)
NUMBERED_RE = re.compile(r'^\d+\)[ \t\u00A0]+(\S.*)$', re.DOTALL)
HEADING_RE = re.compile(r'^(\d+(?:\.\d+){0,4})\.?[ \t\u00A0]+([А-ЯЁ\d].*)$', re.DOTALL)

MAX_HEADING_LEVEL = 5  # столько уровней сейчас определено в dpt_styles_fragment.xml


# ---------------------------------------------------------------------------
# Полное разрешение уровня заголовка по цепочке basedOn (не только прямой
# outlineLvl в самом стиле — он часто наследуется от базового стиля через
# несколько звеньев, простая прямая проверка это пропускает).
# ---------------------------------------------------------------------------

def resolve_outline_level(style_id, styles_root, cache):
    if style_id in cache:
        return cache[style_id]
    sid = style_id
    seen = set()
    result = None
    while sid and sid not in seen:
        seen.add(sid)
        s = next((st for st in styles_root.findall('w:style', NS) if st.get(qn('styleId')) == sid), None)
        if s is None:
            break
        outline = s.find('.//w:outlineLvl', NS)
        if outline is not None:
            result = int(outline.get(qn('val')))
            break
        based = s.find('w:basedOn', NS)
        sid = based.get(qn('val')) if based is not None else None
    cache[style_id] = result
    return result


def is_real_data_table(tbl):
    grid = tbl.find('w:tblGrid', NS)
    cols = len(grid.findall('w:gridCol', NS)) if grid is not None else 0
    if cols > 1:
        return True
    trs = tbl.findall('w:tr', NS)
    if trs and max(len(tr.findall('w:tc', NS)) for tr in trs) > 1:
        return True
    return False


def paragraph_in_table(p):
    parent = p.getparent()
    while parent is not None:
        if parent.tag == qn('tbl'):
            return parent
        parent = parent.getparent()
    return None


def get_paragraph_text(p):
    return ''.join(t.text or '' for t in p.findall('.//w:t', NS))


def clear_conflicting_ppr(ppr):
    """Убираем прямое форматирование, которое должно наследоваться от стиля."""
    for tag in ('ind', 'jc', 'spacing'):
        el = ppr.find(f'w:{tag}', NS)
        if el is not None:
            ppr.remove(el)


def clear_conflicting_rpr_fonts(p):
    """Убираем прямые переопределения шрифта/размера в run'ах — оставляем
    жирный/курсив/подчёркивание (это часто осмысленное локальное выделение)."""
    for rpr in p.findall('.//w:r/w:rPr', NS):
        for tag in ('rFonts', 'sz', 'szCs'):
            el = rpr.find(f'w:{tag}', NS)
            if el is not None:
                rpr.remove(el)


def classify_heading_text(full):
    """Возвращает (level, rest_start_pos) для 'сломанного' (без стиля)
    заголовка вида '13.1.1 Текст', либо None. rest_start_pos — позиция
    начала текста после номера (для вырезания номера)."""
    full_stripped = full.strip()
    if not full_stripped:
        return None
    m = HEADING_RE.match(full_stripped)
    if not m:
        return None
    num, rest = m.groups()
    level = num.count('.') + 1
    # похоже на год (сегмент из 4 цифр) - не заголовок
    if level >= 2 and any(len(seg) == 4 for seg in num.split('.')):
        return None
    words = rest.split()
    if len(words) > 14:
        return None
    if rest.rstrip().endswith(('.', ',', ';')) and not rest.isupper():
        return None
    # одноуровневый номер + текст оканчивается на ':' -> похоже на локальный
    # подсписок внутри абзаца ("1. Взрывоопасные предметы (ВОП):"), не на
    # структурный заголовок документа
    if level == 1 and rest.rstrip().endswith(':'):
        return None
    # позиция начала rest в исходной (не strip'нутой) строке
    leading_ws = len(full) - len(full.lstrip())
    start_in_full = leading_ws + m.start(2)
    return level, start_in_full


def clear_heading_run_formatting(p):
    """Для заголовков — в отличие от обычного текста, очищаем ВСЁ прямое
    форматирование run'ов (включая bold/italic/caps), т.к. оформление
    заголовка должно целиком определяться стилем DPTHeadingN, а не
    случайными остатками форматирования из старого (несовместимого)
    источника."""
    for rpr in p.findall('.//w:r/w:rPr', NS):
        for tag in ('rFonts', 'sz', 'szCs', 'b', 'bCs', 'i', 'iCs', 'caps', 'u', 'color', 'highlight'):
            el = rpr.find(f'w:{tag}', NS)
            if el is not None:
                rpr.remove(el)


def set_heading_pstyle(p, level):
    level = min(level, MAX_HEADING_LEVEL)
    style_id = f'DPTHeading{level}'
    ppr = p.find('w:pPr', NS)
    if ppr is None:
        ppr = etree.Element(qn('pPr'))
        p.insert(0, ppr)
    # чистим прямые переопределения, которые должны целиком определяться стилем
    for tag in ('ind', 'jc', 'spacing', 'numPr', 'outlineLvl', 'keepNext'):
        el = ppr.find(f'w:{tag}', NS)
        if el is not None:
            ppr.remove(el)
    existing = ppr.find('w:pStyle', NS)
    if existing is None:
        existing = etree.SubElement(ppr, qn('pStyle'))
        ppr.remove(existing)
        ppr.insert(0, existing)
    existing.set(qn('val'), style_id)
    clear_heading_run_formatting(p)
    return level


def set_pstyle(p, style_id):
    ppr = p.find('w:pPr', NS)
    if ppr is None:
        ppr = etree.Element(qn('pPr'))
        p.insert(0, ppr)
    existing = ppr.find('w:pStyle', NS)
    if existing is None:
        existing = etree.SubElement(ppr, qn('pStyle'))
        ppr.remove(existing)
        ppr.insert(0, existing)
    existing.set(qn('val'), style_id)
    clear_conflicting_ppr(ppr)


def apply_page_margins(root, spec):
    """Устанавливает поля страницы для книжных (portrait) секций документа.
    Альбомные (landscape) секции НЕ трогаем — для них нет отдельной
    спецификации полей (обычно landscape используется под широкие таблицы,
    и портретные значения для них не подходят без отдельного решения)."""
    margins = spec['page']['margins_twips']
    sect_prs = root.findall('.//w:sectPr', NS)
    changed = 0
    skipped_landscape = 0
    for sectpr in sect_prs:
        pgsz = sectpr.find('w:pgSz', NS)
        orient = pgsz.get(qn('orient')) if pgsz is not None else None
        if orient == 'landscape':
            skipped_landscape += 1
            continue
        pgmar = sectpr.find('w:pgMar', NS)
        if pgmar is None:
            pgmar = etree.SubElement(sectpr, qn('pgMar'))
        for side in ('top', 'bottom', 'left', 'right'):
            pgmar.set(qn(side), str(margins[side]))
        changed += 1
    return changed, skipped_landscape


def process_document_xml(doc_path, styles_root, report_lines, spec):
    tree = etree.parse(doc_path)
    root = tree.getroot()

    outline_cache = {}
    stats = {'body': 0, 'list_bullet': 0, 'list_numbered': 0,
             'heading_from_style': 0, 'heading_from_text': 0,
             'heading_level_clamped': 0}

    paras = root.findall('.//w:p', NS)
    for p in paras:
        tbl = paragraph_in_table(p)
        full = get_paragraph_text(p)
        if not full.strip():
            continue

        ppr = p.find('w:pPr', NS)
        cur_style = None
        if ppr is not None:
            pstyle_el = ppr.find('w:pStyle', NS)
            cur_style = pstyle_el.get(qn('val')) if pstyle_el is not None else None

        # --- 1. Уже стилизованный заголовок (через ПОЛНУЮ цепочку basedOn) ---
        resolved_level = resolve_outline_level(cur_style, styles_root, outline_cache) if cur_style else None
        if resolved_level is not None:
            level = resolved_level + 1
            if level > MAX_HEADING_LEVEL:
                stats['heading_level_clamped'] += 1
            applied = set_heading_pstyle(p, level)
            stats['heading_from_style'] += 1
            report_lines.append(f"[заголовок из стиля] L{applied} {full[:60]!r}")
            continue

        if tbl is not None:
            continue  # таблицы (кроме случая выше, если бы заголовок уже был реальным) — не трогаем per-паграф

        # --- 2. "Сломанный" заголовок: номер напечатан в тексте вручную ---
        heading_match = classify_heading_text(full)
        if heading_match:
            level, strip_pos = heading_match
            _strip_leading(p, strip_pos)
            applied = set_heading_pstyle(p, level)
            if level > MAX_HEADING_LEVEL:
                stats['heading_level_clamped'] += 1
            stats['heading_from_text'] += 1
            report_lines.append(f"[заголовок из текста] L{applied} {full[:60]!r}")
            continue

        m_bullet = BULLET_RE.match(full)
        m_num = NUMBERED_RE.match(full)

        if m_bullet:
            _strip_leading(p, m_bullet.start(1))
            set_pstyle(p, 'DPTListBullet')
            clear_conflicting_rpr_fonts(p)
            stats['list_bullet'] += 1
            report_lines.append(f"[список-маркер] {full[:60]!r}")
        elif m_num:
            _strip_leading(p, m_num.start(1))
            set_pstyle(p, 'DPTListNumbered')
            clear_conflicting_rpr_fonts(p)
            stats['list_numbered'] += 1
            report_lines.append(f"[список-номер] {full[:60]!r}")
        else:
            set_pstyle(p, 'DPTBody')
            clear_conflicting_rpr_fonts(p)
            stats['body'] += 1

    # --- таблицы: единая рамка/шрифт/шапка ---
    tbls = root.findall('.//w:tbl', NS)
    tables_touched = 0
    for tbl in tbls:
        if not is_real_data_table(tbl):
            continue
        tables_touched += 1
        apply_table_style(tbl)

    sections_touched = apply_page_margins(root, spec)

    tree.write(doc_path, xml_declaration=True, encoding='UTF-8', standalone=True)
    stats['tables'] = tables_touched
    stats['sections'], stats['sections_landscape_skipped'] = sections_touched
    return stats


def _strip_leading(p, n_chars):
    """Убирает первые n_chars символов текста абзаца (маркер списка),
    последовательно обрезая по <w:t> узлам (маркер может быть размазан
    по нескольким run'ам)."""
    remaining = n_chars
    for t in p.findall('.//w:t', NS):
        cur = t.text or ''
        if remaining <= 0:
            break
        if len(cur) <= remaining:
            remaining -= len(cur)
            t.text = ''
        else:
            t.text = cur[remaining:]
            remaining = 0


TABLE_BORDER_SZ = '4'  # 0.5pt в восьмых долях пункта (sz=4 -> 0.5pt в OOXML для tbl borders)


def apply_table_style(tbl):
    tblpr = tbl.find('w:tblPr', NS)
    if tblpr is None:
        tblpr = etree.SubElement(tbl, qn('tblPr'))
        tbl.insert(0, tblpr)

    borders = tblpr.find('w:tblBorders', NS)
    if borders is not None:
        tblpr.remove(borders)
    borders = etree.SubElement(tblpr, qn('tblBorders'))
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = etree.SubElement(borders, qn(side))
        el.set(qn('val'), 'single')
        el.set(qn('sz'), TABLE_BORDER_SZ)
        el.set(qn('space'), '0')
        el.set(qn('color'), 'auto')

    trs = tbl.findall('w:tr', NS)
    for ri, tr in enumerate(trs):
        is_header = (ri == 0)
        for tc in tr.findall('w:tc', NS):
            for p in tc.findall('w:p', NS):
                ppr = p.find('w:pPr', NS)
                if ppr is None:
                    ppr = etree.Element(qn('pPr'))
                    p.insert(0, ppr)
                jc = ppr.find('w:jc', NS)
                if jc is None:
                    jc = etree.SubElement(ppr, qn('jc'))
                jc.set(qn('val'), 'center' if is_header else 'left')
                for r in p.findall('w:r', NS):
                    rpr = r.find('w:rPr', NS)
                    if rpr is None:
                        rpr = etree.Element(qn('rPr'))
                        r.insert(0, rpr)
                    for tag in ('rFonts', 'sz', 'szCs', 'b', 'bCs'):
                        el = rpr.find(f'w:{tag}', NS)
                        if el is not None:
                            rpr.remove(el)
                    rf = etree.SubElement(rpr, qn('rFonts'))
                    rf.set(qn('ascii'), 'Times New Roman')
                    rf.set(qn('hAnsi'), 'Times New Roman')
                    rf.set(qn('cs'), 'Times New Roman')
                    sz = etree.SubElement(rpr, qn('sz'))
                    sz.set(qn('val'), '22')  # 11pt
                    szcs = etree.SubElement(rpr, qn('szCs'))
                    szcs.set(qn('val'), '22')
                    if is_header:
                        b = etree.SubElement(rpr, qn('b'))


def process(input_path, output_path, dry_run=False, report_path=None):
    import tempfile
    workdir = tempfile.mkdtemp(prefix='normalize_')
    with zipfile.ZipFile(input_path) as z:
        z.extractall(workdir)
    for dirpath, dirnames, filenames in os.walk(workdir):
        for fn in filenames:
            full = os.path.join(dirpath, fn)
            if os.path.islink(full):
                os.remove(full)

    inject_result = style_injector.inject(workdir)

    styles_path = os.path.join(workdir, 'word', 'styles.xml')
    styles_root = etree.parse(styles_path).getroot()

    with open(SPEC_PATH, 'r', encoding='utf-8') as f:
        spec = json.load(f)

    doc_path = os.path.join(workdir, 'word', 'document.xml')
    report_lines = [f"Инъекция стилей: {inject_result['status']}"]
    stats = process_document_xml(doc_path, styles_root, report_lines, spec)

    report_lines.insert(1, f"Обычный текст (ДПТ - Основной): {stats['body']}")
    report_lines.insert(2, f"Маркированные списки: {stats['list_bullet']}")
    report_lines.insert(3, f"Нумерованные списки: {stats['list_numbered']}")
    report_lines.insert(4, f"Таблиц переоформлено: {stats['tables']}")
    report_lines.insert(5, f"Заголовки — переприсвоено из уже готового стиля: {stats['heading_from_style']}")
    report_lines.insert(6, f"Заголовки — распознано по тексту (номер вручную): {stats['heading_from_text']}")
    report_lines.insert(7, f"Заголовки — уровень обрезан до {MAX_HEADING_LEVEL} (было глубже): {stats['heading_level_clamped']}")
    report_lines.insert(8, f"Секций (поля страницы применены): {stats['sections']}, пропущено альбомных: {stats['sections_landscape_skipped']}")

    print('\n'.join(report_lines[:9]))

    if report_path:
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(report_lines))

    if not dry_run:
        if os.path.exists(output_path):
            os.remove(output_path)
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for dirpath, dirnames, filenames in os.walk(workdir):
                for fn in filenames:
                    full = os.path.join(dirpath, fn)
                    rel = os.path.relpath(full, workdir)
                    zf.write(full, rel)
        print(f"Сохранено: {output_path}")

    shutil.rmtree(workdir)


def main():
    ap = argparse.ArgumentParser(description="Нормализация структуры .docx (стили, списки, таблицы) по style_spec.json")
    ap.add_argument('input')
    ap.add_argument('output', nargs='?')
    ap.add_argument('--dry-run', action='store_true')
    ap.add_argument('--report')
    args = ap.parse_args()
    process(args.input, args.output, dry_run=args.dry_run, report_path=args.report)


if __name__ == '__main__':
    main()
