# -*- coding: utf-8 -*-
"""
Вклеивает провалидированные стили «ДПТ - ...» (dpt_styles_fragment.xml) и их
нумерацию (dpt_numbering_fragment.xml) в word/styles.xml и word/numbering.xml
целевого документа. Эти фрагменты извлечены из demo_formatting.docx,
сгенерированного make_demo.js — то есть это ровно те XML-определения,
которые уже проверены визуально в Word, а не заново собранная в Python логика.

Идемпотентно: повторный запуск на уже обработанном документе не создаёт
дублей (проверяет наличие styleId "DPTBody" и пропускает инъекцию, если он
уже есть).
"""
import os
from lxml import etree

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W_NS}


def qn(tag):
    return '{%s}%s' % (W_NS, tag)


HERE = os.path.dirname(__file__)
STYLES_FRAGMENT_PATH = os.path.join(HERE, 'dpt_styles_fragment.xml')
NUMBERING_FRAGMENT_PATH = os.path.join(HERE, 'dpt_numbering_fragment.xml')


def _load_fragment_elements(path, root_tag='root'):
    """dpt_*_fragment.xml — это последовательность элементов без общего корня,
    оборачиваем во временный корень для парсинга."""
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    wrapper = f'<{root_tag} xmlns:w="{W_NS}">{content}</{root_tag}>'
    root = etree.fromstring(wrapper.encode('utf-8'))
    return list(root)


def inject(workdir):
    """workdir — распакованный .docx (содержит word/styles.xml, word/numbering.xml)."""
    styles_path = os.path.join(workdir, 'word', 'styles.xml')
    numbering_path = os.path.join(workdir, 'word', 'numbering.xml')

    styles_tree = etree.parse(styles_path)
    styles_root = styles_tree.getroot()

    # идемпотентность: если уже вклеено — ничего не делаем
    existing = [s for s in styles_root.findall('w:style', NS)
                if s.get(qn('styleId')) == 'DPTBody']
    if existing:
        return {'status': 'already-injected'}

    if not os.path.exists(numbering_path):
        raise RuntimeError(
            f'{numbering_path} не найден — у этого документа нет word/numbering.xml. '
            'Нужно создать часть numbering.xml с нуля (не реализовано, в исследованных '
            '4 документах пользователя такого случая не встретилось).'
        )

    numbering_tree = etree.parse(numbering_path)
    numbering_root = numbering_tree.getroot()

    # --- 1. Определяем сдвиг ID, чтобы не столкнуться с существующей нумерацией ---
    existing_abstract_ids = [int(a.get(qn('abstractNumId')))
                              for a in numbering_root.findall('w:abstractNum', NS)]
    existing_num_ids = [int(n.get(qn('numId')))
                         for n in numbering_root.findall('w:num', NS)]
    abs_offset = (max(existing_abstract_ids) + 1) if existing_abstract_ids else 0
    num_offset = (max(existing_num_ids) + 1) if existing_num_ids else 0

    # --- 2. Загружаем наши фрагменты нумерации и переприсваиваем ID ---
    numbering_elems = _load_fragment_elements(NUMBERING_FRAGMENT_PATH)
    abstract_id_map = {}  # старый -> новый
    num_id_map = {}

    abstract_elems = [e for e in numbering_elems if e.tag == qn('abstractNum')]
    num_elems = [e for e in numbering_elems if e.tag == qn('num')]

    for a in abstract_elems:
        old_id = a.get(qn('abstractNumId'))
        new_id = str(int(old_id) + abs_offset)
        abstract_id_map[old_id] = new_id
        a.set(qn('abstractNumId'), new_id)

    for n in num_elems:
        old_id = n.get(qn('numId'))
        new_id = str(int(old_id) + num_offset)
        num_id_map[old_id] = new_id
        n.set(qn('numId'), new_id)
        abs_ref = n.find('w:abstractNumId', NS)
        abs_ref.set(qn('val'), abstract_id_map[abs_ref.get(qn('val'))])

    for e in abstract_elems + num_elems:
        numbering_root.append(e)

    # --- 3. Загружаем наши стили, переприсваиваем ссылки numId на новые ID ---
    style_elems = _load_fragment_elements(STYLES_FRAGMENT_PATH)
    for s in style_elems:
        for numpr in s.findall('.//w:numPr', NS):
            numid_el = numpr.find('w:numId', NS)
            if numid_el is not None:
                old = numid_el.get(qn('val'))
                if old in num_id_map:
                    numid_el.set(qn('val'), num_id_map[old])
        styles_root.append(s)

    styles_tree.write(styles_path, xml_declaration=True, encoding='UTF-8', standalone=True)
    numbering_tree.write(numbering_path, xml_declaration=True, encoding='UTF-8', standalone=True)

    return {
        'status': 'injected',
        'abstract_id_map': abstract_id_map,
        'num_id_map': num_id_map,
    }
