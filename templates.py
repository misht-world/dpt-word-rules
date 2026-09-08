# -*- coding: utf-8 -*-
"""
Шаблоны оформления .docx — консервативный слой ПОВЕРХ типографики.

Идея: пользователь выбирает шаблон (напр. «Коммерческое предложение»), и к
документу применяется единое оформление СТРАНИЦЫ:
  - поля страницы (только для книжных/portrait секций);
  - колонтитулы: шапка (текст) и нижний колонтитул с номером страницы.

СОЗНАТЕЛЬНО НЕ ТРОГАЕМ заголовки/списки/таблицы/стили абзацев — это самая
хрупкая часть (см. normalize_structure.py, из-за которой структура отключена в
GUI). Здесь только поля и колонтитулы — низкий риск сломать документ.

Значения полей — в твипах (1 см = 567 твипов). Шаблоны можно править ниже —
это обычный словарь TEMPLATES.

Использование:
    import templates
    templates.apply_template('in.docx', 'out.docx', 'Коммерческое предложение')
"""
import os
import re
import shutil
import zipfile
import tempfile
from lxml import etree

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
PR_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS = {'w': W_NS, 'r': R_NS}

CM = 567  # твипов в сантиметре


def qn(tag):
    return '{%s}%s' % (W_NS, tag)


def rqn(tag):
    return '{%s}%s' % (R_NS, tag)


# ---------------------------------------------------------------------------
# Шаблоны. Правь смело — это просто данные.
# margins: top/bottom/left/right/header/footer в твипах (или None — не менять).
# header:  текст шапки (или None — без шапки).
# footer_page: True -> нижний колонтитул с номером страницы по центру.
#
# ЕДИНЫЙ СТИЛЬ: все шаблоны используют общую базу полей _BASE_MARGINS (совпадает
# с реальными документами — «Коммерческое предложение» и «Градостроительная
# справка»: A4, поля 2/1.5/2/3 см). Различается только текст шапки. Так
# «Коммерческое» и «Градостроительный анализ» выглядят единообразно.
# ---------------------------------------------------------------------------

# поля 2 см сверху/снизу, 3 см слева (под подшивку), 1.5 см справа; колонтитулы 1.25 см
_BASE_MARGINS = {'top': 2 * CM, 'bottom': 2 * CM, 'left': 3 * CM, 'right': int(1.5 * CM),
                 'header': int(1.25 * CM), 'footer': int(1.25 * CM)}

TEMPLATES = {
    'Коммерческое предложение': {
        'margins': _BASE_MARGINS,
        'header': 'Коммерческое предложение',
        'footer_page': True,
    },
    'Градостроительный анализ': {
        'margins': _BASE_MARGINS,
        'header': 'Градостроительный анализ',
        'footer_page': True,
    },
}


def template_names():
    return list(TEMPLATES.keys())


# ---------------------------------------------------------------------------
# Заготовки XML колонтитулов
# ---------------------------------------------------------------------------

def _header_xml(text):
    safe = (text or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<w:hdr xmlns:w="%s">'
        '<w:p><w:pPr><w:jc w:val="right"/></w:pPr>'
        '<w:r><w:rPr><w:i/><w:color w:val="808080"/><w:sz w:val="18"/></w:rPr>'
        '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
        '</w:hdr>' % (W_NS, safe)
    )


def _footer_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<w:ftr xmlns:w="%s">'
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        '</w:p></w:ftr>' % W_NS
    )


# ---------------------------------------------------------------------------
# Работа с пакетом
# ---------------------------------------------------------------------------

def _is_landscape(sectpr):
    pgsz = sectpr.find('w:pgSz', NS)
    if pgsz is None:
        return False
    if pgsz.get(qn('orient')) == 'landscape':
        return True
    w = pgsz.get(qn('w'))
    h = pgsz.get(qn('h'))
    try:
        return int(w) > int(h)
    except (TypeError, ValueError):
        return False


def _set_margins(sectpr, margins):
    pgmar = sectpr.find('w:pgMar', NS)
    if pgmar is None:
        pgmar = etree.SubElement(sectpr, qn('pgMar'))
    for key in ('top', 'bottom', 'left', 'right', 'header', 'footer'):
        val = margins.get(key)
        if val is not None:
            pgmar.set(qn(key), str(int(val)))
    if pgmar.get(qn('gutter')) is None:
        pgmar.set(qn('gutter'), '0')


def _ensure_relationship(rels_root, rel_type, target):
    """Возвращает Id связи на target; создаёт при отсутствии. Идемпотентно."""
    for rel in rels_root:
        if rel.get('Target') == target and rel.get('Type') == rel_type:
            return rel.get('Id')
    nums = [int(m.group(1)) for rel in rels_root
            for m in [re.match(r'rId(\d+)$', rel.get('Id') or '')] if m]
    new_id = 'rId%d' % ((max(nums) + 1) if nums else 1)
    rel = etree.SubElement(rels_root, '{%s}Relationship' % PR_NS)
    rel.set('Id', new_id)
    rel.set('Type', rel_type)
    rel.set('Target', target)
    return new_id


def _ensure_content_type(ct_root, partname, content_type):
    for ov in ct_root.findall('{%s}Override' % CT_NS):
        if ov.get('PartName') == partname:
            return
    ov = etree.SubElement(ct_root, '{%s}Override' % CT_NS)
    ov.set('PartName', partname)
    ov.set('ContentType', content_type)


def _clear_refs(sectpr, tag):
    for ref in sectpr.findall('w:' + tag, NS):
        sectpr.remove(ref)


def _add_ref_first(sectpr, tag, rid):
    """Вставляет headerReference/footerReference первым ребёнком sectPr
    (порядок элементов в CT_SectPr важен — ссылки идут в начале)."""
    ref = etree.Element(qn(tag))
    ref.set(qn('type'), 'default')
    ref.set(rqn('id'), rid)
    sectpr.insert(0, ref)


def apply_template(input_path, output_path, template_name):
    """Применяет шаблон оформления. Возвращает dict со статистикой.
    Бросает KeyError, если шаблон неизвестен."""
    spec = TEMPLATES[template_name]
    workdir = tempfile.mkdtemp(prefix='dpt_tpl_')
    stats = {'sections': 0, 'landscape_skipped': 0, 'header': False, 'footer': False}
    try:
        with zipfile.ZipFile(input_path) as z:
            z.extractall(workdir)
        for dirpath, dirnames, filenames in os.walk(workdir):
            for fn in filenames:
                full = os.path.join(dirpath, fn)
                if os.path.islink(full):
                    os.remove(full)

        doc_path = os.path.join(workdir, 'word', 'document.xml')
        tree = etree.parse(doc_path)
        root = tree.getroot()

        header_id = footer_id = None
        want_header = bool(spec.get('header'))
        want_footer = bool(spec.get('footer_page'))

        if want_header or want_footer:
            # части колонтитулов
            if want_header:
                with open(os.path.join(workdir, 'word', 'header1.xml'), 'w', encoding='utf-8') as f:
                    f.write(_header_xml(spec['header']))
            if want_footer:
                with open(os.path.join(workdir, 'word', 'footer1.xml'), 'w', encoding='utf-8') as f:
                    f.write(_footer_xml())

            # relationships
            rels_path = os.path.join(workdir, 'word', '_rels', 'document.xml.rels')
            rels_tree = etree.parse(rels_path)
            rels_root = rels_tree.getroot()
            if want_header:
                header_id = _ensure_relationship(
                    rels_root, R_NS + '/header', 'header1.xml')
            if want_footer:
                footer_id = _ensure_relationship(
                    rels_root, R_NS + '/footer', 'footer1.xml')
            rels_tree.write(rels_path, xml_declaration=True, encoding='UTF-8', standalone=True)

            # content types
            ct_path = os.path.join(workdir, '[Content_Types].xml')
            ct_tree = etree.parse(ct_path)
            ct_root = ct_tree.getroot()
            if want_header:
                _ensure_content_type(ct_root, '/word/header1.xml',
                                     'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml')
            if want_footer:
                _ensure_content_type(ct_root, '/word/footer1.xml',
                                     'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml')
            ct_tree.write(ct_path, xml_declaration=True, encoding='UTF-8', standalone=True)

        # секции: поля (только portrait) + ссылки на колонтитулы (все)
        for sectpr in root.findall('.//w:sectPr', NS):
            stats['sections'] += 1
            if spec.get('margins'):
                if _is_landscape(sectpr):
                    stats['landscape_skipped'] += 1
                else:
                    _set_margins(sectpr, spec['margins'])
            if want_header:
                _clear_refs(sectpr, 'headerReference')
            if want_footer:
                _clear_refs(sectpr, 'footerReference')
            # порядок: сначала footer, потом header вставляем первым -> header окажется перед footer
            if want_footer and footer_id:
                _add_ref_first(sectpr, 'footerReference', footer_id)
                stats['footer'] = True
            if want_header and header_id:
                _add_ref_first(sectpr, 'headerReference', header_id)
                stats['header'] = True

        tree.write(doc_path, xml_declaration=True, encoding='UTF-8', standalone=True)

        if os.path.exists(output_path):
            os.remove(output_path)
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for dirpath, dirnames, filenames in os.walk(workdir):
                for fn in filenames:
                    full = os.path.join(dirpath, fn)
                    rel = os.path.relpath(full, workdir)
                    zf.write(full, rel)
    finally:
        shutil.rmtree(workdir, ignore_errors=True)
    return stats
