# -*- coding: utf-8 -*-
"""
Фирменные именованные стили Word (по образцу skill «gradostroitelnaya-spravka»
и generate.js) для единого оформления Коммерческих предложений и
градостроительных справок/анализов.

Две операции, соответствуют двум галочкам в GUI:
  add_styles=True   — ВКЛЕИТЬ определения стилей в word/styles.xml (идемпотентно,
                      по styleId). Существующий текст НЕ переоформляется — стили
                      просто становятся доступны в панели Word и срабатывают там,
                      где документ их уже использует. Низкий риск.
  reformat=True     — ПЕРЕОФОРМИТЬ: первый абзац -> «Заголовок документа»,
                      распознанные заголовки («1. …», «1.1 …», а также абзацы с
                      outline-стилями Word) -> «Заголовок раздела»/«Подзаголовок»,
                      с очисткой прямого форматирования этих абзацев. Тело,
                      таблицы, врезки НЕ трогаются (нет надёжного признака).
                      Экспериментально — проверять глазами.

Значения (кегли в полукеглях, цвета hex) — из generate.js:
  accent 1F4E79, callout FDEBD0/D68910, шрифт Times New Roman.
"""
import os
import re
import shutil
import zipfile
import tempfile
from lxml import etree

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W_NS}
ACCENT = '1F4E79'
FONT = 'Times New Roman'


def qn(tag):
    return '{%s}%s' % (W_NS, tag)


# styleId -> русское имя (для отчёта/справки)
STYLE_NAMES = {
    'DPTgTitle': 'Заголовок документа',
    'DPTgSubtitle': 'Подзаголовок документа',
    'DPTgH1': 'Заголовок раздела',
    'DPTgH2': 'Подзаголовок',
    'DPTgBody': 'Обычный (ДПТ)',
    'DPTgCallout': 'Вывод / Важно',
    'DPTgSource': 'Источник',
}

# Определения стилей одной строкой-фрагментом (оборачивается w:styles при разборе).
_STYLES_FRAGMENT = f'''<w:styles xmlns:w="{W_NS}">
  <w:style w:type="paragraph" w:styleId="DPTgTitle">
    <w:name w:val="Заголовок документа"/>
    <w:qFormat/>
    <w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="120"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/><w:b/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="DPTgSubtitle">
    <w:name w:val="Подзаголовок документа"/>
    <w:qFormat/>
    <w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="200"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/><w:i/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="DPTgH1">
    <w:name w:val="Заголовок раздела"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="DPTgBody"/>
    <w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/><w:b/><w:color w:val="{ACCENT}"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="DPTgH2">
    <w:name w:val="Подзаголовок"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="DPTgBody"/>
    <w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="160" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/><w:b/><w:sz w:val="23"/><w:szCs w:val="23"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="DPTgBody">
    <w:name w:val="Обычный (ДПТ)"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr><w:jc w:val="both"/><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="DPTgCallout">
    <w:name w:val="Вывод / Важно"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="D68910"/></w:pBdr>
      <w:shd w:val="clear" w:color="auto" w:fill="FDEBD0"/>
      <w:spacing w:before="120" w:after="120"/><w:ind w:left="120"/>
    </w:pPr>
    <w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/><w:b/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="DPTgSource">
    <w:name w:val="Источник"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr><w:spacing w:before="200" w:after="0"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/><w:i/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
  </w:style>
</w:styles>'''


def _inject_styles(styles_root):
    """Добавляет фирменные стили в w:styles. Если стиль с таким ИМЕНЕМ уже есть
    (напр. документ создан скиллом-генератором) — переиспользуем его, а не
    добавляем дубликат (иначе Word ругается на одинаковые имена стилей).
    Возвращает (число добавленных, mapping: наш styleId -> фактический styleId)."""
    existing_ids = {st.get(qn('styleId')) for st in styles_root.findall('w:style', NS)}
    existing_by_name = {}
    for st in styles_root.findall('w:style', NS):
        nm = st.find('w:name', NS)
        if nm is not None:
            existing_by_name[nm.get(qn('val'))] = st.get(qn('styleId'))

    frag = etree.fromstring(_STYLES_FRAGMENT.encode('utf-8'))
    added = 0
    mapping = {}
    for st in frag.findall('w:style', NS):
        sid = st.get(qn('styleId'))
        nm = st.find('w:name', NS).get(qn('val'))
        if nm in existing_by_name:
            mapping[sid] = existing_by_name[nm]      # уже есть по имени — используем его
        elif sid in existing_ids:
            mapping[sid] = sid                       # id занят (имя другое) — маловероятно
        else:
            styles_root.append(st)
            added += 1
            mapping[sid] = sid
    return added, mapping


# --- переоформление ---

HEADING1_RE = re.compile(r'^\d{1,2}\.\s+\S')
HEADING2_RE = re.compile(r'^\d{1,2}\.\d{1,2}\.?\s+\S')


def _para_text(p):
    return ''.join((t.text or '') for t in p.findall('.//w:t', NS))


def _set_pstyle(p, style_id):
    ppr = p.find('w:pPr', NS)
    if ppr is None:
        ppr = etree.Element(qn('pPr'))
        p.insert(0, ppr)
    pstyle = ppr.find('w:pStyle', NS)
    if pstyle is None:
        pstyle = etree.SubElement(ppr, qn('pStyle'))
        ppr.insert(0, pstyle)
    pstyle.set(qn('val'), style_id)


def _clear_heading_run_formatting(p):
    """Убирает прямое форматирование run'ов заголовка, чтобы вид определялся
    стилем (цвет/кегль/жирность из стиля, а не из прямого форматирования)."""
    for r in p.findall('w:r', NS):
        rpr = r.find('w:rPr', NS)
        if rpr is None:
            continue
        for tag in ('b', 'bCs', 'i', 'iCs', 'color', 'sz', 'szCs', 'u', 'caps', 'smallCaps', 'highlight'):
            for el in rpr.findall('w:' + tag, NS):
                rpr.remove(el)


def _resolve_outline_level(style_id, styles_root, cache):
    """Уровень outline по цепочке basedOn (как в normalize_structure)."""
    if style_id in cache:
        return cache[style_id]
    sid = style_id
    seen = set()
    result = None
    while sid and sid not in seen:
        seen.add(sid)
        s = next((st for st in styles_root.findall('w:style', NS)
                  if st.get(qn('styleId')) == sid), None)
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


def _reformat(root, styles_root, mapping):
    """Переоформляет заголовки/титул под фирменные стили. mapping — наш styleId
    -> фактический (на случай, если стиль уже есть по имени). Возвращает
    статистику. Осторожно: только абзацы верхнего уровня тела (не в таблицах)."""
    def sid(key):
        return mapping.get(key, key)
    body = root.find('w:body', NS)
    if body is None:
        return {'title': 0, 'h1': 0, 'h2': 0}
    stats = {'title': 0, 'h1': 0, 'h2': 0}
    cache = {}
    top_paras = [c for c in body if c.tag == qn('p')]
    title_done = False
    for idx, p in enumerate(top_paras):
        txt = _para_text(p).strip()
        if not txt:
            continue

        # первый непустой абзац -> заголовок документа
        if not title_done:
            title_done = True
            # титул только если короткий (не абзац-простыня)
            if len(txt) <= 160:
                _set_pstyle(p, sid('DPTgTitle'))
                _clear_heading_run_formatting(p)
                stats['title'] += 1
                continue

        # существующий outline-стиль документа
        ppr = p.find('w:pPr', NS)
        cur = ppr.find('w:pStyle', NS) if ppr is not None else None
        lvl = None
        if cur is not None:
            lvl = _resolve_outline_level(cur.get(qn('val')), styles_root, cache)

        target = None
        if HEADING2_RE.match(txt) and len(txt) <= 160:
            target = 'DPTgH2'
        elif HEADING1_RE.match(txt) and len(txt) <= 160:
            target = 'DPTgH1'
        elif lvl is not None:
            target = 'DPTgH1' if lvl == 0 else 'DPTgH2'

        if target:
            _set_pstyle(p, sid(target))
            _clear_heading_run_formatting(p)
            stats['h1' if target == 'DPTgH1' else 'h2'] += 1
    return stats


def apply(input_path, output_path, add_styles=True, reformat=False):
    """Вклеивает стили и/или переоформляет. Возвращает статистику."""
    workdir = tempfile.mkdtemp(prefix='dpt_grad_')
    stats = {'styles_added': 0, 'title': 0, 'h1': 0, 'h2': 0}
    try:
        with zipfile.ZipFile(input_path) as z:
            z.extractall(workdir)
        for dirpath, dirnames, filenames in os.walk(workdir):
            for fn in filenames:
                full = os.path.join(dirpath, fn)
                if os.path.islink(full):
                    os.remove(full)

        styles_path = os.path.join(workdir, 'word', 'styles.xml')
        styles_tree = etree.parse(styles_path)
        styles_root = styles_tree.getroot()

        # reformat требует наличия стилей -> вклеиваем в обоих случаях
        added, mapping = _inject_styles(styles_root)
        stats['styles_added'] = added
        styles_tree.write(styles_path, xml_declaration=True, encoding='UTF-8', standalone=True)

        if reformat:
            doc_path = os.path.join(workdir, 'word', 'document.xml')
            doc_tree = etree.parse(doc_path)
            rstats = _reformat(doc_tree.getroot(), styles_root, mapping)
            stats.update(rstats)
            doc_tree.write(doc_path, xml_declaration=True, encoding='UTF-8', standalone=True)

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
