const fs = require('fs');
const {
  Document, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, BorderStyle, Packer, LevelFormat,
} = require('docx');

// Генерирует demo_formatting.docx строго по style_spec.json — визуальный
// пример спецификации оформления (заголовки с автонумерацией, оба вида
// списков, таблицы) для проверки пользователем.
// Запуск: node make_demo.js  (требует npm-пакет "docx")

const spec = JSON.parse(fs.readFileSync('style_spec.json', 'utf8'));
const FONT = spec.body.font;
const pt = (p) => p * 2; // half-points
const TABLE_SZ = pt(spec.table.size_pt);

function noteParagraph(text) {
  return new Paragraph({
    children: [new TextRun({ text, italics: true, color: '888888', size: 18, font: FONT })],
    spacing: { before: 200, after: 100 },
  });
}

// ---------------------------------------------------------------------
// Нумерация: заголовки (многоуровневая 1. / 1.1. / 1.1.1. / 1.1.1.1.),
// маркированный список (-), нумерованный список (1))
// Всё — настоящие Word numPr-списки, не вручную напечатанные номера.
// ---------------------------------------------------------------------
const headingLevels = spec.headings.levels.map((lvl, i) => ({
  level: i,
  format: LevelFormat.DECIMAL,
  text: Array.from({ length: i + 1 }, (_, j) => `%${j + 1}`).join('.') + '.',
  alignment: AlignmentType.START,
  style: {
    paragraph: { indent: { left: 0, hanging: 0 } },
    run: { bold: !!lvl.bold, italics: !!lvl.italics, size: pt(lvl.size_pt), font: lvl.font },
  },
}));

const numberingConfig = {
  config: [
    { reference: 'dptHeadingNum', levels: headingLevels },
    {
      reference: 'dptBulletNum',
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: '-',
        alignment: AlignmentType.START,
        style: { paragraph: { indent: { left: spec.list_bullet.indent_twips, hanging: 0 } } },
      }],
    },
    {
      reference: 'dptNumberedNum',
      levels: [{
        level: 0, format: LevelFormat.DECIMAL, text: '%1)',
        alignment: AlignmentType.START,
        style: { paragraph: { indent: { left: spec.list_numbered.indent_twips, hanging: 0 } } },
      }],
    },
  ],
};

function headingParagraph(level, text) {
  const spc = spec.headings.levels[level - 1];
  return new Paragraph({
    outlineLevel: level - 1,
    numbering: { reference: 'dptHeadingNum', level: level - 1 },
    spacing: { before: spc.space_before_pt * 20, after: spc.space_after_pt * 20 },
    children: [new TextRun({
      text, font: spc.font, size: pt(spc.size_pt), bold: !!spc.bold, italics: !!spc.italics,
    })],
  });
}

function bodyParagraph(text) {
  return new Paragraph({
    children: [new TextRun({ text, font: FONT, size: pt(spec.body.size_pt) })],
    alignment: AlignmentType.JUSTIFIED,
    indent: { firstLine: spec.body.first_line_indent_twips },
    spacing: { line: 360, lineRule: 'auto', after: spec.body.space_after_pt * 20 },
  });
}

function bulletParagraph(text) {
  return new Paragraph({
    numbering: { reference: 'dptBulletNum', level: 0 },
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 360, lineRule: 'auto' },
    children: [new TextRun({ text, font: FONT, size: pt(spec.list_bullet.size_pt) })],
  });
}

function numberedParagraph(text) {
  return new Paragraph({
    numbering: { reference: 'dptNumberedNum', level: 0 },
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 360, lineRule: 'auto' },
    children: [new TextRun({ text, font: FONT, size: pt(spec.list_numbered.size_pt) })],
  });
}

const borderSpec = { style: BorderStyle.SINGLE, size: spec.table.border_width_pt * 8, color: '000000' };

function tableCell(text, { header = false, width } = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    margins: spec.table.cell_margin_twips,
    children: [new Paragraph({
      alignment: header ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: [new TextRun({ text, font: FONT, size: TABLE_SZ, bold: header })],
    })],
  });
}

const colWidths = [1600, 3200, 2400, 2400];
const table = new Table({
  width: { size: colWidths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
  columnWidths: colWidths,
  borders: {
    top: borderSpec, bottom: borderSpec, left: borderSpec, right: borderSpec,
    insideHorizontal: borderSpec, insideVertical: borderSpec,
  },
  rows: [
    new TableRow({ children: [
      tableCell('№ точки', { header: true, width: colWidths[0] }),
      tableCell('Наименование', { header: true, width: colWidths[1] }),
      tableCell('X (м)', { header: true, width: colWidths[2] }),
      tableCell('Y (м)', { header: true, width: colWidths[3] }),
    ] }),
    new TableRow({ children: [
      tableCell('1', { width: colWidths[0] }),
      tableCell('ЗУ №1', { width: colWidths[1] }),
      tableCell('70108,51', { width: colWidths[2] }),
      tableCell('113167,36', { width: colWidths[3] }),
    ] }),
    new TableRow({ children: [
      tableCell('2', { width: colWidths[0] }),
      tableCell('ЗУ №1', { width: colWidths[1] }),
      tableCell('70081,23', { width: colWidths[2] }),
      tableCell('113231,21', { width: colWidths[3] }),
    ] }),
  ],
});

const doc = new Document({
  numbering: numberingConfig,
  sections: [{
    properties: {
      page: {
        size: { width: 11907, height: 16840 },
        margin: spec.page.margins_twips,
      },
    },
    children: [
      new Paragraph({
        children: [new TextRun({ text: 'Пример оформления — на проверку', bold: true, size: 28, font: FONT })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 100 },
      }),
      noteParagraph('Все номера ниже — настоящая автонумерация Word (не напечатаны вручную). Попробуйте удалить один заголовок или пункт списка в Word — остальные перенумеруются сами.'),

      headingParagraph(1, 'Исходные данные'),
      bodyParagraph('Настоящим разделом устанавливаются границы зоны планируемого размещения линейного объекта регионального значения в соответствии с утверждённой документацией по планировке территории.'),

      headingParagraph(2, 'Существующее положение'),
      bodyParagraph('Абзац оформлен шрифтом Times New Roman 12 пт, выравнивание по ширине, отступ первой строки 1,25 см, межстрочный интервал полуторный.'),

      headingParagraph(3, 'Инженерное обеспечение'),
      bodyParagraph('Пример того, как выглядит заголовок третьего уровня — при том же размере шрифта, что и второй уровень, но с большей вложенностью номера.'),

      headingParagraph(4, 'Водоснабжение'),
      bodyParagraph('Четвёртый уровень — курсивом, чтобы визуально отличаться при одинаковом размере шрифта.'),

      headingParagraph(2, 'Перечисления'),
      bodyParagraph('Ниже — два реальных вида перечислений, которые встречаются в документах: маркированный список (пункты через дефис) и нумерованный список (пункты вида «1)», «2)»). Оба — настоящие списки Word.'),

      new Paragraph({ children: [new TextRun({ text: 'Маркированный список:', font: FONT, size: pt(12), bold: false })], spacing: { before: 120 } }),
      bulletParagraph('Запрещается размещать объекты, способствующие привлечению и массовому скоплению птиц;'),
      bulletParagraph('Допускается размещать в границах шестой подзоны объекты по обращению с твёрдыми коммунальными отходами при наличии заключения по результатам орнитологического исследования;'),
      bulletParagraph('Требуется согласование с уполномоченным органом в случаях, установленных законодательством.'),

      new Paragraph({ children: [new TextRun({ text: 'Нумерованный список:', font: FONT, size: pt(12), bold: false })], spacing: { before: 200 } }),
      numberedParagraph('для расчёта распределительных сетей водопровода расход на наружное пожаротушение принимается по нормативным документам;'),
      numberedParagraph('для расчёта магистральных (расчётных кольцевых) сетей водопровода расход принимается с учётом одновременного тушения нескольких пожаров.'),

      headingParagraph(2, 'Таблицы'),
      noteParagraph('Рамка 0,5 pt, шапка — полужирный по центру, БЕЗ заливки цветом. Шрифт внутри — Times New Roman 11 пт.'),
      table,
      new Paragraph({ text: '', spacing: { before: 200 } }),
      noteParagraph('Если что-то не так — поправьте прямо в этом Word-файле и пришлите обратно, я сниму параметры и обновлю style_spec.json под них.'),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync('demo_formatting.docx', buf);
  console.log('written demo_formatting.docx');
});
