const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, TabStopType,
  TabStopPosition, UnderlineType
} = require('docx');
const fs = require('fs');

// ── DATA (injected by caller) ─────────────────────────────
const DATA = JSON.parse(process.argv[2]);
const {
  proyecto, codigoProyecto, ordenCompra, fecha, ordenTrabajo,
  actividades, hallazgos, fotos, // fotos = [{legend, b64}]
  situacionGeneral, cumplimientoCronograma, avanceProyecto, accionesRequeridas,
  avance, supervisor, supervisorIniciales, codigo
} = DATA;

// ── HELPERS ───────────────────────────────────────────────
const FONT = 'Segoe UI';
const BLUE = '1F3864';
const W_PAGE = 11906;  // A4
const M_LEFT = 1701; const M_RIGHT = 1701;
const CONTENT_W = W_PAGE - M_LEFT - M_RIGHT; // 8504 DXA

function tx(text, opts={}) {
  return new TextRun({ text, font: FONT, ...opts });
}
function para(children, opts={}) {
  return new Paragraph({ children: Array.isArray(children) ? children : [children], ...opts });
}
function emptyPara() {
  return para([tx('')], { spacing: { before: 0, after: 80 } });
}

// ── THIN BORDER ───────────────────────────────────────────
const thinBorder = { style: BorderStyle.SINGLE, size: 4, color: 'BBBBBB' };
const cellBorders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// ── LOGO ──────────────────────────────────────────────────
const logoData = fs.readFileSync('/home/claude/logo.jpg');

// ── HEADER TABLE: MPE box + Sistema de Reportes ───────────
// Blue box with MPE white text + subtitle below
const headerTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [CONTENT_W],
  borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder,
             insideH: noBorder, insideV: noBorder },
  rows: [
    new TableRow({
      children: [
        new TableCell({
          borders: noBorders,
          width: { size: CONTENT_W, type: WidthType.DXA },
          shading: { fill: BLUE, type: ShadingType.CLEAR },
          margins: { top: 120, bottom: 80, left: 160, right: 160 },
          children: [
            para([
              tx('MPE', { bold: true, size: 40, color: 'FFFFFF' }),
              tx('   Sistema de Reportes de Visita', { size: 22, color: 'FFFFFF' })
            ], { spacing: { before: 0, after: 0 } })
          ]
        })
      ]
    })
  ]
});

// ── IDENTIFICATION TABLE ──────────────────────────────────
function idRow(label, value) {
  return new TableRow({
    children: [
      new TableCell({
        borders: cellBorders,
        width: { size: 2800, type: WidthType.DXA },
        shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        children: [para([tx(label, { bold: true, size: 18 })], { spacing: { before: 0, after: 0 } })]
      }),
      new TableCell({
        borders: cellBorders,
        width: { size: CONTENT_W - 2800, type: WidthType.DXA },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        children: [para([tx(value || '—', { size: 18 })], { spacing: { before: 0, after: 0 } })]
      })
    ]
  });
}

const idTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [2800, CONTENT_W - 2800],
  rows: [
    idRow('Proyecto', proyecto),
    idRow('Código Proyecto', codigoProyecto),
    idRow('Orden de Compra', ordenCompra),
    idRow('Fecha', fecha),
    idRow('Avance estimado', avance + '%'),
    idRow('Calendario', cumplimientoCronograma),
    idRow('Situación', situacionGeneral),
    idRow('Supervisor', supervisor + ' (' + supervisorIniciales + ')'),
  ]
});

// ── SUBTITLE ──────────────────────────────────────────────
const subtitlePara = para(
  [tx('Reporte Técnico de Visita de Supervisión', { bold: true, size: 24 })],
  { alignment: AlignmentType.CENTER, spacing: { before: 280, after: 120 } }
);

// ── MINI-FICHA ────────────────────────────────────────────
function miniRow(label, value) {
  return para([
    tx(label + ' :  ', { bold: false, size: 20 }),
    tx(value || '—', { size: 20 })
  ], { spacing: { before: 0, after: 40 } });
}

// ── SECTION TITLE ─────────────────────────────────────────
function secTitle(text) {
  return para([tx(text, { bold: true, size: 22 })],
    { spacing: { before: 200, after: 100 } });
}

// ── SUBSECTION TITLE ──────────────────────────────────────
function subSecTitle(text) {
  return para([tx(text, { bold: true, size: 20 })],
    { spacing: { before: 160, after: 60 } });
}

// ── BODY TEXT ─────────────────────────────────────────────
function bodyText(text) {
  return para([tx(text, { size: 20 })],
    { alignment: AlignmentType.BOTH, spacing: { before: 0, after: 80 },
      indent: { left: 720 } });
}

// ── ACTIVIDADES as roman numeral list ─────────────────────
function actividadItem(text, num) {
  const roman = ['i','ii','iii','iv','v','vi','vii','viii','ix','x'];
  const label = roman[num] || (num+1).toString();
  return para([
    tx(label + '.', { size: 20 }),
    tx('    ' + text, { size: 20 })
  ], {
    alignment: AlignmentType.BOTH,
    spacing: { before: 40, after: 60 },
    indent: { left: 720, hanging: 360 }
  });
}

// ── PHOTOS in 2-column table ──────────────────────────────
function buildPhotoTable(fotosArr) {
  const rows = [];
  for (let i = 0; i < fotosArr.length; i += 2) {
    const f1 = fotosArr[i];
    const f2 = fotosArr[i + 1];
    const colW = Math.floor(CONTENT_W / 2) - 100;
    const imgW = 220; const imgH = 160;

    function photoCell(f) {
      if (!f) return new TableCell({
        borders: noBorders,
        width: { size: colW, type: WidthType.DXA },
        children: [emptyPara()]
      });

      let imgData, imgType;
      try {
        const b64 = f.b64.split(',')[1];
        imgData = Buffer.from(b64, 'base64');
        imgType = f.b64.startsWith('data:image/png') ? 'png' : 'jpg';
      } catch(e) {
        imgData = fs.readFileSync('/home/claude/image1.png');
        imgType = 'png';
      }

      return new TableCell({
        borders: noBorders,
        width: { size: colW, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 80, right: 80 },
        children: [
          para([new ImageRun({
            type: imgType,
            data: imgData,
            transformation: { width: imgW, height: imgH },
            altText: { title: f.legend||'', description: f.legend||'', name: f.legend||'' }
          })], { alignment: AlignmentType.CENTER, spacing: { before: 0, after: 60 } }),
          para([tx(f.legend || '', { size: 16 })],
            { alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 } })
        ]
      });
    }

    rows.push(new TableRow({
      children: [photoCell(f1), photoCell(f2)]
    }));
  }

  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [Math.floor(CONTENT_W/2)-100, Math.floor(CONTENT_W/2)-100],
    borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder,
               insideH: noBorder, insideV: noBorder },
    rows
  });
}

// ── FOOTER ────────────────────────────────────────────────
// codigo left, Pág X/Y right — same line using tab stop
const footer = new Footer({
  children: [
    para([
      tx(codigo, { size: 16, color: '666666' }),
      new TextRun({ text: '\t', font: FONT }),
      tx('Pág ', { size: 16, color: '666666' }),
      new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: '666666' }),
      tx(' / ', { size: 16, color: '666666' }),
      new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 16, color: '666666' }),
    ], {
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'BBBBBB', space: 4 } },
      tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
      spacing: { before: 80, after: 0 }
    })
  ]
});

// ── ACTIVIDADES items ─────────────────────────────────────
const actLines = (actividades || '').split('\n').filter(l => l.trim());
const actItems = actLines.length
  ? actLines.map((l, i) => actividadItem(l.trim(), i))
  : [bodyText('No se ingresó información.')];

// ── HALLAZGOS ─────────────────────────────────────────────
const hallLines = (hallazgos || '').split('\n').filter(l => l.trim());
const hallItems = hallLines.length
  ? hallLines.map(l => bodyText(l.trim()))
  : [bodyText('No se registraron hallazgos significativos durante la presente visita de supervisión.')];

// ── FOTOS section ─────────────────────────────────────────
const fotosSection = [];
if (fotos && fotos.length) {
  fotosSection.push(buildPhotoTable(fotos));
} else {
  fotosSection.push(bodyText('No se registraron fotografías en esta visita.'));
}

// ── COMENTARIOS subsections ───────────────────────────────
const accionLines = (accionesRequeridas || '').split('\n').filter(l => l.trim());

// ── BUILD DOCUMENT ────────────────────────────────────────
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: FONT, size: 20 } }
    }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1417, right: M_RIGHT, bottom: 1417, left: M_LEFT,
                  header: 708, footer: 708, gutter: 0 }
      }
    },
    footers: { default: footer },
    children: [
      // HEADER BLOCK
      headerTable,
      emptyPara(),

      // IDENTIFICATION TABLE
      idTable,
      emptyPara(),

      // SUBTITLE
      subtitlePara,

      // MINI-FICHA
      miniRow('Proyecto', proyecto),
      miniRow('Orden Trabajo', ordenTrabajo || ordenCompra),
      miniRow('Fecha Visita', fecha),
      emptyPara(),

      // 1.0 ACTIVIDADES
      secTitle('1.0 Actividades Realizadas'),
      ...actItems,
      emptyPara(),

      // 2.0 HALLAZGOS
      secTitle('2.0 Hallazgos'),
      ...hallItems,
      emptyPara(),

      // 3.0 FOTOS
      secTitle('3.0 Registro Fotográfico'),
      ...fotosSection,
      emptyPara(),

      // 4.0 COMENTARIOS
      secTitle('4.0 Comentarios y Acciones'),

      subSecTitle('4.1 Situación General'),
      bodyText(situacionGeneral || 'No se registran observaciones críticas en la presente etapa.'),

      subSecTitle('4.2 Cumplimiento de Cronograma'),
      bodyText(cumplimientoCronograma || 'Las actividades se desarrollan dentro del cronograma.'),

      subSecTitle('4.3 Avance del Proyecto'),
      bodyText('Avance físico estimado ' + avance + '%' + (avanceProyecto ? ' - ' + avanceProyecto : '.')),

      subSecTitle('4.4 Acciones Requeridas'),
      ...(accionLines.length
        ? accionLines.map(l => bodyText(l.trim()))
        : [bodyText('No se identifican acciones inmediatas.')]),

      emptyPara(),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  const outPath = '/home/claude/reporte_output.docx';
  fs.writeFileSync(outPath, buf);
  console.log('OK:' + outPath);
}).catch(e => {
  console.error('ERR:' + e.message);
  process.exit(1);
});
