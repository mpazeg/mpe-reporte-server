const express = require('express');
const cors = require('cors');
const JSZip = require('jszip');
const { generatePdfBuffer } = require('./pdf_generator');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Footer, AlignmentType, BorderStyle, WidthType,
  ShadingType, PageNumber, TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

const FONT = 'Segoe UI';
const BLUE = '1A3A6B';
const W_PAGE = 11906;
const M_LEFT = 1701;
const M_RIGHT = 1701;
const CONTENT_W = W_PAGE - M_LEFT - M_RIGHT;

function tx(text, opts={}) {
  return new TextRun({ text: String(text||''), font: FONT, ...opts });
}
function para(children, opts={}) {
  return new Paragraph({ children: Array.isArray(children)?children:[children], ...opts });
}
function emptyPara(space=100) {
  return para([tx('')], { spacing: { before:0, after:space } });
}

const dottedBorder  = { style: BorderStyle.DOTTED, size: 4, color: '767171' };
const dottedBorders = { top:dottedBorder, bottom:dottedBorder, left:dottedBorder, right:dottedBorder };
const noBorder      = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders     = { top:noBorder, bottom:noBorder, left:noBorder, right:noBorder };

function cleanLine(line) {
  return (line||'')
    .replace(/^[\s\u00a0]*(([ivxlcdmIVXLCDM]+|[0-9]+|[a-zA-Z])[\.\)\-]\s*)+/, '')
    .replace(/^[\s\u00a0]*[-•·*▪▸→►]\s*/, '')
    .replace(/^[\s\u00a0]+/, '')
    .trim();
}

const ROMAN = ['i','ii','iii','iv','v','vi','vii','viii','ix','x',
               'xi','xii','xiii','xiv','xv','xvi','xvii','xviii','xix','xx'];

function secTitle(text) {
  return para([tx(text, { bold:true, size:22 })], { spacing:{ before:240, after:100 } });
}
function subSecTitle(text) {
  return para([tx(text, { bold:true, size:20 })], { spacing:{ before:180, after:60 } });
}
function bodyText(text) {
  return para([tx(text, { size:20 })], {
    alignment: AlignmentType.BOTH, spacing: { before:0, after:80 }, indent: { left:720 }
  });
}
function actividadItem(rawText, num) {
  const label = ROMAN[num] || String(num+1);
  const text  = cleanLine(rawText);
  return para([tx(label+'.', { size:20 }), tx('\t'+text, { size:20 })], {
    alignment: AlignmentType.BOTH, spacing: { before:40, after:60 },
    indent: { left:900, hanging:360 },
    tabStops: [{ type:TabStopType.LEFT, position:540 }]
  });
}
function idRow(label, value) {
  return new TableRow({ children:[
    new TableCell({
      borders: dottedBorders, width: { size:3397, type:WidthType.DXA },
      shading: { fill:'EDF2FA', type:ShadingType.CLEAR },
      margins: { top:15, bottom:15, left:120, right:120 },
      children: [para([tx(label, { bold:true, size:18 })], { spacing:{ before:0, after:0 } })]
    }),
    new TableCell({
      borders: dottedBorders, width: { size:CONTENT_W-3397, type:WidthType.DXA },
      margins: { top:15, bottom:15, left:120, right:120 },
      children: [para([tx(String(value||'—'), { size:18 })], { spacing:{ before:0, after:0 } })]
    })
  ]});
}
function buildPhotoTable(fotosArr) {
  const colW = Math.floor(CONTENT_W/2)-60;
  const MAX_W = 215;
  const MAX_H = 170;
  const rows = [];
  for (let i=0; i<fotosArr.length; i+=2) {
    const f1=fotosArr[i], f2=fotosArr[i+1]||null;
    function photoCell(f) {
      if (!f) return new TableCell({ borders:noBorders, width:{size:colW,type:WidthType.DXA}, children:[emptyPara()] });
      let imgData, imgType, imgW=MAX_W, imgH=MAX_H;
      try {
        imgData=Buffer.from(f.b64.split(',')[1],'base64');
        imgType=f.b64.startsWith('data:image/png')?'png':'jpg';
        // Read actual dimensions to preserve aspect ratio
        try {
          let w=0, h=0;
          if (imgType==='jpg') {
            for (let pos=0; pos<imgData.length-8; pos++) {
              if (imgData[pos]===0xFF&&(imgData[pos+1]===0xC0||imgData[pos+1]===0xC2)) {
                h=(imgData[pos+5]<<8)|imgData[pos+6];
                w=(imgData[pos+7]<<8)|imgData[pos+8];
                break;
              }
            }
          } else {
            w=(imgData[16]<<24)|(imgData[17]<<16)|(imgData[18]<<8)|imgData[19];
            h=(imgData[20]<<24)|(imgData[21]<<16)|(imgData[22]<<8)|imgData[23];
          }
          if (w>0&&h>0) {
            const ratio=w/h;
            imgW=MAX_W; imgH=Math.round(MAX_W/ratio);
            if (imgH>MAX_H){imgH=MAX_H;imgW=Math.round(MAX_H*ratio);}
          }
        } catch(ex){}
      } catch(e) {
        return new TableCell({ borders:noBorders, width:{size:colW,type:WidthType.DXA},
          children:[para([tx(f.legend||'',{size:16})],{alignment:AlignmentType.CENTER})] });
      }
      const legend=(f.legend||'').trim();
      return new TableCell({
        borders:noBorders, width:{size:colW,type:WidthType.DXA},
        margins:{top:80,bottom:80,left:60,right:60},
        children:[
          para([new ImageRun({type:imgType,data:imgData,
            transformation:{width:imgW,height:imgH},
            altText:{title:legend||'foto',description:legend,name:legend||'foto'}})],
            {alignment:AlignmentType.CENTER,spacing:{before:0,after:40}}),
          para([tx(legend,{size:16})],{alignment:AlignmentType.CENTER,spacing:{before:0,after:80}})
        ]
      });
    }
    rows.push(new TableRow({ children:[photoCell(f1),photoCell(f2)] }));
  }
  return new Table({
    width:{size:CONTENT_W,type:WidthType.DXA}, columnWidths:[colW,colW],
    borders:{top:noBorder,bottom:noBorder,left:noBorder,right:noBorder,insideH:noBorder,insideV:noBorder},
    rows
  });
}

function makeVerticalRunXml(codigo) {
  return `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:rPr><w:noProof/><w:color w:val="888888"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>
<w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<wp:anchor distT="0" distB="0" distL="114300" distR="114300"
           simplePos="0" relativeHeight="251659264" behindDoc="0"
           locked="0" layoutInCell="1" allowOverlap="1">
  <wp:simplePos x="0" y="0"/>
  <wp:positionH relativeFrom="column"><wp:posOffset>-1667323</wp:posOffset></wp:positionH>
  <wp:positionV relativeFrom="paragraph"><wp:posOffset>-1041363</wp:posOffset></wp:positionV>
  <wp:extent cx="2142894" cy="249382"/>
  <wp:effectExtent l="0" t="0" r="0" b="0"/>
  <wp:wrapNone/>
  <wp:docPr id="99" name="code-vertical"/>
  <wp:cNvGraphicFramePr/>
  <a:graphic>
    <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
      <wps:wsp>
        <wps:cNvSpPr txBox="1"/>
        <wps:spPr>
          <a:xfrm rot="16200000"><a:off x="0" y="0"/><a:ext cx="2142894" cy="249382"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/><a:ln w="6350"><a:noFill/></a:ln>
        </wps:spPr>
        <wps:txbx>
          <w:txbxContent>
            <w:p>
              <w:r>
                <w:rPr>
                  <w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI" w:cs="Segoe UI"/>
                  <w:color w:val="888888"/><w:sz w:val="16"/><w:szCs w:val="16"/>
                </w:rPr>
                <w:t>${codigo}</w:t>
              </w:r>
            </w:p>
          </w:txbxContent>
        </wps:txbx>
        <wps:bodyPr rot="0" vert="horz" wrap="square" anchor="t" anchorCtr="0">
          <a:noAutofit/>
        </wps:bodyPr>
      </wps:wsp>
    </a:graphicData>
  </a:graphic>
</wp:anchor>
</w:drawing>
</w:r>`;
}

// Build HTML version for PDF generation
function buildHtmlDoc(data, logoB64) {
  const { proyecto, codigoProyecto, ordenCompra, ordenTrabajo, fecha,
          actividades, hallazgos, fotos, situacionGeneral,
          cumplimientoCronograma, avanceProyecto, accionesRequeridas,
          avance, supervisor, supervisorIniciales, codigo } = data;

  function romanItems(text) {
    if (!text) return '<p style="margin:2pt 0 6pt 40pt;font-size:11pt">No se ingresó información.</p>';
    const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
    return lines.map((l, i) => {
      const label = ['i','ii','iii','iv','v','vi','vii','viii','ix','x',
                     'xi','xii','xiii','xiv','xv','xvi','xvii','xviii','xix','xx'][i] || String(i+1);
      const clean = l.replace(/^[\s]*(([ivxlcdmIVXLCDM]+|[0-9]+)[\.\)]\s*)+/i,'')
                     .replace(/^[\s]*[-•·*]\s*/,'').trim();
      return `<p style="margin:3pt 0 5pt 0;font-size:11pt;text-align:justify;padding-left:60pt;text-indent:-20pt"><span style="font-weight:normal">${label}.&nbsp;&nbsp;&nbsp;${clean}</span></p>`;
    }).join('');
  }

  function bodyItems(text, fallback) {
    if (!text || !text.trim()) return `<p style="margin:2pt 0 6pt 40pt;font-size:11pt">${fallback}</p>`;
    const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
    return lines.map((l,i) => {
      const label = ['i','ii','iii','iv','v','vi','vii','viii','ix','x'][i] || String(i+1);
      const clean = l.replace(/^[\s]*(([ivxlcdmIVXLCDM]+|[0-9]+)[\.\)]\s*)+/i,'')
                     .replace(/^[\s]*[-•·*]\s*/,'').trim();
      return `<p style="margin:3pt 0 5pt 0;font-size:11pt;text-align:justify;padding-left:60pt;text-indent:-20pt">${label}.&nbsp;&nbsp;&nbsp;${clean}</p>`;
    }).join('');
  }

  // Photo grid
  let photoHtml = '';
  if (fotos && fotos.length) {
    photoHtml = '<table style="width:100%;border-collapse:collapse;margin-top:8pt">';
    for (let i = 0; i < fotos.length; i += 2) {
      const f1 = fotos[i], f2 = fotos[i+1]||null;
      photoHtml += '<tr>';
      [f1, f2].forEach(f => {
        if (!f) { photoHtml += '<td style="width:50%;padding:6pt"></td>'; return; }
        photoHtml += `<td style="width:50%;padding:6pt;text-align:center;vertical-align:top">
          <img src="${f.b64}" style="width:220pt;height:155pt;object-fit:cover;display:block;margin:0 auto"/>
          <p style="font-size:9pt;text-align:center;margin:4pt 0 0;color:#333">${f.legend||''}</p>
        </td>`;
      });
      photoHtml += '</tr>';
    }
    photoHtml += '</table>';
  } else {
    photoHtml = '<p style="margin:2pt 0 6pt 40pt;font-size:11pt">No se registraron fotografías en esta visita.</p>';
  }

  const otDisplay = (ordenTrabajo&&ordenTrabajo.trim()&&ordenTrabajo.trim()!==codigoProyecto)
    ? ordenTrabajo.trim() : (ordenCompra||'—');

  return `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  @page { size: A4; margin: 2cm 2.5cm 2.5cm 2.5cm; }
  body { font-family: "Segoe UI", Calibri, sans-serif; font-size: 11pt; color: #1a1a1a; margin:0; padding:0; }
  .hdr-blue { background: #1A3A6B; padding: 8pt 12pt 6pt; }
  .hdr-mpe { font-size: 16pt; font-weight: bold; color: white; margin:0; }
  .hdr-sub { font-size: 11pt; color: #1a1a1a; margin: 4pt 0 12pt; }
  table.meta { width:100%; border-collapse:collapse; margin-bottom:12pt; }
  table.meta td { padding:3pt 6pt; border:1px dotted #767171; font-size:10pt; }
  table.meta td:first-child { background:#EDF2FA; font-weight:bold; width:35%; }
  .subtitle { text-align:center; font-weight:bold; font-size:12pt; margin:12pt 0 8pt; }
  .mini p { margin:2pt 0; font-size:10pt; }
  .sec-title { font-weight:bold; font-size:11pt; margin:14pt 0 6pt; }
  .sub-title { font-weight:bold; font-size:10.5pt; margin:10pt 0 4pt; padding-left:20pt; }
  .footer { position:fixed; bottom:1cm; left:2.5cm; right:2.5cm; font-size:8pt; color:#888; }
  .footer-code { position:fixed; bottom:2cm; left:0.5cm; font-size:8pt; color:#888;
                 transform: rotate(-90deg); transform-origin: left bottom; white-space:nowrap; }
  .footer-page { position:fixed; bottom:1cm; right:2.5cm; font-size:8pt; color:#888; }
</style>
</head>
<body>
<div class="hdr-blue"><p class="hdr-mpe">MPE</p></div>
<p class="hdr-sub">Sistema de Reportes de Visita</p>

<table class="meta">
  <tr><td>Proyecto</td><td>${proyecto}</td></tr>
  <tr><td>Código Proyecto</td><td>${codigoProyecto}</td></tr>
  <tr><td>Orden de Compra</td><td>${ordenCompra}</td></tr>
  <tr><td>Fecha</td><td>${fecha}</td></tr>
  <tr><td>Avance estimado</td><td>${avance}%</td></tr>
  <tr><td>Calendario</td><td>${cumplimientoCronograma}</td></tr>
  <tr><td>Situación</td><td>${situacionGeneral}</td></tr>
  <tr><td>Supervisor</td><td>${supervisor} (${supervisorIniciales})</td></tr>
</table>

<p class="subtitle">Reporte Técnico de Visita de Supervisión</p>
<div class="mini">
  <p>Proyecto &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : &nbsp; ${proyecto}</p>
  <p>Orden Trabajo : &nbsp; ${otDisplay}</p>
  <p>Fecha Visita &nbsp;&nbsp; : &nbsp; ${fecha}</p>
</div>

<p class="sec-title">1.0 Actividades Realizadas</p>
${romanItems(actividades)}

<p class="sec-title">2.0 Hallazgos</p>
${bodyItems(hallazgos, 'No se registraron hallazgos significativos durante la presente visita de supervisión.')}

<p class="sec-title">3.0 Registro Fotográfico</p>
${photoHtml}

<p class="sec-title">4.0 Comentarios y Acciones</p>
<p class="sub-title">4.1 Situación General</p>
<p style="margin:2pt 0 6pt 40pt;font-size:11pt">${situacionGeneral||'No se registran observaciones críticas.'}</p>
<p class="sub-title">4.2 Cumplimiento de Cronograma</p>
<p style="margin:2pt 0 6pt 40pt;font-size:11pt">${cumplimientoCronograma}</p>
<p class="sub-title">4.3 Avance del Proyecto</p>
<p style="margin:2pt 0 6pt 40pt;font-size:11pt">Avance físico estimado ${avance}%${avanceProyecto&&avanceProyecto.trim()?' - '+avanceProyecto.trim():'.'}</p>
<p class="sub-title">4.4 Acciones Requeridas</p>
${bodyItems(accionesRequeridas, 'No se identifican acciones inmediatas.')}

<div class="footer-code">${codigo}</div>
</body>
</html>`;
}

async function generateDocx(data) {
  const { proyecto, codigoProyecto, ordenCompra, ordenTrabajo, fecha,
          actividades, hallazgos, fotos, situacionGeneral,
          cumplimientoCronograma, avanceProyecto, accionesRequeridas,
          avance, supervisor, supervisorIniciales, codigo } = data;

  const headerMPE = para(
    [tx('MPE', { bold:true, size:28, color:'FFFFFF', characterSpacing:8 })],
    { shading:{ fill:BLUE, type:ShadingType.CLEAR, color:'auto' }, spacing:{ before:0, after:0 } }
  );
  const headerSubtitle = para(
    [tx('Sistema de Reportes de Visita', { size:23, color:'141413' })],
    { spacing:{ before:0, after:160 } }
  );

  const idTable = new Table({
    width:{size:CONTENT_W,type:WidthType.DXA}, columnWidths:[3397, CONTENT_W-3397],
    rows:[
      idRow('Proyecto', proyecto), idRow('Código Proyecto', codigoProyecto),
      idRow('Orden de Compra', ordenCompra), idRow('Fecha', fecha),
      idRow('Avance estimado', avance+'%'), idRow('Calendario', cumplimientoCronograma),
      idRow('Situación', situacionGeneral), idRow('Supervisor', supervisor+' ('+supervisorIniciales+')'),
    ]
  });

  const footer = new Footer({ children:[
    para([
      tx('Pág ', { size:16, color:'888888' }),
      new TextRun({ children:[PageNumber.CURRENT], font:FONT, size:16, color:'888888' }),
      tx(' / ', { size:16, color:'888888' }),
      new TextRun({ children:[PageNumber.TOTAL_PAGES], font:FONT, size:16, color:'888888' }),
    ], { tabStops:[{ type:TabStopType.RIGHT, position:TabStopPosition.MAX }], spacing:{ before:60, after:0 } })
  ]});

  const actLines = (actividades||'').split('\n').map(l=>l.trim()).filter(l=>l.length>0);
  const actItems = actLines.length ? actLines.map((l,i)=>actividadItem(l,i)) : [bodyText('No se ingresó información.')];

  const hallLines = (hallazgos||'').split('\n').map(l=>l.trim()).filter(l=>l.length>0);
  const hallItems = hallLines.length ? hallLines.map((l,i)=>actividadItem(l,i)) : [bodyText('No se registraron hallazgos significativos durante la presente visita de supervisión.')];

  const fotosItems = (fotos&&fotos.length) ? [buildPhotoTable(fotos)] : [bodyText('No se registraron fotografías en esta visita.')];

  const accLines = (accionesRequeridas||'').split('\n').map(l=>l.trim()).filter(l=>l.length>0);
  const accItems = accLines.length ? accLines.map((l,i)=>actividadItem(l,i)) : [bodyText('No se identifican acciones inmediatas.')];

  const otDisplay = (ordenTrabajo&&ordenTrabajo.trim()&&ordenTrabajo.trim()!==codigoProyecto)
    ? ordenTrabajo.trim() : (ordenCompra||'—');

  const doc = new Document({
    styles:{ default:{ document:{ run:{ font:FONT, size:20 } } } },
    sections:[{
      properties:{ page:{ size:{ width:11906, height:16838 }, margin:{ top:1417, right:M_RIGHT, bottom:1417, left:M_LEFT, header:708, footer:708, gutter:0 } } },
      footers:{ default:footer },
      children:[
        headerMPE, headerSubtitle, idTable, emptyPara(160),
        para([tx('Reporte Técnico de Visita de Supervisión',{bold:true,size:24})],{alignment:AlignmentType.CENTER,spacing:{before:200,after:120}}),
        para([tx('Proyecto      :  '+proyecto,{size:20})],{spacing:{before:0,after:40}}),
        para([tx('Orden Trabajo :  '+otDisplay,{size:20})],{spacing:{before:0,after:40}}),
        para([tx('Fecha Visita  :  '+fecha,{size:20})],{spacing:{before:0,after:80}}),
        emptyPara(120),
        secTitle('1.0 Actividades Realizadas'), ...actItems, emptyPara(),
        secTitle('2.0 Hallazgos'), ...hallItems, emptyPara(),
        secTitle('3.0 Registro Fotográfico'), ...fotosItems, emptyPara(),
        secTitle('4.0 Comentarios y Acciones'),
        subSecTitle('4.1 Situación General'),
        bodyText(cleanLine(situacionGeneral||'No se registran observaciones críticas.')),
        subSecTitle('4.2 Cumplimiento de Cronograma'),
        bodyText(cleanLine(cumplimientoCronograma||'Las actividades se desarrollan según el programa.')),
        subSecTitle('4.3 Avance del Proyecto'),
        bodyText('Avance físico estimado '+avance+'%'+(avanceProyecto&&avanceProyecto.trim()?' - '+avanceProyecto.trim():'.')),
        subSecTitle('4.4 Acciones Requeridas'), ...accItems,
        emptyPara(200),
      ]
    }]
  });

  let buffer = await Packer.toBuffer(doc);
  const zip = await JSZip.loadAsync(buffer);
  const footerFiles = Object.keys(zip.files).filter(f=>f.match(/word\/footer\d+\.xml/));
  if (footerFiles.length > 0) {
    let fXml = await zip.files[footerFiles[0]].async('string');
    fXml = fXml.replace(/(<\/w:pPr>)/, '$1' + makeVerticalRunXml(codigo));
    zip.file(footerFiles[0], fXml);
    buffer = await zip.generateAsync({ type:'nodebuffer', compression:'DEFLATE' });
  }
  return buffer;
}

async function generatePdf(data) {
  return await generatePdfBuffer(data);
}

// ── ROUTES ────────────────────────────────────────────────
app.get('/', (req,res) => res.json({status:'MPE Reporte Server OK',version:'3.2'}));

app.post('/generar', async (req,res) => {
  try {
    const data = req.body;
    if (!data.proyecto) return res.status(400).json({error:'Faltan datos.'});
    const buffer = await generateDocx(data);
    const filename = 'Reporte_'+(data.codigo||'MPE').replace(/[^a-zA-Z0-9]/g,'_')+'.docx';
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition','attachment; filename="'+filename+'"');
    res.setHeader('Access-Control-Allow-Origin','*');
    res.send(buffer);
  } catch(e) { console.error(e); res.status(500).json({error:e.message}); }
});

app.post('/generar-pdf', async (req,res) => {
  try {
    const data = req.body;
    if (!data.proyecto) return res.status(400).json({error:'Faltan datos.'});
    const buffer = await generatePdf(data);
    const filename = 'Reporte_'+(data.codigo||'MPE').replace(/[^a-zA-Z0-9]/g,'_')+'.pdf';
    res.setHeader('Content-Type','application/pdf');
    res.setHeader('Content-Disposition','attachment; filename="'+filename+'"');
    res.setHeader('Access-Control-Allow-Origin','*');
    res.send(buffer);
  } catch(e) { console.error(e); res.status(500).json({error:e.message}); }
});

app.options('*', (req,res) => {
  res.setHeader('Access-Control-Allow-Origin','*');
  res.setHeader('Access-Control-Allow-Methods','POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers','Content-Type');
  res.sendStatus(200);
});

const PORT = process.env.PORT||3000;
app.listen(PORT, ()=>console.log('MPE Server v3.1 running on port '+PORT));
