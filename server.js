const express = require('express');
const cors = require('cors');
const JSZip = require('jszip');
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
  return para([tx(text, { bold:true, size:22 })],
    { spacing:{ before:240, after:100 } });
}
function subSecTitle(text) {
  return para([tx(text, { bold:true, size:20 })],
    { spacing:{ before:180, after:60 } });
}
function bodyText(text) {
  return para([tx(text, { size:20 })], {
    alignment: AlignmentType.BOTH,
    spacing: { before:0, after:80 },
    indent: { left:720 }
  });
}
function actividadItem(rawText, num) {
  const label = ROMAN[num] || String(num+1);
  const text  = cleanLine(rawText);
  return para([
    tx(label+'.', { size:20 }),
    tx('\t'+text, { size:20 })
  ], {
    alignment: AlignmentType.BOTH,
    spacing: { before:40, after:60 },
    indent: { left:900, hanging:360 },
    tabStops: [{ type:TabStopType.LEFT, position:540 }]
  });
}
function idRow(label, value) {
  return new TableRow({ children:[
    new TableCell({
      borders: dottedBorders,
      width: { size:3397, type:WidthType.DXA },
      shading: { fill:'EDF2FA', type:ShadingType.CLEAR },
      margins: { top:15, bottom:15, left:120, right:120 },
      children: [para([tx(label, { bold:true, size:18 })],
        { spacing:{ before:0, after:0 } })]
    }),
    new TableCell({
      borders: dottedBorders,
      width: { size:CONTENT_W-3397, type:WidthType.DXA },
      margins: { top:15, bottom:15, left:120, right:120 },
      children: [para([tx(String(value||'—'), { size:18 })],
        { spacing:{ before:0, after:0 } })]
    })
  ]});
}
function buildPhotoTable(fotosArr) {
  const colW = Math.floor(CONTENT_W/2)-60;
  const rows = [];
  for (let i=0; i<fotosArr.length; i+=2) {
    const f1=fotosArr[i], f2=fotosArr[i+1]||null;
    function photoCell(f) {
      if (!f) return new TableCell({ borders:noBorders, width:{size:colW,type:WidthType.DXA}, children:[emptyPara()] });
      let imgData, imgType;
      try {
        imgData=Buffer.from(f.b64.split(',')[1],'base64');
        imgType=f.b64.startsWith('data:image/png')?'png':'jpg';
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
            transformation:{width:215,height:155},
            altText:{title:legend||'foto',description:legend,name:legend||'foto'}})],
            {alignment:AlignmentType.CENTER,spacing:{before:0,after:40}}),
          para([tx(legend,{size:16})],
            {alignment:AlignmentType.CENTER,spacing:{before:0,after:80}})
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
          <a:xfrm rot="16200000">
            <a:off x="0" y="0"/>
            <a:ext cx="2142894" cy="249382"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/>
          <a:ln w="6350"><a:noFill/></a:ln>
        </wps:spPr>
        <wps:txbx>
          <w:txbxContent>
            <w:p>
              <w:r>
                <w:rPr>
                  <w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI" w:cs="Segoe UI"/>
                  <w:color w:val="888888"/>
                  <w:sz w:val="16"/><w:szCs w:val="16"/>
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

async function generateDocx(data) {
  const { proyecto, codigoProyecto, ordenCompra, ordenTrabajo, fecha,
          actividades, hallazgos, fotos, situacionGeneral,
          cumplimientoCronograma, avanceProyecto, accionesRequeridas,
          avance, supervisor, supervisorIniciales, codigo } = data;

  // Header: Para 1 = blue+MPE, Para 2 = subtitle black
  const headerMPE = para(
    [tx('MPE', { bold:true, size:28, color:'FFFFFF', characterSpacing:8 })],
    { shading:{ fill:BLUE, type:ShadingType.CLEAR, color:'auto' },
      spacing:{ before:0, after:0 } }
  );
  const headerSubtitle = para(
    [tx('Sistema de Reportes de Visita', { size:23, color:'141413' })],
    { spacing:{ before:0, after:160 } }
  );

  const idTable = new Table({
    width:{size:CONTENT_W,type:WidthType.DXA},
    columnWidths:[3397, CONTENT_W-3397],
    rows:[
      idRow('Proyecto', proyecto),
      idRow('Código Proyecto', codigoProyecto),
      idRow('Orden de Compra', ordenCompra),
      idRow('Fecha', fecha),
      idRow('Avance estimado', avance+'%'),
      idRow('Calendario', cumplimientoCronograma),
      idRow('Situación', situacionGeneral),
      idRow('Supervisor', supervisor+' ('+supervisorIniciales+')'),
    ]
  });

  // Footer: Pág X/Y right only — vertical code injected via JSZip later
  const footer = new Footer({
    children:[
      para([
        tx('Pág ', { size:16, color:'888888' }),
        new TextRun({ children:[PageNumber.CURRENT], font:FONT, size:16, color:'888888' }),
        tx(' / ', { size:16, color:'888888' }),
        new TextRun({ children:[PageNumber.TOTAL_PAGES], font:FONT, size:16, color:'888888' }),
      ], {
        // NO border line
        tabStops:[{ type:TabStopType.RIGHT, position:TabStopPosition.MAX }],
        spacing:{ before:60, after:0 }
      })
    ]
  });

  const actLines = (actividades||'').split('\n').map(l=>l.trim()).filter(l=>l.length>0);
  const actItems = actLines.length
    ? actLines.map((l,i)=>actividadItem(l,i))
    : [bodyText('No se ingresó información.')];

  const hallLines = (hallazgos||'').split('\n').map(l=>l.trim()).filter(l=>l.length>0);
  const hallItems = hallLines.length
    ? hallLines.map((l,i)=>actividadItem(l,i))
    : [bodyText('No se registraron hallazgos significativos durante la presente visita de supervisión.')];

  const fotosItems = (fotos&&fotos.length)
    ? [buildPhotoTable(fotos)]
    : [bodyText('No se registraron fotografías en esta visita.')];

  const accLines = (accionesRequeridas||'').split('\n').map(l=>l.trim()).filter(l=>l.length>0);
  const accItems = accLines.length
    ? accLines.map((l,i)=>actividadItem(l,i))
    : [bodyText('No se identifican acciones inmediatas.')];

  const otDisplay = (ordenTrabajo&&ordenTrabajo.trim()&&ordenTrabajo.trim()!==codigoProyecto)
    ? ordenTrabajo.trim() : (ordenCompra||'—');

  const doc = new Document({
    styles:{ default:{ document:{ run:{ font:FONT, size:20 } } } },
    sections:[{
      properties:{
        page:{
          size:{ width:11906, height:16838 },
          margin:{ top:1417, right:M_RIGHT, bottom:1417, left:M_LEFT,
                   header:708, footer:708, gutter:0 }
        }
      },
      footers:{ default:footer },
      children:[
        headerMPE,
        headerSubtitle,
        idTable,
        emptyPara(160),
        para([tx('Reporte Técnico de Visita de Supervisión',{bold:true,size:24})],
          {alignment:AlignmentType.CENTER,spacing:{before:200,after:120}}),
        para([tx('Proyecto      :  '+proyecto,{size:20})],{spacing:{before:0,after:40}}),
        para([tx('Orden Trabajo :  '+otDisplay,{size:20})],{spacing:{before:0,after:40}}),
        para([tx('Fecha Visita  :  '+fecha,{size:20})],{spacing:{before:0,after:80}}),
        emptyPara(120),
        secTitle('1.0 Actividades Realizadas'),
        ...actItems,
        emptyPara(),
        secTitle('2.0 Hallazgos'),
        ...hallItems,
        emptyPara(),
        secTitle('3.0 Registro Fotográfico'),
        ...fotosItems,
        emptyPara(),
        secTitle('4.0 Comentarios y Acciones'),
        subSecTitle('4.1 Situación General'),
        bodyText(cleanLine(situacionGeneral||'No se registran observaciones críticas.')),
        subSecTitle('4.2 Cumplimiento de Cronograma'),
        bodyText(cleanLine(cumplimientoCronograma||'Las actividades se desarrollan según el programa.')),
        subSecTitle('4.3 Avance del Proyecto'),
        bodyText('Avance físico estimado '+avance+'%'+
          (avanceProyecto&&avanceProyecto.trim()?' - '+avanceProyecto.trim():'.')),
        subSecTitle('4.4 Acciones Requeridas'),
        ...accItems,
        emptyPara(200),
      ]
    }]
  });

  // Generate base buffer
  let buffer = await Packer.toBuffer(doc);

  // Inject vertical rotated textbox into footer via JSZip
  const zip = await JSZip.loadAsync(buffer);
  const footerFiles = Object.keys(zip.files).filter(f=>f.match(/word\/footer\d+\.xml/));

  if (footerFiles.length > 0) {
    let fXml = await zip.files[footerFiles[0]].async('string');
    const vertRun = makeVerticalRunXml(codigo);
    // Insert vertical run right after the first <w:pPr>...</w:pPr> in footer
    fXml = fXml.replace(/(<\/w:pPr>)/, '$1' + vertRun);
    zip.file(footerFiles[0], fXml);
    buffer = await zip.generateAsync({ type:'nodebuffer', compression:'DEFLATE' });
  }

  return buffer;
}

app.get('/', (req,res) => res.json({status:'MPE Reporte Server OK',version:'3.0'}));

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

app.options('/generar', (req,res) => {
  res.setHeader('Access-Control-Allow-Origin','*');
  res.setHeader('Access-Control-Allow-Methods','POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers','Content-Type');
  res.sendStatus(200);
});

const PORT = process.env.PORT||3000;
app.listen(PORT, ()=>console.log('MPE Server v3.0 running on port '+PORT));
