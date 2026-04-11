const express = require('express');
const cors = require('cors');
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
const BLUE = '1F3864';
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
  return para([tx('')], { spacing: { before: 0, after: space } });
}

const thinBorder  = { style: BorderStyle.SINGLE, size: 4, color: 'BBBBBB' };
const cellBorders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
const noBorder    = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders   = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// Strip leading bullets/numbers/roman from any line
function cleanLine(line) {
  return (line||'')
    .replace(/^[\s]*([ivxlcdmIVXLCDM]+[\.\)]\s*)+/i, '')
    .replace(/^[\s]*\d+[\.\)]\s*/, '')
    .replace(/^[\s]*[-•·*▪▸→]\s*/, '')
    .replace(/^[\s]*[a-zA-Z][\.\)]\s*/, '')
    .trim();
}

const ROMAN = ['i','ii','iii','iv','v','vi','vii','viii','ix','x',
               'xi','xii','xiii','xiv','xv','xvi','xvii','xviii','xix','xx'];

function secTitle(text) {
  return para([tx(text, { bold: true, size: 22 })],
    { spacing: { before: 240, after: 100 } });
}
function subSecTitle(text) {
  return para([tx(text, { bold: true, size: 20 })],
    { spacing: { before: 180, after: 60 } });
}
function bodyText(text) {
  return para([tx(text, { size: 20 })], {
    alignment: AlignmentType.BOTH,
    spacing: { before: 0, after: 80 },
    indent: { left: 720 }
  });
}
function actividadItem(rawText, num) {
  const label = ROMAN[num] || String(num+1);
  const text  = cleanLine(rawText);
  return para([tx(label+'.    '+text, { size: 20 })], {
    alignment: AlignmentType.BOTH,
    spacing: { before: 40, after: 60 },
    indent: { left: 1080, hanging: 360 }
  });
}
function idRow(label, value) {
  return new TableRow({ children: [
    new TableCell({
      borders: cellBorders, width: { size: 2800, type: WidthType.DXA },
      margins: { top: 60, bottom: 60, left: 120, right: 120 },
      children: [para([tx(label, { bold: true, size: 18 })], { spacing: { before:0, after:0 } })]
    }),
    new TableCell({
      borders: cellBorders, width: { size: CONTENT_W-2800, type: WidthType.DXA },
      margins: { top: 60, bottom: 60, left: 120, right: 120 },
      children: [para([tx(String(value||'—'), { size: 18 })], { spacing: { before:0, after:0 } })]
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
      try { imgData=Buffer.from(f.b64.split(',')[1],'base64'); imgType=f.b64.startsWith('data:image/png')?'png':'jpg'; }
      catch(e) { return new TableCell({ borders:noBorders, width:{size:colW,type:WidthType.DXA}, children:[para([tx(f.legend||'',{size:16})],{alignment:AlignmentType.CENTER})] }); }
      const legend=(f.legend||'').trim();
      return new TableCell({
        borders:noBorders, width:{size:colW,type:WidthType.DXA},
        margins:{top:80,bottom:80,left:60,right:60},
        children:[
          para([new ImageRun({type:imgType,data:imgData,transformation:{width:215,height:155},altText:{title:legend||'foto',description:legend,name:legend||'foto'}})],
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

async function generateDocx(data) {
  const { proyecto, codigoProyecto, ordenCompra, ordenTrabajo, fecha,
          actividades, hallazgos, fotos, situacionGeneral,
          cumplimientoCronograma, avanceProyecto, accionesRequeridas,
          avance, supervisor, supervisorIniciales, codigo } = data;

  // Header: blue bar, MPE bold left + subtitle right
  const headerTable = new Table({
    width:{size:CONTENT_W,type:WidthType.DXA}, columnWidths:[1500, CONTENT_W-1500],
    borders:{top:noBorder,bottom:noBorder,left:noBorder,right:noBorder,insideH:noBorder,insideV:noBorder},
    rows:[new TableRow({ children:[
      new TableCell({
        borders:noBorders, width:{size:1500,type:WidthType.DXA},
        shading:{fill:BLUE,type:'clear'},
        margins:{top:120,bottom:100,left:180,right:0},
        children:[para([tx('MPE',{bold:true,size:46,color:'FFFFFF'})],{spacing:{before:0,after:0}})]
      }),
      new TableCell({
        borders:noBorders, width:{size:CONTENT_W-1500,type:WidthType.DXA},
        shading:{fill:BLUE,type:'clear'},
        margins:{top:150,bottom:100,left:160,right:160},
        children:[para([tx('Sistema de Reportes de Visita',{size:22,color:'FFFFFF'})],
          {spacing:{before:0,after:0},alignment:AlignmentType.LEFT})]
      })
    ]})]
  });

  const idTable = new Table({
    width:{size:CONTENT_W,type:WidthType.DXA}, columnWidths:[2800,CONTENT_W-2800],
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

  const footer = new Footer({ children:[
    para([
      tx(codigo, {size:16,color:'888888'}),
      new TextRun({text:'\t',font:FONT}),
      tx('Pág ',{size:16,color:'888888'}),
      new TextRun({children:[PageNumber.CURRENT],font:FONT,size:16,color:'888888'}),
      tx(' / ',{size:16,color:'888888'}),
      new TextRun({children:[PageNumber.TOTAL_PAGES],font:FONT,size:16,color:'888888'}),
    ], {
      border:{top:{style:BorderStyle.SINGLE,size:4,color:'BBBBBB',space:4}},
      tabStops:[{type:TabStopType.RIGHT,position:TabStopPosition.MAX}],
      spacing:{before:60,after:0}
    })
  ]});

  const actLines = (actividades||'').split('\n').map(l=>l.trim()).filter(l=>l.length>0);
  const actItems = actLines.length ? actLines.map((l,i)=>actividadItem(l,i)) : [bodyText('No se ingresó información.')];

  const hallLines = (hallazgos||'').split('\n').map(l=>cleanLine(l)).filter(l=>l.length>0);
  const hallItems = hallLines.length ? hallLines.map(l=>bodyText(l)) : [bodyText('No se registraron hallazgos significativos durante la presente visita de supervisión.')];

  const fotosItems = (fotos&&fotos.length) ? [buildPhotoTable(fotos)] : [bodyText('No se registraron fotografías en esta visita.')];

  const accLines = (accionesRequeridas||'').split('\n').map(l=>cleanLine(l)).filter(l=>l.length>0);
  const accItems = accLines.length ? accLines.map(l=>bodyText(l)) : [bodyText('No se identifican acciones inmediatas.')];

  // OT display: prefer ordenTrabajo field, fallback to ordenCompra
  const otDisplay = (ordenTrabajo&&ordenTrabajo.trim()&&ordenTrabajo.trim()!==codigoProyecto)
    ? ordenTrabajo.trim() : (ordenCompra||'—');

  const doc = new Document({
    styles:{ default:{ document:{ run:{ font:FONT, size:20 } } } },
    sections:[{
      properties:{ page:{
        size:{width:11906,height:16838},
        margin:{top:1417,right:M_RIGHT,bottom:1417,left:M_LEFT,header:708,footer:708,gutter:0}
      }},
      footers:{default:footer},
      children:[
        headerTable,
        emptyPara(120),
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
        bodyText('Avance físico estimado '+avance+'%'+(avanceProyecto&&avanceProyecto.trim()?' - '+avanceProyecto.trim():'.')),
        subSecTitle('4.4 Acciones Requeridas'),
        ...accItems,
        emptyPara(200),
      ]
    }]
  });
  return await Packer.toBuffer(doc);
}

app.get('/', (req,res) => res.json({status:'MPE Reporte Server OK',version:'2.0'}));

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
app.listen(PORT, ()=>console.log('MPE Server v2.0 running on port '+PORT));
