const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,AlignmentType,BorderStyle,HeadingLevel,LevelFormat,Footer,PageNumber}=require('docx');
const W=[500,1150,4300,3688]; const TW=9638;
const b={style:BorderStyle.SINGLE,size:4,color:"808080"}; const borders={top:b,bottom:b,left:b,right:b};
const F="Times New Roman";
const p=(t,o={})=>new Paragraph({alignment:o.al,spacing:{after:o.after??120,before:o.before??0},indent:o.ind,children:[new TextRun({text:t,bold:o.bold,size:o.size??24,font:F,italics:o.it})]});
const cell=(t,i,h)=>new TableCell({width:{size:W[i],type:WidthType.DXA},borders,margins:{top:60,bottom:60,left:90,right:90},
  shading:h?{fill:"D9E2F3",type:ShadingType.CLEAR,color:"auto"}:undefined,
  children:[new Paragraph({alignment:(i<2)?AlignmentType.CENTER:AlignmentType.LEFT,children:[new TextRun({text:t,bold:h,size:21,font:F})]})]});
const H=["№","Дата","Выполненная работа","Результат / документ"];
let n=0;
const table=rows=>new Table({width:{size:TW,type:WidthType.DXA},columnWidths:W,rows:[
  new TableRow({tableHeader:true,children:H.map((t,i)=>cell(t,i,true))}),
  ...rows.map(r=>new TableRow({cantSplit:true,children:[cell(String(++n),0),cell(r[0],1),cell(r[1],2),cell(r[2],3)]}))]});
const h=t=>new Paragraph({heading:HeadingLevel.HEADING_2,spacing:{before:240,after:120},children:[new TextRun({text:t,bold:true,size:26,font:F})]});
const bullet=t=>new Paragraph({numbering:{reference:"bl",level:0},spacing:{after:60},children:[new TextRun({text:t,size:24,font:F})]});

const S=JSON.parse(fs.readFileSync(process.argv[2],'utf8'));
const children=[
 p("ОТЧЁТ",{al:AlignmentType.CENTER,bold:true,size:32,after:60}),
 p("о выполненных работах отдела охраны труда и промышленной безопасности",{al:AlignmentType.CENTER,after:60}),
 p("ООО «Авиационные интерьеры-СН»",{al:AlignmentType.CENTER,after:60}),
 p(S.period,{al:AlignmentType.CENTER,after:240}),
 p("Исполнитель: начальник отдела охраны труда и промышленной безопасности Айкешева А.К.",{after:60}),
 p("Направления: охрана труда, промышленная безопасность, экология.",{after:240}),
 h("1. Сводные показатели"),
 ...S.summary.map(bullet),
];
S.sections.forEach((s,i)=>{children.push(h(`${i+2}. ${s.title}`)); children.push(table(s.rows));});
children.push(h(`${S.sections.length+2}. Ключевые результаты в цифрах`));
S.keys.forEach(t=>children.push(bullet(t)));
children.push(h(`${S.sections.length+3}. Примечание`));
S.notes.forEach(t=>children.push(p(t,{it:true,size:22})));
children.push(p("",{after:360}));
children.push(p("Начальник отдела охраны труда",{after:0}));
children.push(p("и промышленной безопасности                    _______________ / А.К. Айкешева /",{after:120}));
children.push(p("«____» _______________ 2026 г."));
const doc=new Document({numbering:{config:[{reference:"bl",levels:[{level:0,format:LevelFormat.BULLET,text:"–",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:500,hanging:280}}}}]}]},
 styles:{default:{document:{run:{font:F,size:24}}}},
 sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:1134,bottom:1134,left:1134,right:1134}}},
  footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.RIGHT,children:[new TextRun({children:[PageNumber.CURRENT],size:20,font:F})]})]})},
  children}]});
Packer.toBuffer(doc).then(buf=>fs.writeFileSync(process.argv[3],buf));
