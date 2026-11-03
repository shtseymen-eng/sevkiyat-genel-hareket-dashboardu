import {TYPES,availableTypes,groupRows,summarize} from './analysis.mjs';
const COLORS={vehicle:'087E83',shipIn:'408ABE',shipOut:'405774',pipeline:'B58127'};
const numberFormat='#,##0.000;[Red](#,##0.000);"–"';
// Explicit OOXML styles keep browser exports formatted without a server dependency.
function stylesXml(){
 const fonts=[['182F43',11,false],['FFFFFF',21,true],['FFFFFF',11,true],['102941',11,true],['FFFFFF',11,true],['6D8090',10,false]];
 const fills=['FFFFFF','102941','087E83','E7F3F2','F3F7FA','DCE7EF','408ABE','405774','B58127'];
 const xfs=[[0,2,0,0],[1,3,0,0],[5,2,0,0],[2,4,0,0],[0,2,1,0],[0,6,1,0],[0,2,1,164],[0,6,1,164],[0,2,1,165],[0,6,1,165],[3,5,1,0],[3,5,1,164],[3,5,1,165],[4,3,1,0],[4,3,1,164],[4,3,1,165],[2,8,0,0],[2,9,0,0],[2,10,0,0]];
 return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="2"><numFmt numFmtId="164" formatCode="${numberFormat.replace(/"/g,'&quot;')}"/><numFmt numFmtId="165" formatCode="#,##0;[Red](#,##0);&quot;–&quot;"/></numFmts><fonts count="${fonts.length}">${fonts.map(([color,size,bold])=>`<font>${bold?'<b/>':''}<sz val="${size}"/><color rgb="FF${color}"/><name val="Calibri"/><family val="2"/></font>`).join('')}</fonts><fills count="${fills.length+2}"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>${fills.map(c=>`<fill><patternFill patternType="solid"><fgColor rgb="FF${c}"/><bgColor indexed="64"/></patternFill></fill>`).join('')}</fills><borders count="2"><border/><border><bottom style="hair"><color rgb="FFE1E8ED"/></bottom></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.map(([font,fill,border,num])=>`<xf numFmtId="${num}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment vertical="center" wrapText="1"${num?' horizontal="right"':''}/></xf>`).join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
}
export async function buildExcelReport(rows,options,XLSX,JSZip){
 const book=XLSX.utils.book_new(),specs=[];
 const period=options.periodLabel||'Tüm dönemler';
 const filterLabel=[options.owner,options.product,options.tank?`Tank ${options.tank}`:'',options.query?`Arama: ${options.query}`:''].filter(Boolean).join(' / ');
 function sheet(name,title,headers,widths,accent='087E83'){
  const spec={name,accent,headers,widths,headerRow:5,data:[[title], [`Dönem: ${period}${filterLabel?' / '+filterLabel:''}`],['Miktar: ton (1.000 kg) / Araç seferi: aynı plaka + tank + kantar çıkış zamanı'],[],headers],kinds:['title','note','note','blank','header']};specs.push(spec);return spec;
 }
 function push(s,row,kind='body'){s.data.push(row);s.kinds.push(kind);}
 const sum=summarize(rows);
 const overview=sheet('Genel Dashboard','GENEL HAREKET RAPORU',['Hareket türü','Giriş (ton)','Çıkış (ton)','Araç seferi','Hareket satırı'],[54,22,22,18,18]);
 for(const type of availableTypes(rows)){const part=rows.filter(r=>r.type===type),s=summarize(part);push(overview,[TYPES[type],s.incoming/1000,s.outgoing/1000,type==='vehicle'?s.trips:null,part.length]);}
 push(overview,['GENEL TOPLAM',sum.incoming/1000,sum.outgoing/1000,sum.trips,rows.length],'total');
 push(overview,[],'blank');push(overview,['SÖZLEŞME SAHİBİ BAZINDA ÖZET'],'section');
 push(overview,['Sözleşme sahibi','Giriş (ton)','Çıkış (ton)','Araç seferi','Hareket satırı'],'header');
 for(const g of groupRows(rows,['owner'])){const s=summarize(g.rows);push(overview,[g.owner,s.incoming/1000,s.outgoing/1000,s.trips,g.rows.length]);}
 push(overview,['GENEL TOPLAM',sum.incoming/1000,sum.outgoing/1000,sum.trips,rows.length],'total');
 if(sum.missingScale){push(overview,[],'blank');push(overview,[`${sum.missingScale} araç satırında plaka veya kantar zamanı eksik; bu kayıtlar ayrı sefer sayıldı.`],'note');}
 for(const type of availableTypes(rows)){
  const part=rows.filter(r=>r.type===type),hasDestination=type==='shipOut',isVehicle=type==='vehicle';
  const headers=['Sözleşme sahibi','Beyan sahibi',...(hasDestination?['Gideceği firma']:[]),'Ürün','Tank','Miktar (ton)',...(isVehicle?['Araç seferi']:[]),'Hareket satırı'];
  const widths=[43,43,...(hasDestination?[43]:[]),36,13,21,...(isVehicle?[16]:[]),18];
  const s=sheet(TYPES[type],TYPES[type].toLocaleUpperCase('tr-TR')+' RAPORU',headers,widths,COLORS[type]);
  for(const owner of groupRows(part,['owner']).sort((a,b)=>a.owner.localeCompare(b.owner,'tr'))){
   for(const g of groupRows(owner.rows,['declarant',...(hasDestination?['destination']:[]),'product','tank']).sort((a,b)=>[a.declarant,a.product,a.tank].join().localeCompare([b.declarant,b.product,b.tank].join(),'tr'))){
    push(s,[owner.owner,g.declarant,...(hasDestination?[g.destination]:[]),g.product,g.tank,g.quantity/1000,...(isVehicle?[summarize(g.rows).trips]:[]),g.rows.length]);
   }
   push(s,[owner.owner+' / ARA TOPLAM',...Array(headers.length-1).fill('')],'subtotal');const r=s.data.at(-1),q=headers.indexOf('Miktar (ton)');r[q]=owner.quantity/1000;if(isVehicle)r[q+1]=summarize(owner.rows).trips;r[headers.length-1]=owner.rows.length;
  }
  const grand=['GENEL TOPLAM',...Array(headers.length-1).fill('')],q=headers.indexOf('Miktar (ton)');grand[q]=part.reduce((a,r)=>a+r.quantity,0)/1000;if(isVehicle)grand[q+1]=summarize(part).trips;grand[headers.length-1]=part.length;push(s,grand,'total');
 }
 const raw=sheet('Hareket Kayıtları','HAREKET KAYITLARI',['Hareket türü','İşlem tarihi','Sözleşme sahibi','Beyan sahibi','Gideceği firma','Ürün','Tank','Miktar (ton)','Plaka / gemi','Kantar çıkış','Belge no','Beyanname','Kaynak dosya','Kaynak satır'],[20,26,44,44,44,36,13,21,30,26,26,28,38,16]);
 for(const r of [...rows].sort((a,b)=>a.owner.localeCompare(b.owner,'tr')||a.date.localeCompare(b.date)))push(raw,[TYPES[r.type],r.date.replace('T',' '),r.owner,r.declarant,r.destination,r.product,r.tank,r.quantity/1000,r.plate||r.ship,r.scale.replace('T',' '),r.document,r.declaration,r.source,r.row]);
 for(const s of specs){
  const cols=s.headers.length,last=XLSX.utils.encode_col(cols-1);const ws=XLSX.utils.aoa_to_sheet(s.data);
  ws['!cols']=s.widths.map(wch=>({wch}));ws['!merges']=[];
  ws['!rows']=s.kinds.map((kind,i)=>{const base=kind==='title'?40:kind==='note'?26:kind==='blank'?12:kind==='header'?30:kind==='section'?29:kind==='body'?46:34;const height=['body','subtotal','total'].includes(kind)?Math.max(base,...s.data[i].map((v,c)=>typeof v==='string'?Math.ceil(v.length/(s.widths[c]*.85))*15+12:base)):base;return {hpt:height};});
  s.kinds.forEach((kind,i)=>{if(['title','note','section'].includes(kind))ws['!merges'].push({s:{r:i,c:0},e:{r:i,c:cols-1}});
   for(let c=0;c<cols;c++){const addr=XLSX.utils.encode_cell({r:i,c});if(!ws[addr])ws[addr]={t:'s',v:''};}
  });
  if(s===raw)ws['!autofilter']={ref:`A5:${last}${s.data.length}`};
  XLSX.utils.book_append_sheet(book,ws,s.name);
 }
 const zip=await JSZip.loadAsync(XLSX.write(book,{type:'array',bookType:'xlsx'}));zip.file('xl/styles.xml',stylesXml());
 for(let i=0;i<specs.length;i++){
  const s=specs[i],path=`xl/worksheets/sheet${i+1}.xml`;let xml=await zip.file(path).async('string');
  xml=xml.replace(/<c\b([^>]*?)\br="([A-Z]+)(\d+)"([^>]*?)>/g,(match,before,col,r,after)=>{
   const row=Number(r)-1,c=XLSX.utils.decode_col(col),kind=s.kinds[row],header=s.headers[c]||'',isNumber=typeof s.data[row]?.[c]==='number',isQty=/ton/.test(header),stripe=row%2===0;
   let style=kind==='title'?1:kind==='note'?2:kind==='section'?10:kind==='header'?({"408ABE":16,"405774":17,"B58127":18}[s.accent]||3):kind==='total'?(isNumber?(isQty?14:15):13):kind==='subtotal'?(isNumber?(isQty?11:12):10):isNumber?(isQty?(stripe?7:6):(stripe?9:8)):(stripe?5:4);
   return `<c${before.replace(/\s+s="\d+"/g,'')}r="${col}${r}"${after.replace(/\s+s="\d+"/g,'')} s="${style}">`;
  });
  const view=`<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane xSplit="1" ySplit="5" topLeftCell="B6" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="B6" sqref="B6"/></sheetView></sheetViews>`;
  xml=/<sheetViews[\s>]/.test(xml)?xml.replace(/<sheetViews[\s\S]*?<\/sheetViews>/,view):xml.replace(/(<dimension[^>]*\/>)/,'$1'+view);
  xml=xml.replace(/<pageMargins[^>]*\/>/,'');
  const printSettings='<printOptions horizontalCentered="1"/><pageMargins left="0.3" right="0.3" top="0.5" bottom="0.5" header="0.2" footer="0.2"/><pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>';
  // Worksheet elements have a strict order; print settings precede ignoredErrors.
  xml=xml.replace(/(?=<ignoredErrors\b|<\/worksheet>)/,printSettings);
  zip.file(path,xml);
 }
 return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}
