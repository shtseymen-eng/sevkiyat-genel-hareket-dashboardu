import {buildExcelReport} from "./export.mjs";
import {TYPES,normalizeRows,filterRows,summarize,groupRows,detectType,availableTypes} from './analysis.mjs';
const $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
const colors={vehicle:'#0b8488',shipIn:'#408abe',shipOut:'#405774',pipeline:'#d3a24e'};
const icons={vehicle:'↗',shipIn:'↙',shipOut:'↗',pipeline:'⇢'};
const n=v=>new Intl.NumberFormat('tr-TR',{maximumFractionDigits:3}).format(v);
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let sources=[],activeTab='all',expanded=new Set();
const allRows=()=>sources.flatMap(s=>s.rows);
const periods=()=>[...new Set(allRows().map(r=>r.period))].sort();
const monthName=p=>new Intl.DateTimeFormat('tr-TR',{month:'long',year:'numeric'}).format(new Date(Number(p.slice(0,4)),Number(p.slice(5))-1,1));
function selection(){return {periods:$('#period').value?[$('#period').value]:[],types:$$('#typeChecks input:checked').map(e=>e.value),owner:$('#owner').value,product:$('#product').value,tank:$('#tank').value,query:$('#search').value};}
function visible(){return filterRows(allRows(),selection());}
for(const id of ['uploadOpen','sourcesOpen'])$('#'+id).onclick=()=>$('#uploadDialog').showModal();
$$('[data-close]').forEach(b=>b.onclick=()=>$('#'+b.dataset.close).close());
function updateOptions(){
 const oldTypes=new Map($$('#typeChecks input').map(e=>[e.value,e.checked]));
 $('#typeChecks').innerHTML=availableTypes(allRows()).map(key=>`<label><input type="checkbox" value="${key}" ${oldTypes.get(key)!==false?'checked':''}>${TYPES[key]}</label>`).join('')||'<span>Henüz rapor yüklenmedi</span>';
 $$('#typeChecks input').forEach(e=>e.onchange=()=>{expanded.clear();render();});
 for(const [id,values,label] of [['period',periods(),'Tüm dönemler'],['owner',[...new Set(allRows().map(r=>r.owner))].sort(),'Tüm firmalar'],['product',[...new Set(allRows().map(r=>r.product))].sort(),'Tüm ürünler'],['tank',[...new Set(allRows().map(r=>r.tank))].sort((a,b)=>a.localeCompare(b,'tr',{numeric:true})),'Tüm tanklar']]){
  const old=$('#'+id).value;$('#'+id).innerHTML=`<option value="">${label}</option>`+values.map(v=>`<option value="${esc(v)}">${esc(id==='period'?monthName(v):v)}</option>`).join('');if(values.includes(old))$('#'+id).value=old;
 }
 $('#fileCount').textContent=sources.length;
 $('#sourceList').innerHTML=sources.map(s=>`<div class="source"><span><strong>${esc(s.name)}</strong><br>${availableTypes(s.rows).map(type=>`${TYPES[type]}: ${s.rows.filter(r=>r.type===type).length} satır`).join(' · ')}${s.rejected?` · ${s.rejected} satır tanınamadı`:''}</span><button data-remove="${s.hash}">Kaldır</button></div>`).join('');
 $$('[data-remove]').forEach(b=>b.onclick=()=>{sources=sources.filter(s=>s.hash!==b.dataset.remove);updateOptions();render();});
}
async function readReport(file){
 const data=await file.arrayBuffer();const hash=Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',data)),x=>x.toString(16).padStart(2,'0')).join('');
 if(sources.some(s=>s.hash===hash))return `${file.name}: daha önce eklendi, atlandı.`;
 const wb=XLSX.read(data,{type:'array',cellDates:true});const rows=[];let found=false,rejected=0;
 for(const name of wb.SheetNames){
  const matrix=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:null});
  const index=matrix.findIndex(r=>r.includes('Depo')&&(r.includes('MalzemeAdı')||r.includes('Malzeme Adı'))&&(r.includes('Çıkış Miktarı')||r.includes('Tesbit Miktarı')));
  if(index<0)continue;
  found=true;
  for(let i=index+1;i<matrix.length;i++){
   const raw=Object.fromEntries(matrix[index].map((h,j)=>[h,matrix[i][j]]));
   if(!(raw['SözleşmeSahibi']||raw['Söz. Sahibi Adı'])||!(raw['MalzemeAdı']||raw['Malzeme Adı']))continue;
   try{const type=detectType(raw);const parsed=normalizeRows([raw],type,file.name);rows.push(...parsed.map(r=>({...r,row:i+1,sourceId:hash+':'+name})));}
   catch(e){rejected++;}
  }
 }
 if(!found)throw Error(`${file.name}: ham hareket sayfası bulunamadı.`);
 if(!rows.length)throw Error(`${file.name}: hareket türü tanınamadı veya geçerli satır bulunamadı.`);
 sources.push({name:file.name,hash,rows,rejected});return `${file.name}: ${rows.length} hareket eklendi.${rejected?` ${rejected} satır tanınamadığı veya geçersiz olduğu için dahil edilmedi; kontrol edin.`:''}`;
}
$('#reportFiles').onchange=async event=>{
 const input=event.target;
 input.disabled=true;const messages=[];
 try{for(const file of input.files){try{messages.push(await readReport(file));}catch(e){messages.push(e.message);}}}finally{input.disabled=false;input.value='';updateOptions();render();$('#importStatus').textContent=messages.join(' ');}
};
function render(){
 const rows=visible(),s=summarize(rows),scope=selection();
 $('#exportOpen').disabled=!rows.length;
 $('#scopeNote').textContent=`${scope.types.length} / ${availableTypes(allRows()).length} hareket türü seçili`;
 $('#filterSummary').textContent=[scope.periods.length?monthName(scope.periods[0]):'Tüm dönemler',scope.owner||'Tüm firmalar',scope.product,scope.tank?`Tank ${scope.tank}`:'',`${scope.types.length} hareket türü`].filter(Boolean).join(' · ');
 const ps=[...new Set(rows.map(r=>r.period))].sort();
 $('#dateRange').textContent=ps.map(monthName).join(' · ');
 const rejected=sources.reduce((sum,source)=>sum+source.rejected,0);
 $('#status').textContent=allRows().length?`${sources.length} rapor dosyası · ${s.records} hareket satırı · ${s.owners} sözleşme sahibi${rejected?` · ${rejected} satır tanınamadı veya geçersiz; analize alınmadı.`:''}${s.missingScale?` · ${s.missingScale} araç satırında kantar zamanı veya plaka eksik; ayrı sayıldı.`:''}`:'Başlamak için rapor dosyalarınızı ekleyin. Pivot gerektirmez.';
 $('#kpis').innerHTML=Object.entries(TYPES).filter(([key])=>scope.types.includes(key)&&rows.some(r=>r.type===key)).map(([key,label])=>{
  const part=rows.filter(r=>r.type===key),quantity=part.reduce((a,r)=>a+r.quantity,0),count=key==='vehicle'?summarize(part).trips:new Set(part.map(r=>key==='pipeline'?r.tank:r.plate||r.ship).filter(Boolean)).size;
  const detail=key==='vehicle'?`${count} araç seferi`:key==='pipeline'?`${count} tank`:`${count} gemi / sefer`;
  return `<article class="kpi" style="--accent:${colors[key]};--tint:${colors[key]}12"><div class="kpi-top">${label}<span class="kpi-icon">${icons[key]}</span></div><div class="kpi-value">${n(quantity/1000)}<small>ton</small></div><div class="kpi-foot"><b>${detail}</b> <span>· ${part.length} hareket satırı</span></div></article>`;
 }).join('');
 const groups=groupRows(rows,['owner']),max=groups[0]?.quantity||1;
 $('#ownerChart').innerHTML=groups.slice(0,5).map(g=>`<div class="bar-row"><div class="bar-label"><span title="${esc(g.owner)}">${esc(g.owner)}</span><b>${n(g.quantity/1000)}</b></div><div class="bar-track"><div class="bar-fill" style="width:${g.quantity/max*100}%"></div></div></div>`).join('')||'<div class="empty">Seçili kapsamda veri yok.</div>';
 const volumes=availableTypes(rows).map(type=>({type,quantity:rows.filter(r=>r.type===type).reduce((a,r)=>a+r.quantity,0)})),total=volumes.reduce((a,r)=>a+r.quantity,0);let angle=0;
 const gradient=volumes.map(v=>{const start=angle;angle+=total?v.quantity/total*360:0;return `${colors[v.type]} ${start}deg ${angle}deg`;}).join(',');
 $('#donut').style.background=total?`conic-gradient(${gradient})`:'#edf2f5';$('#volume').textContent=n(total/1000);
 $('#legend').innerHTML=volumes.map(v=>`<div class="legend-row"><span class="legend-dot" style="background:${colors[v.type]}"></span><span>${TYPES[v.type]}<small>${n(total?v.quantity/total*100:0)}%</small></span><b>${n(v.quantity/1000)}</b></div>`).join('');
 if(activeTab!=='all'&&!scope.types.includes(activeTab))activeTab='all';
 const shownTypes=availableTypes(rows);if(activeTab!=='all'&&!shownTypes.includes(activeTab))activeTab='all';
 $('#tabs').innerHTML=[['all','Genel özet'],...Object.entries(TYPES).filter(([key])=>shownTypes.includes(key))].map(([key,label])=>`<button data-tab="${key}" class="${activeTab===key?'selected':''}">${label}<small>${key==='all'?s.owners:rows.filter(r=>r.type===key).length}</small></button>`).join('');
 $$('[data-tab]').forEach(b=>b.onclick=()=>{activeTab=b.dataset.tab;expanded.clear();render();});
 renderTable(activeTab==='all'?rows:rows.filter(r=>r.type===activeTab));
}
function renderTable(rows){
 const groups=groupRows(rows,['owner']);
 $('#tableHead').innerHTML='<tr><th>SÖZLEŞME SAHİBİ / DETAY</th><th class="number">GİRİŞ (TON)</th><th class="number">ÇIKIŞ (TON)</th><th class="number">ARAÇ SEFERİ</th><th class="number">HAREKET SATIRI</th></tr>';
 $('#tableBody').innerHTML=groups.map((g,i)=>{
  const s=summarize(g.rows),open=expanded.has(g.owner);
  let html=`<tr class="owner-row"><td><button data-expand="${i}" aria-label="${esc(g.owner)} detayları" aria-expanded="${open}">${open?'−':'+'}</button>${esc(g.owner)}</td><td class="number">${n(s.incoming/1000)}</td><td class="number">${n(s.outgoing/1000)}</td><td class="number">${s.trips||'—'}</td><td class="number">${g.rows.length}</td></tr>`;
  if(open)for(const d of groupRows(g.rows,['type','declarant','destination','product','tank'])){
   const ds=summarize(d.rows);html+=`<tr class="detail-row"><td>${esc(TYPES[d.type])} · ${esc(d.declarant||'Beyan sahibi belirtilmemiş')}<br><strong>${esc(d.product)}</strong> · Tank ${esc(d.tank)}${d.type==='shipOut'?`<br>Gideceği firma: ${esc(d.destination||'Belirtilmemiş')}`:''}</td><td class="number">${n(ds.incoming/1000)}</td><td class="number">${n(ds.outgoing/1000)}</td><td class="number">${ds.trips||'—'}</td><td class="number">${d.rows.length}</td></tr>`;
  }return html;
 }).join('')||'<tr><td colspan="5" class="empty">Bu seçimde hareket bulunamadı.</td></tr>';
 const s=summarize(rows);$('#tableFoot').innerHTML=`<tr><td>GENEL TOPLAM</td><td class="number">${n(s.incoming/1000)}</td><td class="number">${n(s.outgoing/1000)}</td><td class="number">${s.trips||'—'}</td><td class="number">${rows.length}</td></tr>`;
 $('#tableCount').textContent=`${groups.length} sözleşme sahibi · Detayları açmak için + düğmesini kullanın`;
 $$('[data-expand]').forEach(b=>b.onclick=()=>{const key=groups[Number(b.dataset.expand)].owner;expanded.has(key)?expanded.delete(key):expanded.add(key);render();});
}
$$('.filters select, #typeChecks input').forEach(e=>e.onchange=()=>{expanded.clear();render();});$('#search').oninput=render;
$('#reset').onclick=()=>{$$('.filters select').forEach(e=>e.value='');$$('#typeChecks input').forEach(e=>e.checked=true);$('#search').value='';render();};
$('#exportOpen').onclick=()=>{$('#exportPeriods').innerHTML=periods().map(p=>`<label><input type="checkbox" value="${p}" ${!selection().periods.length||selection().periods.includes(p)?'checked':''}>${monthName(p)}</label>`).join('');$('#exportDialog').showModal();};
$('#download').onclick=async()=>{
 const filters=selection();const scope=$('input[name=exportScope]:checked').value;
 filters.periods=scope==='all'?[]:$$('#exportPeriods input:checked').map(e=>e.value);
 if(scope!=='all'&&!filters.periods.length){$('#status').textContent='Excel için en az bir dönem seçin.';$('#exportDialog').close();return;}
 const rows=filterRows(allRows(),filters);
 const button=$('#download');button.disabled=true;button.textContent='Excel hazırlanıyor…';
 try{
  const bytes=await buildExcelReport(rows,{...filters,periodLabel:filters.periods.length?filters.periods.map(monthName).join(', '):'Tüm dönemler'},XLSX,JSZip);
  const url=URL.createObjectURL(new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
  const link=document.createElement('a');link.href=url;link.download='Genel_Hareket_Raporu.xlsx';link.click();setTimeout(()=>URL.revokeObjectURL(url),60000);$('#exportDialog').close();
 }catch(error){$('#status').textContent='Excel hazırlanamadı: '+error.message;$('#exportDialog').close();}
 finally{button.disabled=false;button.textContent='↓ Excel raporunu indir';}

};
render();

const menuToggle=document.querySelector('#menuToggle'),sidebar=document.querySelector('#sidebar');
function setMenu(open){document.body.classList.toggle('menu-collapsed',!open);sidebar.inert=!open;menuToggle.setAttribute('aria-expanded',String(open));menuToggle.setAttribute('aria-label',open?'Menüyü kapat':'Menüyü aç');}
menuToggle.addEventListener('click',()=>setMenu(document.body.classList.contains('menu-collapsed')));
document.addEventListener('keydown',event=>{if(event.key==='Escape'&&!document.body.classList.contains('menu-collapsed')&&!document.querySelector('dialog[open]')){setMenu(false);menuToggle.focus();}});
