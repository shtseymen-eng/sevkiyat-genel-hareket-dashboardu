export const TYPES={vehicle:'Araç çıkışı',shipIn:'Gemi girişi',shipOut:'Gemi çıkışı',pipeline:'Boru hattı'};
const text=v=>String(v??'').trim();
export function availableTypes(rows){return Object.keys(TYPES).filter(type=>rows.some(r=>r.type===type));}
export function detectType(r){
 if(Object.hasOwn(r,'Tesbit Miktarı'))return 'shipIn';
 const plate=text(r.Plaka1||r['Plaka No']).toLocaleUpperCase('tr-TR');
 const trailer=text(r.Plaka2).toLocaleUpperCase('tr-TR');
 if(/BORU\s*HAT|BHS/.test(plate+' '+trailer))return 'pipeline';
 if(/^(0[1-9]|[1-7][0-9]|8[01])\s*[A-ZÇĞİÖŞÜ]{1,3}\s*\d{2,6}(?:\s*-|\s*$)/u.test(plate))return 'vehicle';
 if(/[A-ZÇĞİÖŞÜ]/u.test(plate)&&/(?:-|\/)\s*\d{6}\s*-*\s*$/.test(plate))return 'shipOut';
 throw Error('Hareket türü tanınamadı; plaka veya taşıma bilgisi kontrol edilmeli.');
}
export function dateText(v){
 if(v instanceof Date) return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,'0')}-${String(v.getDate()).padStart(2,'0')}T${String(v.getHours()).padStart(2,'0')}:${String(v.getMinutes()).padStart(2,'0')}:${String(v.getSeconds()).padStart(2,'0')}.${String(v.getMilliseconds()).padStart(3,'0')}`;
 if(typeof v==='number')return new Date(Math.round((v-25569)*86400000)).toISOString().replace('Z','');
 const s=text(v); if(/^\d{4}-\d\d-\d\d/.test(s))return s;
 const m=s.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})(.*)$/);return m?`${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}${m[4]}`:'';
}
export function normalizeRows(raw,type,source=''){
 return raw.flatMap((r,i)=>{
  const owner=text(r['SözleşmeSahibi']??r['Söz. Sahibi Adı']);
  const product=text(r['MalzemeAdı']??r['Malzeme Adı']);
  if(!owner||!product)return [];
  const q=r['Çıkış Miktarı']??r['Tesbit Miktarı'];
  const quantity=typeof q==='number'?q:Number(text(q).replace(/\./g,'').replace(',','.'));
  if(!Number.isFinite(quantity))throw Error(`${source}: ${i+2}. satırda miktar okunamadı.`);
  const date=dateText(r['İşlem Tarihi']);
  if(!date||Number(date.slice(0,4))<2000)throw Error(`${source}: ${i+2}. satırda işlem tarihi okunamadı.`);
  const scale=dateText(r['Kantar Çıkış']);
  return [{type,owner,product,quantity,date,period:date.slice(0,7),tank:text(r.Depo),declarant:text(r['Beyan Sahibi']??r['Byn. Sahibi Adı']),destination:text(r['Gideceği Firma']),plate:text(r.Plaka1??r['Plaka No']).replace(/\s/g,'').toLocaleUpperCase('tr-TR'),scale:Number(scale.slice(0,4))>=2000?scale:'',ship:text(r['Taşıt Adı']),document:text(r['Belge No']),declaration:text(r['Beyanname Numarası']??r['Ref. Beyan']),source,row:i+2}];
 });
}
export function filterRows(rows,{periods=[],types=Object.keys(TYPES),owner='',product='',tank='',query=''}={}){
 const q=query.toLocaleLowerCase('tr-TR');return rows.filter(r=>types.includes(r.type)&&(!periods.length||periods.includes(r.period))&&(!owner||r.owner===owner)&&(!product||r.product===product)&&(!tank||r.tank===tank)&&(!q||[r.owner,r.declarant,r.product,r.tank,r.plate,r.ship,r.destination].join(' ').toLocaleLowerCase('tr-TR').includes(q)));
}
export function summarize(rows){
 const trips=new Set();let incoming=0,outgoing=0,missingScale=0;
 for(const r of rows){if(r.type==='shipIn')incoming+=r.quantity;else outgoing+=r.quantity;
 if(r.type==='vehicle'){if(!r.scale||!r.plate)missingScale++;trips.add(r.scale&&r.plate?JSON.stringify([r.plate,r.tank,r.scale]):JSON.stringify([r.sourceId||r.source,r.row]));}}
 return {incoming,outgoing,trips:trips.size,missingScale,records:rows.length,owners:new Set(rows.map(r=>r.owner)).size};
}
export function groupRows(rows,keys){
 const map=new Map();for(const r of rows){const key=JSON.stringify(keys.map(k=>r[k]));if(!map.has(key))map.set(key,{...Object.fromEntries(keys.map(k=>[k,r[k]])),quantity:0,rows:[]});const g=map.get(key);g.quantity+=r.quantity;g.rows.push(r);}
 return [...map.values()].sort((a,b)=>b.quantity-a.quantity);
}
