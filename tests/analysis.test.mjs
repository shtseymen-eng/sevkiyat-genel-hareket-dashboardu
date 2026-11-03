import test from 'node:test';
import assert from 'node:assert/strict';
import { normalizeRows, summarize, filterRows, groupRows, detectType, availableTypes } from '../analysis.mjs';
const base={'Depo':'90','SözleşmeSahibi':'Firma A','Beyan Sahibi':'Beyan A','MalzemeAdı':'HEXANE','Plaka1':'34ABC123','Çıkış Miktarı':100,'İşlem Tarihi':'2026-11-03T10:32:00','Kantar Çıkış':'2026-11-03T10:30:00'};
test('content identifies movement regardless of identical report filenames',()=>{
 assert.equal(detectType(base),'vehicle');
 assert.equal(detectType({...base,Plaka1:'SST260027969BHS',Plaka2:'BORUHAT'}),'pipeline');
 assert.equal(detectType({...base,Plaka1:'SUMER - 260004','Kantar Çıkış':null}),'shipOut');
 assert.equal(detectType({'Tesbit Miktarı':100,'Malzeme Adı':'XYLENE'}),'shipIn');
 assert.throws(()=>detectType({...base,Plaka1:'', 'Plaka No':''}),/tanınamadı/);
});
test('only movement types actually present are available',()=>{
 assert.deepEqual(availableTypes([{type:'pipeline'},{type:'vehicle'},{type:'pipeline'}]),['vehicle','pipeline']);
 assert.deepEqual(availableTypes([]),[]);
});
test('same plate tank and timestamp counts once while quantities add; later trip remains separate',()=>{
 const rows=normalizeRows([base,{...base,'Çıkış Miktarı':200},{...base,'Kantar Çıkış':'2026-11-03T14:30:00'}],'vehicle');
 assert.equal(summarize(rows).trips,2); assert.equal(summarize(rows).outgoing,400);
});
test('missing scale timestamp never silently merges trips',()=>{
 const rows=normalizeRows([{...base,'Kantar Çıkış':null},{...base,'Kantar Çıkış':null}],'vehicle');
 assert.equal(summarize(rows).trips,2); assert.equal(summarize(rows).missingScale,2);
});
test('report total rows are excluded and dates filter by month',()=>{
 const rows=normalizeRows([base,{...base,'İşlem Tarihi':'2026-10-30T10:00:00'},{'Plaka No':'Çıkan Toplam Miktar','Çıkış Miktarı':200}],'pipeline');
 assert.equal(rows.length,2); assert.equal(filterRows(rows,{periods:['2026-11'],types:['pipeline']}).length,1);
 assert.equal(filterRows(rows,{types:[]}).length,0);
});
test('ship entry aliases and contract product tank totals preserve quantities',()=>{
 const rows=normalizeRows([{'Depo':79,'Söz. Sahibi Adı':'Ravago','Byn. Sahibi Adı':'Ravago','Malzeme Adı':'STRENE','Tesbit Miktarı':120,'İşlem Tarihi':'2026-10-30'}, {'Depo':80,'Söz. Sahibi Adı':'Ravago','Byn. Sahibi Adı':'Ravago','Malzeme Adı':'STRENE','Tesbit Miktarı':80,'İşlem Tarihi':'2026-10-30'}],'shipIn');
 assert.equal(summarize(rows).incoming,200); assert.equal(groupRows(rows,['owner'])[0].quantity,200);
 assert.equal(groupRows(rows,['owner','tank']).length,2);
});
