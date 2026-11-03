import test from 'node:test';
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
const require=createRequire(import.meta.url);
const XLSX=require('../vendor/xlsx.full.min.js');
const JSZip=require('../vendor/jszip.min.js');
import {buildExcelReport} from '../export.mjs';
test('export separates summaries from records, retains totals and writes visible formatting',async()=>{
 const rows=[{type:'vehicle',owner:'Firma A',declarant:'Beyan A',product:'Ürün',tank:'90',quantity:1000,plate:'34ABC123',scale:'2026-11-03T10:00:00',date:'2026-11-03T10:00:00',source:'Rapor.xlsx',row:2},{type:'vehicle',owner:'Firma A',declarant:'Beyan A',product:'Ürün',tank:'90',quantity:2000,plate:'34ABC123',scale:'2026-11-03T10:00:00',date:'2026-11-03T10:00:00',source:'Rapor.xlsx',row:3}];
 const bytes=await buildExcelReport(rows,{periodLabel:'Kasım 2026'},XLSX,JSZip);
 const book=XLSX.read(bytes,{type:'array'});assert.deepEqual(book.SheetNames,['Genel Dashboard','Araç çıkışı','Hareket Kayıtları']);
 const summary=XLSX.utils.sheet_to_json(book.Sheets['Araç çıkışı'],{header:1});
 assert.deepEqual(summary.find(r=>r[0]==='GENEL TOPLAM'),['GENEL TOPLAM','','','',3,1,2]);
 const zip=await JSZip.loadAsync(bytes);const xml=await zip.file('xl/worksheets/sheet2.xml').async('string');
 assert.match(xml,/state="frozen"/);assert.match(xml,/s="[1-9]/);assert.match(xml,/<mergeCell/);
 assert.ok(xml.indexOf('<pageSetup')<xml.indexOf('<ignoredErrors'),'print settings must precede ignoredErrors in worksheet schema order');
 const styles=await zip.file('xl/styles.xml').async('string');assert.match(styles,/102941/);assert.match(styles,/087E83/);
 const raw=XLSX.utils.sheet_to_json(book.Sheets['Hareket Kayıtları'],{header:1});assert.equal(raw.length,7);
});
