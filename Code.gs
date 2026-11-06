// Google Sheets: Estensioni > Apps Script. Incolla, esegui setup() una volta, poi Distribuisci > Nuova distribuzione > App web
// (Esegui come: te stesso; Chi ha accesso: Chiunque). Il codice di accesso sta in Config!B2.
const SH={Veicoli:['id','tipo','marca','modello','anno','km'],
 Rifornimenti:['id','veicolo','datetime','km','litri','prezzo','totale','note'],
 Manutenzioni:['id','veicolo','data','km','tipo','descrizione','costo','note'],
 Config:['chiave','valore']};
function setup(){const ss=SpreadsheetApp.getActive();
 for(const n in SH){const s=ss.getSheetByName(n)||ss.insertSheet(n);if(s.getLastRow()==0)s.appendRow(SH[n]);}
 ss.getSheetByName('Rifornimenti').getRange('C:C').setNumberFormat('@');
 ss.getSheetByName('Manutenzioni').getRange('C:C').setNumberFormat('@');
 const c=ss.getSheetByName('Config');if(c.getLastRow()<2)c.appendRow(['codice','1234']);}
const out=o=>ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const okCode=c=>String(c)===String(SpreadsheetApp.getActive().getSheetByName('Config').getRange('B2').getValue());
function rows(n){const v=SpreadsheetApp.getActive().getSheetByName(n).getDataRange().getValues(),h=v.shift();
 return v.map(r=>Object.fromEntries(h.map((k,i)=>[k,r[i] instanceof Date?r[i].toISOString():r[i]])));}
function doGet(e){if(!okCode(e.parameter.c))return out({error:'codice'});
 return out({veicoli:rows('Veicoli'),rifornimenti:rows('Rifornimenti'),manutenzioni:rows('Manutenzioni')});}
function doPost(e){const d=JSON.parse(e.postData.contents);if(!okCode(d.c))return out({error:'codice'});
 if(!SH[d.sheet]||d.sheet=='Config')return out({error:'foglio'});
 const L=LockService.getScriptLock();L.waitLock(10000);
 try{const s=SpreadsheetApp.getActive().getSheetByName(d.sheet),h=SH[d.sheet];
  if(d.op=='add'){d.row.id=Utilities.getUuid().slice(0,8);s.appendRow(h.map(k=>d.row[k]??''));}
  else{const ids=s.getRange(1,1,s.getLastRow(),1).getValues().map(r=>String(r[0])),i=ids.indexOf(String(d.row.id));
   if(i<1)return out({error:'non trovato'});
   if(d.op=='del')s.deleteRow(i+1);else s.getRange(i+1,1,1,h.length).setValues([h.map(k=>d.row[k]??'')]);}
  return out({ok:1});}finally{L.releaseLock();}}
