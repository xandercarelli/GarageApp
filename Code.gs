// Google Sheets: Estensioni > Apps Script. Incolla, esegui setup() una volta, poi
// Distribuisci > Gestisci distribuzioni > matita > Versione: Nuova versione > Distribuisci.
// Esegui come: te stesso; Chi ha accesso: Chiunque. Il codice di accesso sta in Config!B2.
const SH={Veicoli:['id','tipo','marca','modello','anno','km','colore'],
 Rifornimenti:['id','veicolo','datetime','km','litri','prezzo','totale','note'],
 Manutenzioni:['id','veicolo','data','km','tipo','descrizione','costo','note'],
 Config:['chiave','valore']};
function sh(n){const ss=SpreadsheetApp.getActive();let s=ss.getSheetByName(n)||ss.insertSheet(n);
 if(s.getLastRow()==0){s.appendRow(SH[n]);
  if(n=='Config')s.appendRow(['codice','1234']);
  if(n=='Rifornimenti'||n=='Manutenzioni')s.getRange('C:C').setNumberFormat('@');}
 return s;}
function setup(){Object.keys(SH).forEach(n=>{const s=sh(n);s.getRange(1,1,1,SH[n].length).setValues([SH[n]])});}
const out=o=>ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
// Accesso: codice obbligatorio; dopo 10 tentativi errati si blocca tutto per 15 minuti
const okCode=c=>{const k=String(sh('Config').getRange('B2').getValue()),ch=CacheService.getScriptCache(),n=+ch.get('fail')||0;
 if(!k||!c||n>=10)return false;
 if(String(c)===k){ch.remove('fail');return true}
 ch.put('fail',String(n+1),900);return false};
function rows(n){const v=sh(n).getDataRange().getValues(),h=v.shift();
 return v.map(r=>Object.fromEntries(h.map((k,i)=>[k,r[i] instanceof Date?r[i].toISOString():r[i]])));}
function doGet(e){try{
 if(!okCode(e.parameter.k))return out({error:'codice'});
 return out({veicoli:rows('Veicoli'),rifornimenti:rows('Rifornimenti'),manutenzioni:rows('Manutenzioni')});
}catch(err){return out({error:String(err)});}}
function doPost(e){let L;try{
 const d=JSON.parse(e.postData.contents);if(!okCode(d.c))return out({error:'codice'});
 if(!SH[d.sheet]||d.sheet=='Config')return out({error:'foglio'});
 L=LockService.getScriptLock();L.waitLock(10000);
 const s=sh(d.sheet),h=SH[d.sheet];
 if(d.op=='add'){d.row.id=Utilities.getUuid().slice(0,8);s.appendRow(h.map(k=>d.row[k]??''));}
 else{const ids=s.getRange(1,1,s.getLastRow(),1).getValues().map(r=>String(r[0])),i=ids.indexOf(String(d.row.id));
  if(i<1)return out({error:'non trovato'});
  if(d.op=='del')s.deleteRow(i+1);else s.getRange(i+1,1,1,h.length).setValues([h.map(k=>d.row[k]??'')]);}
 return out({ok:1});
}catch(err){return out({error:String(err)});}finally{if(L)L.releaseLock();}}
