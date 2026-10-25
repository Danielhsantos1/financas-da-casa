/** @OnlyCurrentDoc */
// Finanças da Casa - ponte entre o app e esta planilha.
// 1) Troque TOKEN por uma senha sua (qualquer texto longo, sem espaços).
// 2) Implantar > Nova implantação > App da Web > Executar como: Eu > Acesso: Qualquer pessoa.
const TOKEN = 'TROQUE-ESTE-TEXTO';

const TABS = {
  contas:   ['id','nome','valor','tipo','de','ate','dia'],
  lanc:     ['id','conta','ym','pago','valor'],
  gastos:   ['id','desc','valor','cat','quem','data'],
  entradas: ['id','valor']
};
const TEXT_COLS = ['id','conta','ym','de','ate','data'];

function out_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function sheet_(tab){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(tab);
  if(!sh){ sh = ss.insertSheet(tab); sh.appendRow(TABS[tab]); }
  const head = TABS[tab];
  head.forEach((h,i)=>{ if(TEXT_COLS.indexOf(h)>=0) sh.getRange(1,i+1,sh.getMaxRows(),1).setNumberFormat('@'); });
  return sh;
}

function norm_(v){
  if(v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM');
  return v === '' ? null : v;
}

function readTab_(tab){
  const sh = sheet_(tab), head = TABS[tab];
  const last = sh.getLastRow();
  if(last < 2) return [];
  const vals = sh.getRange(2,1,last-1,head.length).getValues();
  return vals.filter(r=>r[0] !== '').map(r=>{ const o={}; head.forEach((h,i)=>o[h]=norm_(r[i])); return o; });
}

function doGet(e){
  if(!e || !e.parameter || e.parameter.t !== TOKEN) return out_({erro:'acesso negado'});
  const d = {}; Object.keys(TABS).forEach(t=>d[t]=readTab_(t));
  return out_(d);
}

function doPost(e){
  let b; try{ b = JSON.parse(e.postData.contents); }catch(err){ return out_({erro:'json invalido'}); }
  if(b.t !== TOKEN) return out_({erro:'acesso negado'});
  if(!TABS[b.tab]) return out_({erro:'aba invalida'});
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const sh = sheet_(b.tab), head = TABS[b.tab];
    const last = sh.getLastRow();
    const ids = last>=2 ? sh.getRange(2,1,last-1,1).getValues().map(r=>String(r[0])) : [];
    const pos = ids.indexOf(String(b.id));
    if(b.op === 'delete'){ if(pos>=0) sh.deleteRow(pos+2); }
    else if(b.op === 'upsert'){
      const row = head.map(h=> h==='id' ? b.id : (b.data && b.data[h] != null ? b.data[h] : ''));
      if(pos>=0) sh.getRange(pos+2,1,1,head.length).setValues([row]);
      else sh.getRange(last+1,1,1,head.length).setValues([row]);
    } else return out_({erro:'op invalida'});
    rebuildView_();
    return out_({ok:true});
  } finally { lock.release(); }
}

// ---- Aba "Visao": grade mensal pronta para olhar no Google Sheets ou puxar no Excel ----
function addM_(s,n){ let [y,m]=s.split('-').map(Number); m+=n-1; y+=Math.floor(m/12); m=((m%12)+12)%12; return y+'-'+('0'+(m+1)).slice(-2); }
function rebuildView_(){
  const contas=readTab_('contas'), lanc=readTab_('lanc'), gastos=readTab_('gastos'), ent=readTab_('entradas');
  if(!contas.length) return;
  const L={}; lanc.forEach(l=>L[l.id]=l);
  const first=contas.map(c=>c.de).concat(ent.map(e=>e.id)).filter(Boolean).sort()[0];
  let last=contas.map(c=>c.ate).filter(Boolean).sort().pop()||first;
  if(last<addM_(first,11)) last=addM_(first,11);
  const ms=[]; for(let m=first; m<=last && ms.length<36; m=addM_(m,1)) ms.push(m);
  const E={}; ent.forEach(e=>E[e.id]=e.valor);
  const entradaDe=m=>{ if(E[m]!=null) return E[m]; const ks=Object.keys(E).filter(k=>k<m).sort(); return ks.length?E[ks.pop()]:0; };
  const valor=(c,m)=>{
    if(!(c.de<=m && (!c.ate || m<=c.ate))) return null;
    const l=L[c.id+'_'+m];
    if(l && l.valor===0 && !l.pago) return null;
    return (l && l.valor!=null) ? l.valor : c.valor;
  };
  const rows=[['Conta'].concat(ms)], paid=[], fmt=[];
  const tot=ms.map(()=>0);
  ['pagamento','vale'].forEach(tp=>{
    rows.push([tp==='pagamento'?'PAGAMENTOS':'VALES / OUTROS'].concat(ms.map(()=>''))); paid.push(null);
    contas.filter(c=>c.tipo===tp).sort((a,b)=>String(a.nome).localeCompare(String(b.nome))).forEach(c=>{
      const r=[c.nome], p=[''];
      ms.forEach((m,i)=>{ const v=valor(c,m); r.push(v==null?'':v); if(v!=null) tot[i]+=v; const l=L[c.id+'_'+m]; p.push(v!=null && l && l.pago ? '#b6e3c5' : null); });
      rows.push(r); paid.push(p);
    });
  });
  const gv=ms.map(m=>gastos.filter(g=>String(g.data||'').slice(0,7)===m).reduce((a,g)=>a+(Number(g.valor)||0),0));
  rows.push(['TOTAL DE CONTAS'].concat(tot.map(x=>Math.round(x*100)/100)));
  rows.push(['Entrada'].concat(ms.map(entradaDe)));
  rows.push(['Gastos variáveis'].concat(gv));
  rows.push(['SOBRA DO MÊS'].concat(ms.map((m,i)=>Math.round((entradaDe(m)-tot[i]-gv[i])*100)/100)));
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sh=ss.getSheetByName('Visao')||ss.insertSheet('Visao',0);
  sh.clear();
  sh.getRange(1,1,rows.length,rows[0].length).setValues(rows);
  sh.getRange(2,2,rows.length-1,ms.length).setNumberFormat('"R$" #,##0.00');
  sh.getRange(1,1,1,rows[0].length).setBackground('#00607f').setFontColor('#ffffff').setFontWeight('bold');
  paid.forEach((p,i)=>{ if(p) p.forEach((c,j)=>{ if(c) sh.getRange(i+2,j+1).setBackground(c); }); });
  sh.getRange(rows.length-3,1,4,1).setFontWeight('bold');
  sh.setFrozenRows(1); sh.setFrozenColumns(1); sh.setColumnWidth(1,230);
}

// Rode esta função uma vez a mão (menu Executar) para criar a aba Visao com os dados atuais.
function montarVisao(){ rebuildView_(); }

// A carga inicial das contas (dados pessoais) fica so dentro da planilha, nunca no GitHub.
