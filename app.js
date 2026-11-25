/* Controle Financeiro — Office.js. Nenhum dado financeiro é enviado para servidores. */
'use strict';
const BUILD_ID = '6.3.0';
const $ = id => document.getElementById(id);
const popupMode = new URLSearchParams(location.search).get('popup') === '1';
let controleDialog = null, rpcNext = 0;
const rpcPending = new Map();
function dialogCall(action,payload={}) {
  return new Promise((resolve,reject) => {
    const rid = 'cf-'+(++rpcNext);
    const timeout = setTimeout(() => {rpcPending.delete(rid);reject(new Error('Tempo esgotado na comunicação com o Excel.'));},45000);
    rpcPending.set(rid,{resolve,reject,timeout});
    try { Office.context.ui.messageParent(JSON.stringify({kind:'cf-rpc',rid,action,payload})); }
    catch(err){clearTimeout(timeout);rpcPending.delete(rid);reject(err);}
  });
}
function replyFromHost(arg) {
  let msg; try {msg=JSON.parse(arg.message);} catch {return;}
  if(msg.kind!=='cf-rpc-result')return;
  const slot=rpcPending.get(msg.rid);if(!slot)return;
  clearTimeout(slot.timeout);rpcPending.delete(msg.rid);
  if(msg.build!==BUILD_ID)console.warn('[Controle Financeiro] Versões diferentes no Excel e na janela.', {host:msg.build,dialog:BUILD_ID});
  if(msg.error)slot.reject(new Error(msg.error));else slot.resolve(msg.result);
}
async function messageFromDialog(arg) {
  if(arg.origin && arg.origin!==location.origin)return;
  let msg;try{msg=JSON.parse(arg.message);}catch{return;}
  if(msg.kind!=='cf-rpc')return;
  const p=msg.payload||{};let result=null,error=null;
  try {
    switch(msg.action) {
      case 'readAll':result=await readAll(p.month);break;
      case 'savePlanning': await savePlanning(p.item); break;
      case 'upsertEntry':
        if(String(p.item?.id||'').startsWith('__r_'))db=await readAll(p.month);
        await upsertEntry(p.item,p.insertOnly);break;
      case 'deleteEntry':
        if(String(p.id||'').startsWith('__r_'))db=await readAll(p.month);
        await deleteEntry(p.id);break;
      case 'upsertBudgets':await upsertBudgets(p.changes,p.month);break;
      case 'upsertClosing':await upsertClosing(p.snapshot);break;
      case 'upsertRecurring':await upsertRecurring(p.rule);break;
      case 'deleteRecurring':await deleteRecurring(p.id);break;
      default:throw new Error('Operação não permitida.');
    }
  }catch(e){logExcelError(msg.action,e);error=e.message||String(e);}
  if(controleDialog){try{controleDialog.messageChild(JSON.stringify({kind:'cf-rpc-result',rid:msg.rid,result,error,build:BUILD_ID}));}catch(e){logExcelError('responder ao diálogo',e);}}
}
let dialogOpening = false;
function openSeparateWindow() {
  if(popupMode)return Promise.resolve();
  if(controleDialog){notify('A janela Controle Financeiro já está aberta.');return Promise.resolve();}
  if(dialogOpening)return Promise.resolve();
  if(!Office.context?.requirements?.isSetSupported?.('DialogApi','1.2')){
    const error=new Error('Esta instalação do Office não oferece DialogApi 1.2.');
    notify(error.message);return Promise.reject(error);
  }
  dialogOpening=true;
  return new Promise((resolve,reject)=>{
    try{
      Office.context.ui.displayDialogAsync(location.origin+location.pathname+'?popup=1&build='+encodeURIComponent(BUILD_ID),{height:98,width:98},result=>{
        dialogOpening=false;
        if(result.status!==Office.AsyncResultStatus.Succeeded){
          const error=new Error('Não foi possível abrir a janela: '+(result.error?.message||'erro desconhecido'));
          notify(error.message);reject(error);return;
        }
        controleDialog=result.value;
        controleDialog.addEventHandler(Office.EventType.DialogMessageReceived,messageFromDialog);
        controleDialog.addEventHandler(Office.EventType.DialogEventReceived,()=>{controleDialog=null;refresh();});
        resolve();
      });
    }catch(error){dialogOpening=false;notify(error.message||String(error));reject(error);}
  });
}

const money = v => new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL'}).format(Number(v)||0);
const compact = v => new Intl.NumberFormat('pt-BR',{notation:'compact',maximumFractionDigits:1}).format(Number(v)||0);
const safe = value => String(value??'').replace(/[&<>"']/g,k=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[k]));
const fmtDate = str => {const a=String(str||'').split('-');return a.length===3?`${a[2]}/${a[1]}/${a[0]}`:String(str||'');};
const fmtMonth = key => /^\d{4}-\d{2}$/.test(key||'') ? `${key.slice(5,7)}/${key.slice(0,4)}` : String(key||'');
const todayISO = () => {const d=new Date();return [d.getFullYear(),String(d.getMonth()+1).padStart(2,'0'),String(d.getDate()).padStart(2,'0')].join('-');};
const excelDate = s => {const [y,m,d] = s.split('-').map(Number); return Math.round((Date.UTC(y,m-1,d)-Date.UTC(1899,11,30))/86400000);};
const isoDate = (value,display='') => {
  if(typeof value==='number' && Number.isFinite(value)) return new Date(Math.round(value)*86400000+Date.UTC(1899,11,30)).toISOString().slice(0,10);
  let s=String(value??'').trim();
  if(/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0,10);
  const d=s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/) || String(display).match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
  return d ? `${d[3]}-${d[2].padStart(2,'0')}-${d[1].padStart(2,'0')}` : '';
};
const xlsSafe=s=>{s=String(s??'').trim();return /^[=+\-@]/.test(s)?"'"+s:s;};
const monthKey=x=>String(x.date||'').slice(0,7);
const sum=(arr,type)=>arr.filter(x=>x.type===type).reduce((v,x)=>v+x.amount,0);
const categorySpend=items=>items.filter(x=>x.type==='Despesa').reduce((acc,x)=>(acc[x.category]=(acc[x.category]||0)+x.amount,acc),{});
const colors=['#199672','#e6a18d','#e4c875','#759cdd','#a18fd7','#7bbaa7','#d47b9c','#9ea6a0'];
let connected=false, db={items:[],categories:[],budgets:{},closings:[],recurring:[]}, screen='overview', lastRead=0;
let pendingRead = Promise.resolve();
let analysisCategory='', analysisDay=0;
const TABLES={items:'tblLancamentos',categories:'tblCategorias',budgets:'tblOrcamentos',closings:'tblFechamentos',recurring:'tblRecorrentes'};
const currentMonth=()=>String($('month').value||'');
const currentItems=()=>db.items.filter(x=>monthKey(x)===currentMonth());
function notify(msg){$('toast').textContent=msg;$('toast').classList.add('show');clearTimeout(notify.timeout);notify.timeout=setTimeout(()=>$('toast').classList.remove('show'),5000);}

const recurringHeaders=['ID','Tipo','Descricao','Valor','Dia','Conta','Inicio','Fim','AutoPago','Ativo'];
function ruleRow(r){return [
  xlsSafe(r.id),xlsSafe(r.kind),xlsSafe(r.description),Number(r.amount),Number(r.day),
  xlsSafe(r.account),String(r.start),String(r.end||''),r.autoPaid?'Sim':'Não',r.active===false?'Não':'Sim'
];}
function parseRule(row){return {
  id:String(row[0]||'').trim(),kind:String(row[1]||'').trim(),description:String(row[2]||'').trim(),
  amount:Number(row[3])||0,day:Number(row[4])||1,account:String(row[5]||'').trim(),
  start:String(row[6]||'').trim(),end:String(row[7]||'').trim(),
  autoPaid:String(row[8]||'').toLowerCase()==='sim',
  active:String(row[9]||'').toLowerCase()!=='não'
};}
function recurringDate(month,day){
  const [y,m]=month.split('-').map(Number);
  const last=new Date(Date.UTC(y,m,0)).getUTCDate();
  return month+'-'+String(Math.min(Math.max(1,day),last)).padStart(2,'0');
}
function recurringId(id,month){return 'REC-'+id+'-'+month;}
function logExcelError(stage, error) {
  // Não registrar valores, mensagens do Excel ou instruções da operação.
  console.error('[Controle Financeiro]', {
    build: BUILD_ID,
    runtime: popupMode ? 'janela' : 'host',
    stage,
    name: error.name,
    code: error.code,
    location: error.debugInfo?.errorLocation,
    stack: error.stack?.split('\n').filter(line => /^\s*at\s/.test(line)).join('\n')
  });
}

function hasRecurringHeaders(row) {
  return row.length === recurringHeaders.length &&
    row.every((value, index) => value === recurringHeaders[index]);
}

async function ensureRecurringTable() {
  let stage = 'localizar tblRecorrentes';
  try {
    await Excel.run(async context => {
      const found = context.workbook.tables.getItemOrNullObject(TABLES.recurring);
      await context.sync();
      if (!found.isNullObject) return;

      stage = 'localizar aba Recorrentes';
      let sheet = context.workbook.worksheets.getItemOrNullObject('Recorrentes');
      await context.sync();
      if (sheet.isNullObject) {
        stage = 'worksheets.add / sync';
        sheet = context.workbook.worksheets.add('Recorrentes');
        await context.sync();
      }

      stage = 'inspecionar área e tabelas existentes';
      const used = sheet.getUsedRangeOrNullObject(true);
      const headers = sheet.getRange('A1:J1');
      headers.load(['values', 'formulas']);
      used.load(['columnIndex', 'columnCount', 'rowIndex', 'rowCount']);
      sheet.tables.load('items/name');
      await context.sync();

      const row = headers.values[0];
      const matchingHeaders = hasRecurringHeaders(row);
      const blankHeaders = headers.formulas[0].every(value => value === '');
      const lastRow = used.isNullObject ? 0 : used.rowIndex + used.rowCount;
      const safeArea = used.isNullObject || used.columnIndex + used.columnCount <= recurringHeaders.length;

      if (sheet.tables.items.length === 1 && matchingHeaders) {
        const candidate = sheet.tables.items[0];
        const range = candidate.getRange();
        range.load(['rowIndex', 'columnIndex', 'columnCount']);
        await context.sync();
        if (range.rowIndex === 0 && range.columnIndex === 0 && range.columnCount === recurringHeaders.length) {
          stage = 'recuperar nome da tabela / sync';
          candidate.name = TABLES.recurring;
          await context.sync();
          return;
        }
      }
      if (sheet.tables.items.length || !safeArea || (!matchingHeaders && (!blankHeaders || lastRow > 1))) {
        throw new Error('A aba Recorrentes contém uma estrutura que precisa de revisão.');
      }

      if (!matchingHeaders) {
        stage = 'escrever cabeçalhos / sync';
        headers.values = [recurringHeaders];
        await context.sync();
      }
      stage = 'tables.add / sync';
      const table = sheet.tables.add(sheet.getRange('A1:J' + Math.max(2, lastRow)), true);
      await context.sync();
      stage = 'nomear tblRecorrentes / sync';
      table.name = TABLES.recurring;
      await context.sync();
    });
  } catch (error) {
    logExcelError(stage, error);
    throw error;
  }
}

function rowsFromTable(table){const values=table.values||[],texts=table.text||[];return values.slice(1).map((row,i)=>({row,text:texts[i+1]||[]}));}
async function readAll(monthArg=currentMonth()){
  if(popupMode)return dialogCall('readAll',{month:monthArg});
  const read = pendingRead.then(() => readWorkbook(monthArg));
  pendingRead = read.catch(() => {});
  return read;
}

async function readWorkbook(monthArg) {
  if(!connected)throw Error('Abra este painel como suplemento dentro do Excel, na pasta de trabalho Controle Financeiro.xlsx.');
  const data = await Excel.run(async context=>{
    const ranges={};
    for(const [key,name] of Object.entries(TABLES)){
      if (key === 'recurring') continue;
      ranges[key]=context.workbook.tables.getItem(name).getRange();
      ranges[key].load(['values','text']);
    }
    await context.sync();
    const items=rowsFromTable(ranges.items).map(({row,text},index)=>({
      id:String(row[0]??'').trim()||`__r_${index}`,sourceIndex:index,date:isoDate(row[1],text[1]),type:String(row[2]??'').trim(),
      description:String(row[3]??'').trim(),category:String(row[4]??'').trim(),account:String(row[5]??'').trim(),
      amount:Number(row[6])||0,status:String(row[7]??'').trim()||'Pago',notes:String(row[8]??'').trim()
    })).filter(x=>x.date && (x.type==='Receita'||x.type==='Despesa') && x.amount>0);
    const categories=rowsFromTable(ranges.categories).map(({row})=>({name:String(row[0]??'').trim(),type:String(row[1]??'').trim()})).filter(x=>x.name);
    const month=monthArg;
    const budgets={};
    rowsFromTable(ranges.budgets).forEach(({row})=>{if(String(row[0]??'').trim()===month)budgets[String(row[1]??'').trim()]=Number(row[2])||0;});
    const closings=rowsFromTable(ranges.closings).map(({row})=>({month:String(row[0]??'').trim(),income:Number(row[1])||0,expense:Number(row[2])||0,balance:Number(row[3])||0,pending:Number(row[4])||0,closedAt:String(row[5]??''),notes:String(row[6]??'')})).filter(x=>x.month);
    const entryIds = rowsFromTable(ranges.items).map(({row}) => String(row[0] ?? '').trim());
    return {items,categories,budgets,closings,entryIds};
  });

  const {entryIds, ...result} = data;
  result.recurring = [];
  let stage = 'preparar tabela de receitas fixas';
  try {
    await ensureRecurringTable();
    stage = 'ler tabela de receitas fixas / sync';
    const recurring = await Excel.run(async context => {
      const table = context.workbook.tables.getItem(TABLES.recurring);
      const range = table.getRange();
      range.load('values');
      await context.sync();
      if (!hasRecurringHeaders(range.values[0])) {
        throw new Error('Os cabeçalhos de tblRecorrentes precisam de revisão.');
      }
      return rowsFromTable(range).map(({row}) => parseRule(row))
        .filter(rule => rule.id && rule.amount > 0 && rule.description && rule.start);
    });
    result.recurring = recurring;
    stage = 'gerar lançamentos recorrentes / sync';
    const month = monthArg;
    const existing=new Set(entryIds);
    const generated=[];
    for(const rule of recurring){
      if(data.closings.some(closing => closing.month === month)||!/^\d{4}-\d{2}$/.test(month)||rule.active===false||month<rule.start||(rule.end&&month>rule.end))continue;
      const id=recurringId(rule.id,month);
      if(existing.has(id))continue;
      const date=recurringDate(month,rule.day);
      const kind=rule.kind;
      const entry={id,date,type:'Receita',description:rule.description,category:kind,account:rule.account||(kind==='Vale-alimentação'?'Cartão VA':'Conta principal'),amount:rule.amount,status:rule.autoPaid&&date<=todayISO()?'Pago':'Pendente',notes:'Gerado automaticamente por receita fixa · '+kind};
      generated.push(entry);existing.add(id);
    }
    if(generated.length){
      await Excel.run(async context => {
        context.workbook.tables.getItem(TABLES.items).rows.add(null,generated.map(entryRow));
        await context.sync();
      });
      data.items.push(...generated);
    }
  } catch (error) {
    logExcelError(stage, error);
    result.recurringWarning = 'As receitas fixas não puderam ser atualizadas. Os demais dados foram carregados.';
  }
  try {
    result.planning = await readPlanning();
  } catch (error) {
    logExcelError('ler planejamento e reserva', error);
  }
  return result;
}
async function refresh(){try{db=await readAll(); planningReady=Boolean(db.planning); planning=db.planning||emptyPlanning; lastRead=Date.now();populateCategories();render();$('recurringWarning').textContent=db.recurringWarning||'';$('recurringWarning').hidden=!db.recurringWarning;}catch(err){logExcelError('ler tabelas principais',err);notify('Erro ao ler as tabelas do Excel. Confira a versão exibida e tente Atualizar.');}}
function go(to){screen=to;document.body.dataset.screen=to;document.querySelectorAll('.screen').forEach(x=>x.classList.toggle('on',x.id===to));document.querySelectorAll('.top-nav [data-quick]').forEach(button => {
  if (button.dataset.quick === to) button.setAttribute('aria-current', 'page');
  else button.removeAttribute('aria-current');
});window.scrollTo(0,0);const pages={bills:['Despesas fixas','Escolha o valor e registre o pagamento.'],reserve:['Reserva e imóvel','Acompanhe seus aportes e planeje a entrada.'],overview:['Visão geral','Veja o mês e escolha o que quer fazer.'],transactions:['Lançamentos','Cadastre e edite sem sair do Excel.'],budgets:['Orçamento mensal','Planeje limites para cada categoria.'],closing:['Fechamento mensal','Revise e registre seu resultado.'],analysis:['Análise interativa','Explore despesas por categoria e por dia.'],recurring:['Receitas fixas','Cadastre salário, vale-alimentação e outras entradas mensais.']};$('pageTitle').textContent=pages[to][0];$('pageSub').textContent=pages[to][1];render();}
function render(){const items=currentItems(),income=sum(items,'Receita'),expense=sum(items,'Despesa'),balance=income-expense,pending=items.filter(x=>x.type==='Despesa'&&x.status==='Pendente').reduce((a,x)=>a+x.amount,0);
  $('income').textContent=money(income);$('expenses').textContent=money(expense);$('balance').textContent=money(balance);$('savings').textContent=income?`${Math.round(balance/income*100)}%`:'—';$('balance').style.color=balance<0?'#ffd1be':'white';
  const vaIncome=items.filter(x=>x.type==='Receita'&&x.category==='Vale-alimentação').reduce((n,x)=>n+x.amount,0);
  const vaSpent=items.filter(x=>x.type==='Despesa'&&/^(cartão va|cartao va|vale alimentação|vale-alimentação)$/i.test(x.account.trim())).reduce((n,x)=>n+x.amount,0);
  const cashBalance=(income-vaIncome)-(expense-vaSpent);
  $('vaIncome').textContent=money(vaIncome);
  $('vaSpent').textContent=money(vaSpent);
  $('vaBalance').textContent=money(vaIncome-vaSpent);
  $('cashBalance').textContent=money(cashBalance);
  $('cashBalance').style.color=cashBalance<0?'#c44337':'#13664f';
  $('closingIncome').textContent=money(income);$('closingExpense').textContent=money(expense);$('closingBalance').textContent=money(balance);$('closingPending').textContent=money(pending);
  renderBars();renderDonut(items,expense);renderRecent(items);renderBudgets(items);renderTransactions(items);renderClosings(income,expense,pending);renderAnalysis();renderRecurring();renderPlanning();
}
function renderBars(){const [y,m]=currentMonth().split('-').map(Number);if(!y||!m)return;const months=[];for(let i=5;i>=0;i--){const d=new Date(y,m-1-i,1),key=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`,items=db.items.filter(x=>monthKey(x)===key);months.push({key,label:d.toLocaleDateString('pt-BR',{month:'short'}).replace('.',''),inc:sum(items,'Receita'),exp:sum(items,'Despesa')});}const max=Math.max(1,...months.flatMap(x=>[x.inc,x.exp]));$('barChart').innerHTML=months.map(x=>`<div class="barcol" data-month-link="${safe(x.key)}" title="Clique para analisar este mês"><div class="bars"><div class="bar inc" title="Receitas: ${money(x.inc)}" style="height:${Math.max(1,x.inc/max*100)}%"></div><div class="bar exp" title="Despesas: ${money(x.exp)}" style="height:${Math.max(1,x.exp/max*100)}%"></div></div><span>${safe(x.label)}</span></div>`).join('');}
function renderDonut(items,expense){const data=Object.entries(categorySpend(items)).sort((a,b)=>b[1]-a[1]);$('donutTotal').textContent='R$ '+compact(expense);if(!data.length){$('donut').style.background='#e9efeb';$('categoryLegend').innerHTML='<p class="empty">Ainda não há despesas para este mês.</p>';return;}let offset=0;const stops=data.map(([cat,value],i)=>{const prev=offset;offset+=value/expense*100;return `${colors[i%colors.length]} ${prev}% ${offset}%`;});$('donut').style.background=`conic-gradient(${stops.join(',')})`;$('categoryLegend').innerHTML=data.slice(0,6).map(([cat,v],i)=>`<button class="catitem catlink" data-category-link="${safe(cat)}" title="Explorar os gastos de ${safe(cat)}"><span><i class="dot" style="background:${colors[i%colors.length]}"></i>${safe(cat)}</span><b>${Math.round(v/expense*100)}%</b></button>`).join('');}
function renderRecent(items){let rows=[...items].sort((a,b)=>b.date.localeCompare(a.date)).slice(0,5);$('recent').innerHTML=rows.length?rows.map(a=>`<div class="txrow"><div class="txicon ${a.type==='Receita'?'credit':''}">${a.type==='Receita'?'↗':'↘'}</div><div class="txmain"><strong>${safe(a.description)}</strong><small>${safe(a.category)} · ${fmtDate(a.date)} · ${safe(a.status)}</small></div><div class="txamount ${a.type==='Receita'?'credit':'debit'}">${a.type==='Receita'?'+':'−'} ${money(a.amount)}</div></div>`).join(''):'<p class="empty">Registre o primeiro lançamento para começar.</p>';}
function renderBudgets(items){const spending=categorySpend(items),cats=db.categories.filter(c=>c.type==='Despesa').map(c=>c.name),budgetEntries=cats.filter(cat=>(db.budgets[cat]||0)>0).map(cat=>[cat,db.budgets[cat]]).sort((a,b)=>(spending[b[0]]||0)/b[1]-(spending[a[0]]||0)/a[1]);$('budgetPreview').innerHTML=budgetEntries.length?budgetEntries.slice(0,5).map(([cat,lim])=>{const val=spending[cat]||0,pct=100*val/lim;return `<div class="budgetline"><div class="budgettop"><span>${safe(cat)}</span><b>${money(val)} / ${money(lim)}</b></div><div class="budgettrack"><div class="budgetfill ${pct>100?'over':''}" style="width:${Math.min(100,pct)}%"></div></div></div>`;}).join(''):'<p class="empty">Adicione limites em Orçamento.</p>';
  $('budgetEditor').innerHTML=cats.map(cat=>`<div class="budgetfield"><label>${safe(cat)}<small>Gasto no mês: ${money(spending[cat]||0)}</small></label><input type="number" min="0" step="0.01" data-budget="${safe(cat)}" value="${Number(db.budgets[cat])||0}" aria-label="Limite para ${safe(cat)}"></div>`).join('');
}
function renderTransactions(items){const search=$('search').value.toLowerCase(),type=$('typeFilter').value;const rows=items.filter(x=>(!type||x.type===type)&&`${x.description} ${x.category} ${x.account}`.toLowerCase().includes(search)).sort((a,b)=>b.date.localeCompare(a.date));$('records').innerHTML=rows.length?rows.map(a=>`<tr><td>${fmtDate(a.date)}</td><td><b>${safe(a.description)}</b></td><td>${safe(a.category)}</td><td>${safe(a.account)}</td><td class="txamount ${a.type==='Receita'?'credit':'debit'}">${a.type==='Receita'?'+':'−'} ${money(a.amount)}</td><td>${safe(a.status)}</td><td><button data-edit="${safe(a.id)}">Editar</button><button class="delete" data-delete="${safe(a.id)}">Excluir</button></td></tr>`).join(''):'<tr><td colspan="7" class="empty">Nenhum lançamento para esta competência.</td></tr>';}
function renderClosings(income,expense,pending){const m=currentMonth(),last=db.closings.find(x=>x.month===m),diff=last&&(Math.abs(last.income-income)>0.01||Math.abs(last.expense-expense)>0.01||Math.abs(last.pending-pending)>0.01);const status=$('closingStatus');status.classList.toggle('warn',!!diff || pending>0);if(!last){status.textContent=pending>0?`Este mês ainda não foi fechado. Há ${money(pending)} em despesas pendentes.`:'Este mês ainda não foi fechado. Os totais estão prontos para conferência.';}else if(diff){status.textContent=`Atenção: ${fmtMonth(m)} já foi fechado em ${last.closedAt}, mas os lançamentos mudaram depois. Confira e registre um novo fechamento se necessário.`;}else{status.textContent=`Competência ${fmtMonth(m)} fechada em ${last.closedAt}. Os totais conferem com o Excel.`;}const rows=[...db.closings].sort((a,b)=>b.month.localeCompare(a.month)).slice(0,14);$('closingHistory').innerHTML=rows.length?rows.map(x=>`<div class="historyrow"><div><strong>${fmtMonth(x.month)}</strong><small>${safe(x.closedAt)}</small></div><div>Receitas <b>${money(x.income)}</b></div><div>Despesas <b>${money(x.expense)}</b></div><div>Saldo <b>${money(x.balance)}</b></div><div>Pendente <b>${money(x.pending)}</b></div></div>`).join(''):'<p class="empty">Nenhum mês foi fechado ainda.</p>';}
function populateCategories(){const type=document.querySelector('input[name=type]:checked').value,prior=$('category').value,cats=db.categories.filter(x=>x.type===type).map(x=>x.name);$('category').innerHTML=cats.map(c=>`<option value="${safe(c)}">${safe(c)}</option>`).join('');if(cats.includes(prior))$('category').value=prior;}
function modal(id){fixedPayment=null;$('paymentBudget').hidden=true;$('entryForm').reset();$('entryId').value='';$('date').value=todayISO();$('modalTitle').textContent='Novo lançamento';if(id){const row=db.items.find(x=>x.id===id);if(!row)return notify('Lançamento não encontrado. Atualize o painel.');$('entryId').value=row.id;$('modalTitle').textContent='Editar lançamento';document.querySelector(`input[name=type][value="${row.type}"]`).checked=true;populateCategories();for(const k of ['description','amount','date','category','account','notes'])$(k).value=row[k]??'';$('status').value=row.status||'Pago';}else{populateCategories();}$('modal').classList.add('show');$('description').focus();}
function closeModal(){$('modal').classList.remove('show');}
function entryRow(x){return [xlsSafe(x.id),excelDate(x.date),x.type,xlsSafe(x.description),xlsSafe(x.category),xlsSafe(x.account),Number(x.amount),x.status,xlsSafe(x.notes)];}
function newId(){return 'FC-'+Date.now().toString(36)+'-'+Math.random().toString(36).slice(2,7);}
async function upsertEntry(item,insertOnly=false){if(popupMode)return dialogCall('upsertEntry',{item,month:currentMonth(),insertOnly});await Excel.run(async context=>{const table=context.workbook.tables.getItem(TABLES.items),range=table.getRange();range.load('values');await context.sync();const rows=range.values.slice(1);let index=rows.findIndex(row=>String(row[0])===item.id);if(insertOnly&&index>=0)throw Error('Esta conta já foi lançada. Atualize o painel para editá-la.');if(index<0 && item.id.startsWith('__r_')){const source=db.items.find(x=>x.id===item.id),i=Number(item.id.slice(4));if(source && i>=0 && i<rows.length && !String(rows[i][0]??'').trim() && isoDate(rows[i][1])===source.date && String(rows[i][3]??'').trim()===source.description){index=i;item.id=newId();}else throw Error('A tabela mudou desde a última atualização. Clique em Atualizar antes de editar.');}if(index>=0)table.rows.getItemAt(index).getRange().values=[entryRow(item)];else table.rows.add(null,[entryRow(item)]);await context.sync();});}
async function deleteEntry(id){if(popupMode)return dialogCall('deleteEntry',{id,month:currentMonth()});await Excel.run(async context=>{const table=context.workbook.tables.getItem(TABLES.items),range=table.getRange();range.load('values');await context.sync();let index=range.values.slice(1).findIndex(r=>String(r[0])===id);if(index<0 && id.startsWith('__r_')){const source=db.items.find(x=>x.id===id),i=Number(id.slice(4)),rows=range.values.slice(1);if(source&&i>=0&&i<rows.length&&!String(rows[i][0]??'').trim()&&isoDate(rows[i][1])===source.date&&String(rows[i][3]??'').trim()===source.description)index=i;}if(index<0)throw Error('Lançamento não encontrado ou planilha alterada. Atualize o painel.');table.rows.getItemAt(index).delete();await context.sync();});}
async function upsertBudgets(changes,monthArg=currentMonth()){if(popupMode)return dialogCall('upsertBudgets',{changes,month:monthArg});await Excel.run(async context=>{const table=context.workbook.tables.getItem(TABLES.budgets),range=table.getRange();range.load('values');await context.sync();const rows=range.values.slice(1),add=[];for(const [cat,amount] of Object.entries(changes)){const i=rows.findIndex(r=>String(r[0])===monthArg&&String(r[1])===cat);const row=[monthArg,xlsSafe(cat),Number(amount)];if(i>=0)table.rows.getItemAt(i).getRange().values=[row];else if(amount>0)add.push(row);}if(add.length)table.rows.add(null,add);await context.sync();});}
async function upsertClosing(snapshot){if(popupMode)return dialogCall('upsertClosing',{snapshot});await Excel.run(async context=>{const table=context.workbook.tables.getItem(TABLES.closings),range=table.getRange();range.load('values');await context.sync();const idx=range.values.slice(1).findIndex(r=>String(r[0])===snapshot[0]);if(idx>=0)table.rows.getItemAt(idx).getRange().values=[snapshot];else table.rows.add(null,[snapshot]);await context.sync();});}

async function upsertRecurring(rule){
  if(popupMode)return dialogCall('upsertRecurring',{rule});
  await ensureRecurringTable();
  await Excel.run(async context=>{
    const table=context.workbook.tables.getItem(TABLES.recurring);
    const range=table.getRange();range.load('values');await context.sync();
    if (!hasRecurringHeaders(range.values[0])) throw new Error('Confira os cabeçalhos de tblRecorrentes.');
    const rows=range.values.slice(1);
    const idx=rows.findIndex(row=>String(row[0]||'')===rule.id);
    if(idx>=0)table.rows.getItemAt(idx).getRange().values=[ruleRow(rule)];
    else {
      const empty=rows.findIndex(row=>row.every(value => value === ''));
      if(empty>=0)table.rows.getItemAt(empty).getRange().values=[ruleRow(rule)];
      else table.rows.add(null,[ruleRow(rule)]);
    }
    await context.sync();
  });
}
async function deleteRecurring(id){
  if(popupMode)return dialogCall('deleteRecurring',{id});
  await ensureRecurringTable();
  await Excel.run(async context=>{
    const table=context.workbook.tables.getItem(TABLES.recurring);
    const range=table.getRange();range.load('values');await context.sync();
    if (!hasRecurringHeaders(range.values[0])) throw new Error('Confira os cabeçalhos de tblRecorrentes.');
    const idx=range.values.slice(1).findIndex(row=>String(row[0]||'')===id);
    if(idx<0)throw new Error('Regra não encontrada na planilha.');
    table.rows.getItemAt(idx).delete();await context.sync();
  });
}
function renderRecurring(){
  const root=$('recurringList');if(!root)return;
  const sorted=[...(db.recurring||[])].sort((a,b)=>a.kind.localeCompare(b.kind));
  root.innerHTML=sorted.length?sorted.map(r=>{
    const day=String(r.day).padStart(2,'0'),range=fmtMonth(r.start)+(r.end?' até '+fmtMonth(r.end):' em diante');
    return '<div class="recurring-item"><div class="recurring-detail"><strong>'+safe(r.description)+'</strong><small>'+safe(r.kind)+' · dia '+safe(day)+' · '+safe(range)+'</small><small>'+safe(r.account||'Conta principal')+' · '+(r.autoPaid?'Recebido na data prevista':'Pendente até confirmação')+'</small></div><div class="recurring-actions"><b>'+money(r.amount)+'</b><div><button data-recur-edit="'+safe(r.id)+'" class="secondary">Editar</button><button data-recur-del="'+safe(r.id)+'" class="secondary">Excluir</button></div></div></div>';
  }).join(''):'<p class="empty">Cadastre seu salário ou vale-alimentação para começar.</p>';
  if($('receiptInfo'))$('receiptInfo').textContent=(db.recurring||[]).length+' receita(s) fixa(s) configurada(s) · uma entrada mensal por regra.';
}
function resetRecurringForm(){
  $('recurringForm').reset();$('recurringId').value='';$('recurringStart').value=currentMonth();
  $('recurringKind').value='Salário';$('recurringDay').value='5';
}
function editRecurring(id){
  const r=(db.recurring||[]).find(x=>x.id===id);if(!r)return;
  go('recurring');$('recurringId').value=r.id;
  $('recurringKind').value=r.kind;
  $('recurringDescription').value=r.description;
  $('recurringAmount').value=r.amount;
  $('recurringDay').value=r.day;
  $('recurringAccount').value=r.account;
  $('recurringStart').value=r.start;
  $('recurringEnd').value=r.end||'';
  $('recurringAutoPaid').checked=r.autoPaid;
  $('recurringForm').scrollIntoView({behavior:'smooth',block:'start'});
}
function analysisExpenses(){
  return currentItems().filter(x=>x.type==='Despesa'&&(!analysisCategory||x.category===analysisCategory)&&(!analysisDay||Number(x.date.slice(8,10))===analysisDay));
}
function renderAnalysis(){
  if(!$('analysisCategories'))return;
  const expenses=currentItems().filter(x=>x.type==='Despesa');
  const categoryTotals=Object.entries(categorySpend(expenses)).sort((a,b)=>b[1]-a[1]);
  const maxCategory=Math.max(1,...categoryTotals.map(v=>v[1]));
  $('analysisCategories').innerHTML=categoryTotals.length?categoryTotals.map(([name,value])=>
    '<button class="a-bar'+(analysisCategory===name?' selected':'')+'" data-a-category="'+safe(name)+'" title="Filtrar por '+safe(name)+'"><span>'+safe(name)+'</span><span class="a-track"><span style="width:'+((value/maxCategory)*100).toFixed(1)+'%"></span></span><b>'+money(value)+'</b></button>'
  ).join(''):'<p class="empty">Cadastre despesas neste mês para explorar os gráficos.</p>';
  const dayBase=expenses.filter(x=>!analysisCategory||x.category===analysisCategory);
  const dayTotals={};for(const x of dayBase){const day=Number(x.date.slice(8,10));dayTotals[day]=(dayTotals[day]||0)+x.amount;}
  const [yr,mo]=currentMonth().split('-').map(Number);
  const days=yr&&mo?new Date(Date.UTC(yr,mo,0)).getUTCDate():31;
  const maxDay=Math.max(1,...Object.values(dayTotals));
  $('analysisDays').innerHTML=Array.from({length:days},(_,i)=>i+1).map(day=>{
    const val=dayTotals[day]||0;
    return '<button class="a-day'+(analysisDay===day?' selected':'')+'" data-a-day="'+day+'" title="Dia '+day+': '+money(val)+'"><span class="a-daybar" style="height:'+Math.max(3,val/maxDay*100)+'%"></span><small>'+day+'</small></button>';
  }).join('');
  const selected=analysisExpenses();
  const total=selected.reduce((sum,x)=>sum+x.amount,0);
  $('analysisTotal').textContent=money(total);
  $('analysisCount').textContent=selected.length+' lançamento(s)';
  $('analysisFilter').textContent=fmtMonth(currentMonth())+' · '+(analysisCategory||'Todas as categorias')+(analysisDay?' · dia '+analysisDay:' · mês inteiro');
  const peak=Object.entries(dayTotals).sort((a,b)=>b[1]-a[1])[0];
  $('analysisInsight').textContent=!expenses.length?'Nenhuma despesa cadastrada nesta competência.':
    peak?(analysisCategory||'Todas as categorias')+': o maior gasto ocorreu no dia '+peak[0]+', somando '+money(peak[1])+'. Clique em outro dia para comparar.':'Sem gastos para os filtros selecionados.';
  $('analysisRows').innerHTML=selected.length?[...selected].sort((a,b)=>b.date.localeCompare(a.date)).map(x=>
    '<tr><td>'+fmtDate(x.date)+'</td><td>'+safe(x.description)+'</td><td>'+safe(x.category)+'</td><td>'+safe(x.account)+'</td><td>'+money(x.amount)+'</td></tr>'
  ).join(''):'<tr><td colspan="5" class="empty">Nenhum gasto neste recorte.</td></tr>';
}

async function performSave(fn,ok){if(!connected){notify('Abra o painel pelo Excel antes de salvar dados.');return false;}try{await fn();await refresh();notify(ok);return true;}catch(err){console.error(err);notify('Não foi possível salvar: '+(err.message||err));return false;}}
function bind(){bindPlanning();if($('openWindow')){if(popupMode)$('openWindow').style.display='none';else $('openWindow').onclick=openSeparateWindow;}document.querySelectorAll('input[name=type]').forEach(el=>el.addEventListener('change',populateCategories));
$('configureRecurring').onclick=()=>go('recurring');
$('recurringCancel').onclick=resetRecurringForm;
$('recurringForm').onsubmit=async e=>{e.preventDefault();const rule={id:$('recurringId').value||('R-'+Date.now().toString(36)+'-'+Math.random().toString(36).slice(2,8)),kind:$('recurringKind').value,description:$('recurringDescription').value.trim(),amount:Number($('recurringAmount').value),day:Number($('recurringDay').value),account:$('recurringAccount').value.trim(),start:$('recurringStart').value,end:$('recurringEnd').value,autoPaid:$('recurringAutoPaid').checked,active:true};
  if(!rule.description||!(rule.amount>0)||rule.day<1||rule.day>31||!/^\d{4}-\d{2}$/.test(rule.start)||(rule.end&&rule.end<rule.start))return notify('Confira a descrição, valor, dia e período.');
  const saved=await performSave(()=>upsertRecurring(rule),'Receita fixa salva no Excel.');if(saved)resetRecurringForm();};
$('recurringList').onclick=async e=>{const edit=e.target.closest('[data-recur-edit]'),del=e.target.closest('[data-recur-del]');if(edit)editRecurring(edit.dataset.recurEdit);if(del&&confirm('Excluir esta regra? Lançamentos anteriores permanecerão no Excel.')){await performSave(()=>deleteRecurring(del.dataset.recurDel),'Regra excluída.');}};
$('analysisCategories').onclick=e=>{const b=e.target.closest('[data-a-category]');if(!b)return;analysisCategory=analysisCategory===b.dataset.aCategory?'':b.dataset.aCategory;analysisDay=0;renderAnalysis();};
$('analysisDays').onclick=e=>{const b=e.target.closest('[data-a-day]');if(!b)return;const day=Number(b.dataset.aDay);analysisDay=analysisDay===day?0:day;renderAnalysis();};
$('clearAnalysis').onclick=()=>{analysisCategory='';analysisDay=0;renderAnalysis();};
$('categoryLegend').onclick=e=>{const b=e.target.closest('[data-category-link]');if(!b)return;analysisCategory=b.dataset.categoryLink;analysisDay=0;go('analysis');};
$('barChart').onclick=e=>{const b=e.target.closest('[data-month-link]');if(!b)return;$('month').value=b.dataset.monthLink;analysisCategory='';analysisDay=0;refresh();go('analysis');};
$('viewAll').onclick=()=>go('transactions');$('viewBudgets').onclick=()=>go('budgets');$('newBtn').onclick=()=>modal();$('closeModal').onclick=closeModal;$('cancel').onclick=closeModal;$('modal').onclick=e=>{if(e.target.id==='modal')closeModal();};$('month').addEventListener('change',refresh);$('search').oninput=render;$('typeFilter').onchange=render;$('reloadBtn').onclick=refresh;
$('entryForm').onsubmit=async e=>{e.preventDefault();const type=document.querySelector('input[name=type]:checked').value;const item={id:$('entryId').value||newId(),type,description:$('description').value.trim(),amount:Number($('amount').value),date:$('date').value,category:$('category').value,account:$('account').value,notes:$('notes').value,status:$('status').value};if(!item.date||!item.description||!item.category||!(item.amount>0))return notify('Preencha a data, categoria, descrição e valor positivo.');if(fixedPayment && (item.type!=='Despesa'||item.date.slice(0,7)!==fixedPayment.month))return notify('Mantenha a conta como despesa na competência selecionada.');
const submit=e.currentTarget.querySelector('button[type=submit]');
submit.disabled=true;
let saved;
try{saved=await performSave(()=>upsertEntry(item,Boolean(fixedPayment)),'Lançamento salvo no Excel.');}
finally{submit.disabled=false;}if(saved){closeModal();$('month').value=item.date.slice(0,7);await refresh();}};
$('records').onclick=async e=>{const edit=e.target.closest('[data-edit]'),del=e.target.closest('[data-delete]');if(edit)modal(edit.dataset.edit);if(del&&confirm('Excluir definitivamente este lançamento da tabela Excel?'))await performSave(()=>deleteEntry(del.dataset.delete),'Lançamento excluído do Excel.');};
$('saveBudget').onclick=async()=>{let changes={};for(const input of document.querySelectorAll('[data-budget]')){let v=Number(input.value);if(!Number.isFinite(v)||v<0)return notify('Limites precisam ser valores positivos ou zero.');changes[input.dataset.budget]=v;}await performSave(()=>upsertBudgets(changes),'Orçamentos gravados na tabela Excel.');};
$('closeMonth').onclick=async()=>{const items=currentItems(),inc=sum(items,'Receita'),exp=sum(items,'Despesa'),pending=items.filter(x=>x.type==='Despesa'&&x.status==='Pendente').reduce((a,x)=>a+x.amount,0),month=currentMonth(),last=db.closings.find(x=>x.month===month);if(last&&!confirm(`Já existe fechamento de ${fmtMonth(month)}. Substituir o registro anterior pelos valores atuais?`))return;if(pending>0&&!confirm(`Há ${money(pending)} em despesas pendentes. Deseja registrar este fechamento mesmo assim?`))return;if(!confirm(`Fechar ${fmtMonth(month)}?\nReceitas: ${money(inc)}\nDespesas: ${money(exp)}\nSaldo: ${money(inc-exp)}`))return;const snapshot=[month,inc,exp,inc-exp,pending,new Date().toLocaleString('pt-BR'),'Fechamento efetuado pelo Controle Financeiro HTML'];await performSave(()=>upsertClosing(snapshot),'Fechamento mensal registrado no Excel.');};
window.addEventListener('focus',()=>{if(connected && Date.now()-lastRead>3500)refresh();});
}
async function init(){
  $('buildId').textContent = 'Versão ' + BUILD_ID;
  console.info('[Controle Financeiro]', {build:BUILD_ID,runtime:popupMode?'janela':'host'});
  bind();$('month').value=todayISO().slice(0,7);renderPlanning();$('recurringStart').value=currentMonth();if(typeof Office==='undefined'){notify('Office.js não carregado. Abra o suplemento dentro do Excel, com conexão à internet.');return;}try{await Office.onReady();
  if(popupMode){
    await new Promise((resolve,reject)=>{
      Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived,replyFromHost,result=>{
        if(result.status===Office.AsyncResultStatus.Succeeded)resolve();
        else reject(new Error(result.error?.message||'Não foi possível conectar ao Excel.'));
      });
    });
    connected=true;await refresh();return;
  }
  if(!Office.context||Office.context.host!==Office.HostType.Excel){notify('Prévia HTML: abra a interface como suplemento dentro do Excel para usar suas tabelas.');return;}connected=true;await refresh();}catch(err){console.error(err);notify('Falha ao iniciar suplemento: '+(err.message||err));}}
// Comando nativo na faixa de opções do Excel. Usa runtime compartilhado de longa duração.
function abrirControleFinanceiroPeloExcel(event) {
  // Só encerra o comando do Excel depois que o Office confirmar a abertura.
  Promise.resolve(Office.onReady())
    .then(() => openSeparateWindow())
    .catch(error => {
      console.error('Falha no botão da faixa de opções do Controle Financeiro:', error);
      // Se a janela não puder abrir, apresenta o diagnóstico no painel lateral.
      try { if(typeof Office.addin?.showAsTaskpane==='function') Office.addin.showAsTaskpane(); } catch(e) { console.error(e); }
    })
    .finally(() => { try { event.completed(); } catch (error) { console.error(error); } });
}
if (typeof Office !== 'undefined' && Office.actions && typeof Office.actions.associate === 'function') {
  Office.actions.associate('abrirControleFinanceiro', abrirControleFinanceiroPeloExcel);
}

init();
