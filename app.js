/* Controle Financeiro v4 — Office.js. Nenhum dado financeiro é enviado para servidores. */
'use strict';
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
      case 'upsertEntry':
        if(String(p.item?.id||'').startsWith('__r_'))db=await readAll(p.month);
        await upsertEntry(p.item);break;
      case 'deleteEntry':
        if(String(p.id||'').startsWith('__r_'))db=await readAll(p.month);
        await deleteEntry(p.id);break;
      case 'upsertBudgets':await upsertBudgets(p.changes,p.month);break;
      case 'upsertClosing':await upsertClosing(p.snapshot);break;
      default:throw new Error('Operação não permitida.');
    }
  }catch(e){error=e.message||String(e);}
  if(controleDialog){try{controleDialog.messageChild(JSON.stringify({kind:'cf-rpc-result',rid:msg.rid,result,error}));}catch(e){console.error(e);}}
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
      Office.context.ui.displayDialogAsync(location.origin+location.pathname+'?popup=1',{height:90,width:90},result=>{
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
let connected=false, db={items:[],categories:[],budgets:{},closings:[]}, screen='overview', lastRead=0;
const TABLES={items:'tblLancamentos',categories:'tblCategorias',budgets:'tblOrcamentos',closings:'tblFechamentos'};
const currentMonth=()=>String($('month').value||'');
const currentItems=()=>db.items.filter(x=>monthKey(x)===currentMonth());
function notify(msg){$('toast').textContent=msg;$('toast').classList.add('show');clearTimeout(notify.timeout);notify.timeout=setTimeout(()=>$('toast').classList.remove('show'),5000);}
function rowsFromTable(table){const values=table.values||[],texts=table.text||[];return values.slice(1).map((row,i)=>({row,text:texts[i+1]||[]}));}
async function readAll(monthArg=currentMonth()){
  if(popupMode)return dialogCall('readAll',{month:monthArg});
  if(!connected)throw Error('Abra este painel como suplemento dentro do Excel, na pasta de trabalho Controle Financeiro.xlsx.');
  return Excel.run(async context=>{
    const ranges={};
    for(const [key,name] of Object.entries(TABLES)){
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
    return {items,categories,budgets,closings};
  });
}
async function refresh(){try{db=await readAll(); lastRead=Date.now();populateCategories();render();}catch(err){console.error(err);notify('Erro ao ler Excel: '+(err.message||err));}}
function go(to){screen=to;document.querySelectorAll('.screen').forEach(x=>x.classList.toggle('on',x.id===to));document.querySelectorAll('.nav').forEach(x=>x.classList.toggle('active',x.dataset.screen===to));const pages={overview:['Visão geral','Acompanhe seu dinheiro com clareza.'],transactions:['Lançamentos','Cadastre e edite sem sair do Excel.'],budgets:['Orçamento mensal','Planeje limites para cada categoria.'],closing:['Fechamento mensal','Revise e registre seu resultado.']};$('pageTitle').textContent=pages[to][0];$('pageSub').textContent=pages[to][1];render();}
function render(){const items=currentItems(),income=sum(items,'Receita'),expense=sum(items,'Despesa'),balance=income-expense,pending=items.filter(x=>x.type==='Despesa'&&x.status==='Pendente').reduce((a,x)=>a+x.amount,0);
  $('income').textContent=money(income);$('expenses').textContent=money(expense);$('balance').textContent=money(balance);$('savings').textContent=income?`${Math.round(balance/income*100)}%`:'—';$('balance').style.color=balance<0?'#ffd1be':'white';
  $('closingIncome').textContent=money(income);$('closingExpense').textContent=money(expense);$('closingBalance').textContent=money(balance);$('closingPending').textContent=money(pending);
  renderBars();renderDonut(items,expense);renderRecent(items);renderBudgets(items);renderTransactions(items);renderClosings(income,expense,pending);
}
function renderBars(){const [y,m]=currentMonth().split('-').map(Number);if(!y||!m)return;const months=[];for(let i=5;i>=0;i--){const d=new Date(y,m-1-i,1),key=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`,items=db.items.filter(x=>monthKey(x)===key);months.push({label:d.toLocaleDateString('pt-BR',{month:'short'}).replace('.',''),inc:sum(items,'Receita'),exp:sum(items,'Despesa')});}const max=Math.max(1,...months.flatMap(x=>[x.inc,x.exp]));$('barChart').innerHTML=months.map(x=>`<div class="barcol"><div class="bars"><div class="bar inc" title="Receitas: ${money(x.inc)}" style="height:${Math.max(1,x.inc/max*100)}%"></div><div class="bar exp" title="Despesas: ${money(x.exp)}" style="height:${Math.max(1,x.exp/max*100)}%"></div></div><span>${safe(x.label)}</span></div>`).join('');}
function renderDonut(items,expense){const data=Object.entries(categorySpend(items)).sort((a,b)=>b[1]-a[1]);$('donutTotal').textContent='R$ '+compact(expense);if(!data.length){$('donut').style.background='#e9efeb';$('categoryLegend').innerHTML='<p class="empty">Ainda não há despesas para este mês.</p>';return;}let offset=0;const stops=data.map(([cat,value],i)=>{const prev=offset;offset+=value/expense*100;return `${colors[i%colors.length]} ${prev}% ${offset}%`;});$('donut').style.background=`conic-gradient(${stops.join(',')})`;$('categoryLegend').innerHTML=data.slice(0,6).map(([cat,v],i)=>`<div class="catitem"><span><i class="dot" style="background:${colors[i%colors.length]}"></i>${safe(cat)}</span><b>${Math.round(v/expense*100)}%</b></div>`).join('');}
function renderRecent(items){let rows=[...items].sort((a,b)=>b.date.localeCompare(a.date)).slice(0,5);$('recent').innerHTML=rows.length?rows.map(a=>`<div class="txrow"><div class="txicon ${a.type==='Receita'?'credit':''}">${a.type==='Receita'?'↗':'↘'}</div><div class="txmain"><strong>${safe(a.description)}</strong><small>${safe(a.category)} · ${fmtDate(a.date)} · ${safe(a.status)}</small></div><div class="txamount ${a.type==='Receita'?'credit':'debit'}">${a.type==='Receita'?'+':'−'} ${money(a.amount)}</div></div>`).join(''):'<p class="empty">Registre o primeiro lançamento para começar.</p>';}
function renderBudgets(items){const spending=categorySpend(items),cats=db.categories.filter(c=>c.type==='Despesa').map(c=>c.name),budgetEntries=cats.filter(cat=>(db.budgets[cat]||0)>0).map(cat=>[cat,db.budgets[cat]]).sort((a,b)=>(spending[b[0]]||0)/b[1]-(spending[a[0]]||0)/a[1]);$('budgetPreview').innerHTML=budgetEntries.length?budgetEntries.slice(0,5).map(([cat,lim])=>{const val=spending[cat]||0,pct=100*val/lim;return `<div class="budgetline"><div class="budgettop"><span>${safe(cat)}</span><b>${money(val)} / ${money(lim)}</b></div><div class="budgettrack"><div class="budgetfill ${pct>100?'over':''}" style="width:${Math.min(100,pct)}%"></div></div></div>`;}).join(''):'<p class="empty">Adicione limites em Orçamento.</p>';
  $('budgetEditor').innerHTML=cats.map(cat=>`<div class="budgetfield"><label>${safe(cat)}<small>Gasto no mês: ${money(spending[cat]||0)}</small></label><input type="number" min="0" step="0.01" data-budget="${safe(cat)}" value="${Number(db.budgets[cat])||0}" aria-label="Limite para ${safe(cat)}"></div>`).join('');
}
function renderTransactions(items){const search=$('search').value.toLowerCase(),type=$('typeFilter').value;const rows=items.filter(x=>(!type||x.type===type)&&`${x.description} ${x.category} ${x.account}`.toLowerCase().includes(search)).sort((a,b)=>b.date.localeCompare(a.date));$('records').innerHTML=rows.length?rows.map(a=>`<tr><td>${fmtDate(a.date)}</td><td><b>${safe(a.description)}</b></td><td>${safe(a.category)}</td><td>${safe(a.account)}</td><td class="txamount ${a.type==='Receita'?'credit':'debit'}">${a.type==='Receita'?'+':'−'} ${money(a.amount)}</td><td>${safe(a.status)}</td><td><button data-edit="${safe(a.id)}">Editar</button><button class="delete" data-delete="${safe(a.id)}">Excluir</button></td></tr>`).join(''):'<tr><td colspan="7" class="empty">Nenhum lançamento para esta competência.</td></tr>';}
function renderClosings(income,expense,pending){const m=currentMonth(),last=db.closings.find(x=>x.month===m),diff=last&&(Math.abs(last.income-income)>0.01||Math.abs(last.expense-expense)>0.01||Math.abs(last.pending-pending)>0.01);const status=$('closingStatus');status.classList.toggle('warn',!!diff || pending>0);if(!last){status.textContent=pending>0?`Este mês ainda não foi fechado. Há ${money(pending)} em despesas pendentes.`:'Este mês ainda não foi fechado. Os totais estão prontos para conferência.';}else if(diff){status.textContent=`Atenção: ${fmtMonth(m)} já foi fechado em ${last.closedAt}, mas os lançamentos mudaram depois. Confira e registre um novo fechamento se necessário.`;}else{status.textContent=`Competência ${fmtMonth(m)} fechada em ${last.closedAt}. Os totais conferem com o Excel.`;}const rows=[...db.closings].sort((a,b)=>b.month.localeCompare(a.month)).slice(0,14);$('closingHistory').innerHTML=rows.length?rows.map(x=>`<div class="historyrow"><div><strong>${fmtMonth(x.month)}</strong><small>${safe(x.closedAt)}</small></div><div>Receitas <b>${money(x.income)}</b></div><div>Despesas <b>${money(x.expense)}</b></div><div>Saldo <b>${money(x.balance)}</b></div><div>Pendente <b>${money(x.pending)}</b></div></div>`).join(''):'<p class="empty">Nenhum mês foi fechado ainda.</p>';}
function populateCategories(){const type=document.querySelector('input[name=type]:checked').value,prior=$('category').value,cats=db.categories.filter(x=>x.type===type).map(x=>x.name);$('category').innerHTML=cats.map(c=>`<option value="${safe(c)}">${safe(c)}</option>`).join('');if(cats.includes(prior))$('category').value=prior;}
function modal(id){$('entryForm').reset();$('entryId').value='';$('date').value=todayISO();$('modalTitle').textContent='Novo lançamento';if(id){const row=db.items.find(x=>x.id===id);if(!row)return notify('Lançamento não encontrado. Atualize o painel.');$('entryId').value=row.id;$('modalTitle').textContent='Editar lançamento';document.querySelector(`input[name=type][value="${row.type}"]`).checked=true;populateCategories();for(const k of ['description','amount','date','category','account','notes'])$(k).value=row[k]??'';$('status').value=row.status||'Pago';}else{populateCategories();}$('modal').classList.add('show');$('description').focus();}
function closeModal(){$('modal').classList.remove('show');}
function entryRow(x){return [xlsSafe(x.id),excelDate(x.date),x.type,xlsSafe(x.description),xlsSafe(x.category),xlsSafe(x.account),Number(x.amount),x.status,xlsSafe(x.notes)];}
function newId(){return 'FC-'+Date.now().toString(36)+'-'+Math.random().toString(36).slice(2,7);}
async function upsertEntry(item){if(popupMode)return dialogCall('upsertEntry',{item,month:currentMonth()});await Excel.run(async context=>{const table=context.workbook.tables.getItem(TABLES.items),range=table.getRange();range.load('values');await context.sync();const rows=range.values.slice(1);let index=rows.findIndex(row=>String(row[0])===item.id);if(index<0 && item.id.startsWith('__r_')){const source=db.items.find(x=>x.id===item.id),i=Number(item.id.slice(4));if(source && i>=0 && i<rows.length && !String(rows[i][0]??'').trim() && isoDate(rows[i][1])===source.date && String(rows[i][3]??'').trim()===source.description){index=i;item.id=newId();}else throw Error('A tabela mudou desde a última atualização. Clique em Atualizar antes de editar.');}if(index>=0)table.rows.getItemAt(index).getRange().values=[entryRow(item)];else table.rows.add(null,[entryRow(item)]);await context.sync();});}
async function deleteEntry(id){if(popupMode)return dialogCall('deleteEntry',{id,month:currentMonth()});await Excel.run(async context=>{const table=context.workbook.tables.getItem(TABLES.items),range=table.getRange();range.load('values');await context.sync();let index=range.values.slice(1).findIndex(r=>String(r[0])===id);if(index<0 && id.startsWith('__r_')){const source=db.items.find(x=>x.id===id),i=Number(id.slice(4)),rows=range.values.slice(1);if(source&&i>=0&&i<rows.length&&!String(rows[i][0]??'').trim()&&isoDate(rows[i][1])===source.date&&String(rows[i][3]??'').trim()===source.description)index=i;}if(index<0)throw Error('Lançamento não encontrado ou planilha alterada. Atualize o painel.');table.rows.getItemAt(index).delete();await context.sync();});}
async function upsertBudgets(changes,monthArg=currentMonth()){if(popupMode)return dialogCall('upsertBudgets',{changes,month:monthArg});await Excel.run(async context=>{const table=context.workbook.tables.getItem(TABLES.budgets),range=table.getRange();range.load('values');await context.sync();const rows=range.values.slice(1),add=[];for(const [cat,amount] of Object.entries(changes)){const i=rows.findIndex(r=>String(r[0])===monthArg&&String(r[1])===cat);const row=[monthArg,xlsSafe(cat),Number(amount)];if(i>=0)table.rows.getItemAt(i).getRange().values=[row];else if(amount>0)add.push(row);}if(add.length)table.rows.add(null,add);await context.sync();});}
async function upsertClosing(snapshot){if(popupMode)return dialogCall('upsertClosing',{snapshot});await Excel.run(async context=>{const table=context.workbook.tables.getItem(TABLES.closings),range=table.getRange();range.load('values');await context.sync();const idx=range.values.slice(1).findIndex(r=>String(r[0])===snapshot[0]);if(idx>=0)table.rows.getItemAt(idx).getRange().values=[snapshot];else table.rows.add(null,[snapshot]);await context.sync();});}
async function performSave(fn,ok){if(!connected){notify('Abra o painel pelo Excel antes de salvar dados.');return false;}try{await fn();await refresh();notify(ok);return true;}catch(err){console.error(err);notify('Não foi possível salvar: '+(err.message||err));return false;}}
function bind(){if($('openWindow')){if(popupMode)$('openWindow').style.display='none';else $('openWindow').onclick=openSeparateWindow;}document.querySelectorAll('.nav').forEach(b=>b.addEventListener('click',()=>go(b.dataset.screen)));document.querySelectorAll('input[name=type]').forEach(el=>el.addEventListener('change',populateCategories));
$('viewAll').onclick=()=>go('transactions');$('viewBudgets').onclick=()=>go('budgets');$('newBtn').onclick=()=>modal();$('closeModal').onclick=closeModal;$('cancel').onclick=closeModal;$('modal').onclick=e=>{if(e.target.id==='modal')closeModal();};$('month').addEventListener('change',refresh);$('search').oninput=render;$('typeFilter').onchange=render;$('reloadBtn').onclick=refresh;$('backupBtn').onclick=refresh;
$('entryForm').onsubmit=async e=>{e.preventDefault();const type=document.querySelector('input[name=type]:checked').value;const item={id:$('entryId').value||newId(),type,description:$('description').value.trim(),amount:Number($('amount').value),date:$('date').value,category:$('category').value,account:$('account').value,notes:$('notes').value,status:$('status').value};if(!item.date||!item.description||!item.category||!(item.amount>0))return notify('Preencha a data, categoria, descrição e valor positivo.');const saved=await performSave(()=>upsertEntry(item),'Lançamento salvo no Excel.');if(saved){closeModal();$('month').value=item.date.slice(0,7);await refresh();}};
$('records').onclick=async e=>{const edit=e.target.closest('[data-edit]'),del=e.target.closest('[data-delete]');if(edit)modal(edit.dataset.edit);if(del&&confirm('Excluir definitivamente este lançamento da tabela Excel?'))await performSave(()=>deleteEntry(del.dataset.delete),'Lançamento excluído do Excel.');};
$('saveBudget').onclick=async()=>{let changes={};for(const input of document.querySelectorAll('[data-budget]')){let v=Number(input.value);if(!Number.isFinite(v)||v<0)return notify('Limites precisam ser valores positivos ou zero.');changes[input.dataset.budget]=v;}await performSave(()=>upsertBudgets(changes),'Orçamentos gravados na tabela Excel.');};
$('closeMonth').onclick=async()=>{const items=currentItems(),inc=sum(items,'Receita'),exp=sum(items,'Despesa'),pending=items.filter(x=>x.type==='Despesa'&&x.status==='Pendente').reduce((a,x)=>a+x.amount,0),month=currentMonth(),last=db.closings.find(x=>x.month===month);if(last&&!confirm(`Já existe fechamento de ${fmtMonth(month)}. Substituir o registro anterior pelos valores atuais?`))return;if(pending>0&&!confirm(`Há ${money(pending)} em despesas pendentes. Deseja registrar este fechamento mesmo assim?`))return;if(!confirm(`Fechar ${fmtMonth(month)}?\nReceitas: ${money(inc)}\nDespesas: ${money(exp)}\nSaldo: ${money(inc-exp)}`))return;const snapshot=[month,inc,exp,inc-exp,pending,new Date().toLocaleString('pt-BR'),'Fechamento efetuado pelo Controle Financeiro HTML'];await performSave(()=>upsertClosing(snapshot),'Fechamento mensal registrado no Excel.');};
window.addEventListener('focus',()=>{if(connected && Date.now()-lastRead>3500)refresh();});
}
async function init(){bind();$('month').value=todayISO().slice(0,7);if(typeof Office==='undefined'){notify('Office.js não carregado. Abra o suplemento dentro do Excel, com conexão à internet.');return;}try{await Office.onReady();
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
