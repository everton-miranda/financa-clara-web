'use strict';

const PLANNING_TABLE = 'tblPlanejamento';
const PLANNING_HEADERS = ['ID', 'Tipo', 'Descricao', 'Categoria', 'Conta', 'Valor', 'Dia', 'Data'];
const emptyPlanning = { bills: [], movements: [] };
let planning = emptyPlanning;
let planningReady = false;
let pendingPlanningWrite = Promise.resolve();

async function withPlanningTable(create, operation) {
  return Excel.run(async context => {
    const workbook = context.workbook;
    let table = workbook.tables.getItemOrNullObject(PLANNING_TABLE);
    await context.sync();
    if (table.isNullObject) {
      if (!create) return emptyPlanning;
      let sheet = workbook.worksheets.getItemOrNullObject('Planejamento');
      await context.sync();
      if (sheet.isNullObject) {
        sheet = workbook.worksheets.add('Planejamento');
        await context.sync();
      }
      const used = sheet.getUsedRangeOrNullObject();
      const header = sheet.getRange('A1:H1');
      header.load('values');
      sheet.tables.load('items');
      await context.sync();
      const matches = row => PLANNING_HEADERS.every((value, index) => row[index] === value);
      if (sheet.tables.items.length === 1 && matches(header.values[0])) {
        table = sheet.tables.items[0];
        const range = table.getRange();
        range.load(['rowIndex', 'columnIndex', 'columnCount']);
        await context.sync();
        if (range.rowIndex || range.columnIndex || range.columnCount !== PLANNING_HEADERS.length) {
          throw Error('Confira a estrutura da aba Planejamento.');
        }
      } else {
        if (!used.isNullObject || sheet.tables.items.length) {
          throw Error('A aba Planejamento já contém dados. Confira sua estrutura no Excel.');
        }
        header.values = [PLANNING_HEADERS];
        table = sheet.tables.add(sheet.getRange('A1:H2'), true);
      }
      table.name = PLANNING_TABLE;
      await context.sync();
    }
    const range = table.getRange();
    range.load('values');
    await context.sync();
    if (!PLANNING_HEADERS.every((value, index) => range.values[0][index] === value)) {
      throw Error('Confira os cabeçalhos de tblPlanejamento no Excel.');
    }
    return operation(context, table, range.values.slice(1));
  });
}

function parsePlanning(rows) {
  const bills = [];
  const movements = [];
  for (const [id, type, description, category, account, value, day, rawDate] of rows) {
    const amount = Number(value);
    if (!id || !Number.isFinite(amount) || amount <= 0) continue;
    if (type === 'Conta fixa') bills.push({ id: String(id), description, category, account, amount, day: Number(day) });
    if (['Aporte', 'Resgate', 'Rendimento'].includes(type)) {
      const date = isoDate(rawDate);
      if (date) movements.push({ id: String(id), type, description, amount, date });
    }
  }
  return { bills, movements };
}

function readPlanning() {
  return withPlanningTable(false, (_context, _table, rows) => parsePlanning(rows));
}

const inCents = amount => Math.round(amount * 100);
const signedMovement = movement => inCents(movement.amount) * (movement.type === 'Resgate' ? -1 : 1);
const reserveCents = movements => movements.reduce((total, movement) => total + signedMovement(movement), 0);

function homeEstimate(balance, price, percent) {
  if (!(price > 0 && percent > 0 && percent <= 100)) return null;
  const entry = Math.round(inCents(price) * percent / 100);
  return {
    entry: entry / 100,
    missing: Math.max(0, entry - balance) / 100,
    maximum: Math.floor(Math.max(0, balance) * 100 / percent) / 100
  };
}

async function savePlanning(item) {
  if (popupMode) return dialogCall('savePlanning', { item });
  const isBill = item.type === 'Conta fixa';
  if (!item.id || !item.description || !Number.isFinite(item.amount) || item.amount <= 0 ||
      (isBill ? !item.category || !Number.isInteger(item.day) || item.day < 1 || item.day > 31 :
        !['Aporte', 'Resgate', 'Rendimento'].includes(item.type) || !isoDate(item.date) || item.date > todayISO())) {
    throw Error('Confira os dados e informe um valor positivo.');
  }
  const write = pendingPlanningWrite.then(() => withPlanningTable(true, async (context, table, rows) => {
    const index = rows.findIndex(row => String(row[0]) === item.id);
    if (!isBill && index >= 0) return;
    if (!isBill && item.type === 'Resgate') {
      const timeline = [...parsePlanning(rows).movements, item].sort((a, b) =>
        a.date.localeCompare(b.date) || Number(a.type === 'Resgate') - Number(b.type === 'Resgate'));
      let balance = 0;
      if (timeline.some(movement => (balance += signedMovement(movement)) < 0)) {
        throw Error('O resgate deixaria a reserva negativa na data informada ou em uma data posterior.');
      }
    }
    const row = [xlsSafe(item.id), item.type, xlsSafe(item.description), xlsSafe(item.category),
      xlsSafe(item.account), item.amount, isBill ? item.day : '', isBill ? '' : excelDate(item.date)];
    if (index >= 0) table.rows.getItemAt(index).getRange().values = [row];
    else table.rows.add(null, [row]);
    await context.sync();
  }));
  pendingPlanningWrite = write.catch(() => {});
  await write;
}

function fixedEntryId(bill, month) {
  return 'FIX-' + bill.id + '-' + month;
}

function fixedEntry(bill, month) {
  return db.items.find(entry => entry.id === fixedEntryId(bill, month));
}

function renderEstimate() {
  const estimate = homeEstimate(reserveCents(planning.movements), Number($('homePrice').value), Number($('homePercent').value));
  for (const [id, value] of Object.entries({
    homeEntry: estimate?.entry, homeMissing: estimate?.missing, homeMaximum: estimate?.maximum
  })) $(id).textContent = value === undefined ? '—' : money(value);
}

function renderPlanning() {
  const month = currentMonth();
  $('planningStatus').hidden = planningReady;
  $('fixedCategory').innerHTML = db.categories.filter(category => category.type === 'Despesa')
    .map(category => '<option>' + safe(category.name) + '</option>').join('');
  $('billList').innerHTML = planning.bills.map(bill => {
    const entry = fixedEntry(bill, month);
    const budget = db.budgets[bill.category] || 0;
    return '<div class="bill-row"><div><strong>' + safe(bill.description) + '</strong><small>' +
      safe(bill.category) + ' · dia ' + bill.day + ' · ' + money(entry?.amount ?? bill.amount) +
      (entry ? ' · ' + safe(entry.status) : '') + (budget ? ' · orçamento ' + money(budget) : '') +
      '</small></div><div class="bill-actions"><button class="secondary" data-bill-edit="' + safe(bill.id) +
      '">Editar custo</button><button class="primary" data-bill-pay="' + safe(bill.id) + '">' +
      (entry?.status === 'Pago' ? 'Ver pagamento' : 'Lançar conta') + '</button></div></div>';
  }).join('') || '<p class="empty">Cadastre seu primeiro custo fixo.</p>';
  $('reserveBalance').textContent = money(reserveCents(planning.movements) / 100);
  $('reserveRows').innerHTML = [...planning.movements].sort((a, b) => b.date.localeCompare(a.date)).map(movement =>
    '<div class="bill-row"><div><strong>' + safe(movement.description) + '</strong><small>' +
    fmtDate(movement.date) + ' · ' + movement.type + '</small></div><b>' +
    money(signedMovement(movement) / 100) + '</b></div>').join('') || '<p class="empty">Nenhuma movimentação na reserva.</p>';
  const paid = (db.items || []).filter(entry => monthKey(entry) === month && entry.status === 'Pago');
  const cash = paid.reduce((sum, entry) => {
    if (entry.account === 'Cartão VA' || (entry.type === 'Receita' && entry.category === 'Vale-alimentação')) return sum;
    return sum + inCents(entry.amount) * (entry.type === 'Despesa' ? -1 : 1);
  }, 0);
  const transfers = planning.movements.filter(movement => movement.date.startsWith(month) && movement.type !== 'Rendimento');
  $('spendableCash').textContent = planningReady ? money((cash - reserveCents(transfers)) / 100) : '—';
  renderEstimate();
}

let fixedPayment = null;

function openFixedPayment(bill) {
  const month = currentMonth();
  if (!month) return notify('Selecione a competência.');
  const entry = fixedEntry(bill, month);
  if (entry) return modal(entry.id);
  modal();
  fixedPayment = { id: fixedEntryId(bill, month), month };
  $('entryId').value = fixedPayment.id;
  $('modalTitle').textContent = 'Lançar custo fixo';
  $('description').value = bill.description;
  $('category').value = bill.category;
  if (![...$('account').options].some(option => option.value === bill.account)) $('account').add(new Option(bill.account));
  $('account').value = bill.account;
  $('date').value = recurringDate(month, bill.day);
  $('amount').value = bill.amount;
  $('status').value = 'Pendente';
  const budget = db.budgets[bill.category] || 0;
  $('paymentBudget').hidden = !budget;
  $('useBudgetAmount').textContent = 'Usar orçamento · ' + money(budget);
  $('useBudgetAmount').onclick = () => { $('amount').value = budget.toFixed(2); };
}

function bindPlanning() {
  $('reserveDate').value = todayISO();
  $('homePrice').oninput = $('homePercent').oninput = renderEstimate;
  $('fixedForm').onsubmit = async event => {
    event.preventDefault();
    const form = event.currentTarget;
    const item = { id: $('fixedId').value || newId(), description: $('fixedLabel').value.trim(),
      type: 'Conta fixa', category: $('fixedCategory').value, account: $('fixedAccount').value.trim(),
      amount: Number($('fixedAmount').value), day: Number($('fixedDay').value) };
    const submit = form.querySelector('button[type=submit]');
    submit.disabled = true;
    try { if (await performSave(() => savePlanning(item), 'Custo fixo salvo.')) form.reset(); }
    finally { submit.disabled = false; }
  };
  $('billList').onclick = event => {
    const button = event.target.closest('button');
    if (!button) return;
    const bill = planning.bills.find(item => item.id === (button.dataset.billEdit || button.dataset.billPay));
    if (!bill) return;
    if (button.dataset.billPay) return openFixedPayment(bill);
    for (const [id, value] of Object.entries({
      fixedId: bill.id, fixedLabel: bill.description, fixedCategory: bill.category, fixedAccount: bill.account,
      fixedAmount: bill.amount, fixedDay: bill.day
    })) $(id).value = value;
    $('fixedLabel').focus();
  };
  $('reserveForm').onsubmit = async event => {
    event.preventDefault();
    const form = event.currentTarget;
    const type = $('reserveType').value;
    const item = { id: newId(), type, description: $('reserveLabel').value.trim() || type,
      amount: Number($('reserveAmount').value), date: $('reserveDate').value };
    const submit = form.querySelector('button[type=submit]');
    submit.disabled = true;
    try {
      if (await performSave(() => savePlanning(item), 'Reserva atualizada.')) {
        form.reset();
        $('reserveDate').value = todayISO();
      }
    } finally { submit.disabled = false; }
  };
  $('quickIncome').onclick = () => {
    modal();
    document.querySelector('input[name=type][value=Receita]').checked = true;
    populateCategories();
  };
  document.querySelectorAll('[data-quick]').forEach(button => {
    button.onclick = () => go(button.dataset.quick);
  });
}
