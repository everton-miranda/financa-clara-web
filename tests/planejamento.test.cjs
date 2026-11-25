const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const test = require('node:test');

const source = fs.readFileSync(path.join(__dirname, '..', 'planning.js'), 'utf8');

function planningFixture() {
  const rows = [['ID','Tipo','Descricao','Categoria','Conta','Valor','Dia','Data']];
  const table = {
    isNullObject: false,
    getRange: () => ({ values: rows, load() {} }),
    rows: { add: (_, input) => rows.push(...input) }
  };
  const context = {workbook: {tables: {getItemOrNullObject: () => table}}, sync: async () => {}};
  const sandbox = {
    Excel: {run: callback => callback(context)}, popupMode: false,
    xlsSafe: value => value || '',
    isoDate: value => typeof value === 'string' ? value : '',
    excelDate: date => date,
    todayISO: () => '2026-11-24'
  };
  vm.createContext(sandbox);
  vm.runInContext(source, sandbox);
  return { sandbox, rows };
}

test('o simulador usa percentuais e centavos da reserva', () => {
  const {sandbox} = planningFixture();
  const result = sandbox.homeEstimate(9000000, 500000, 30);
  assert.equal(result.entry, 150000);
  assert.equal(result.missing, 60000);
  assert.equal(result.maximum, 300000);
  assert.equal(sandbox.homeEstimate(0, 500000, 101), null);
});

test('aportes e rendimentos entram na reserva, e resgate excessivo é rejeitado', async () => {
  const {sandbox, rows} = planningFixture();
  const movement = {description:'Teste', date:'2026-11-24'};
  await sandbox.savePlanning({...movement, id:'a', type:'Aporte', amount:100});
  await sandbox.savePlanning({...movement, id:'b', type:'Rendimento', amount:10});
  await sandbox.savePlanning({...movement, id:'c', type:'Resgate', amount:20});
  await assert.rejects(sandbox.savePlanning({...movement, id:'d', type:'Resgate', amount:91}), /reserva negativa/);
  const result = await sandbox.readPlanning();
  assert.equal(vm.runInContext('reserveCents', sandbox)(result.movements), 9000);
  assert.equal(rows.length, 4);
});

test('duas tentativas de registrar o mesmo movimento gravam apenas uma linha', async () => {
  const {sandbox, rows} = planningFixture();
  const movement = {id:'a', description:'Teste', date:'2026-11-24', amount:50, type:'Aporte'};
  await Promise.all([sandbox.savePlanning(movement),sandbox.savePlanning(movement)]);
  assert.equal(rows.length, 2);
});
