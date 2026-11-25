const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const test = require('node:test');

const root = path.join(__dirname, '..');
const source = fs.readFileSync(path.join(root, 'app.js'), 'utf8').replace(/\binit\(\);\s*$/, '');
const headers = ['ID','Tipo','Descricao','Valor','Dia','Conta','Inicio','Fim','AutoPago','Ativo'];
const rule = ['salario','Salário','Salário fictício',1000,31,'Conta principal','2026-01','','Não','Sim'];
const va = ['va','Vale-alimentação','Benefício fictício',300,30,'Cartão VA','2026-01','','Não','Sim'];

function fixture({ cells, tableName, failAdd, closed } = {}) {
  const logs = [];
  const entries = [[], ['manual','2026-02-10','Despesa','Teste','Casa','Conta principal',10,'Pago','']];
  const tables = {
    tblLancamentos: {
      getRange: () => ({ values: entries, load() {} }),
      rows: { add: (_, rows) => entries.push(...rows) }
    },
    ...Object.fromEntries(Object.entries({
      tblCategorias: [[], ['Casa','Despesa']],
      tblOrcamentos: [[], ['2026-02','Casa',50]],
      tblFechamentos: closed ? [[], ['2026-02']] : [[]]
    }).map(([name, values]) => [name, { getRange: () => ({ values, load() {} }) }]))
  };
  const sheet = {
    isNullObject: false,
    tables: {
      items: [],
      load() {},
      add() {
        if (failAdd) throw new TypeError('Cannot set property fill of #<Object> which has only a getter');
        const table = {
          isNullObject: false,
          getRange: () => ({ values: cells, rowIndex: 0, columnIndex: 0, columnCount: 10, load() {} }),
          set name(name) { tables[name] = this; }
        };
        this.items.push(table);
        return table;
      }
    },
    getRange(address) {
      if (address !== 'A1:J1') return {};
      return {
        get values() { return [cells[0] || Array(10).fill('')]; },
        set values(rows) { cells[0] = rows[0]; },
        get formulas() { return this.values; },
        load() {}
      };
    },
    getUsedRangeOrNullObject: () => ({
      isNullObject: !cells.length, rowIndex: 0, columnIndex: 0, columnCount: 10, rowCount: cells.length, load() {}
    })
  };
  if (tableName) sheet.tables.add().name = tableName;
  const context = {
    workbook: {
      tables: {
        getItemOrNullObject: name => tables[name] || {isNullObject: true},
        getItem: name => tables[name]
      },
      worksheets: {
        getItemOrNullObject: () => cells ? sheet : {isNullObject: true},
        add: () => { cells = []; return sheet; }
      }
    },
    sync: async () => {}
  };
  const sandbox = {
    URLSearchParams, location: {search: ''}, Excel: {run: callback => callback(context)},
    console: {error: (_, log) => logs.push(log)}
  };
  vm.createContext(sandbox);
  vm.runInContext(source + '\nconnected = true;', sandbox);
  return {read: () => sandbox.readAll('2026-02'), entries, tables, sheet, logs};
}

test('o identificador do JS coincide com os ativos do HTML', () => {
  const build = source.match(/const BUILD_ID = '([^']+)'/)[1];
  const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
  assert.ok(html.includes('app.js?v=' + build));
  assert.ok(html.includes('style.css?v=' + build));
});

test('cria a tabela uma vez quando a aba está ausente', async () => {
  const app = fixture();
  await Promise.all([app.read(), app.read()]);
  assert.ok(app.tables.tblRecorrentes);
  assert.equal(app.sheet.tables.items.length, 1);
});

for (const tableName of [undefined, 'Tabela1', 'tblRecorrentes']) {
  test('preserva as regras em esquema inicial, parcial ou completo: ' + tableName, async () => {
    const cells = [headers, rule, va];
    const app = fixture({cells, tableName});
    const [first, second] = await Promise.all([app.read(), app.read()]);
    assert.equal(first.recurringWarning, undefined);
    assert.equal(first.items[1].date, '2026-02-28');
    assert.equal(second.items[2].account, 'Cartão VA');
    assert.deepEqual(cells, [headers, rule, va]);
    assert.equal(app.entries.length, 4);
    app.entries[2][6] = 700;
    assert.equal((await app.read()).items[1].amount, 700);
  });
}

test('um erro de migração deixa as tabelas principais disponíveis', async () => {
  const app = fixture({failAdd: true});
  const result = await app.read();
  assert.equal(result.items[0].id, 'manual');
  assert.equal(result.budgets.Casa, 50);
  assert.ok(result.recurringWarning);
  assert.equal(app.logs[0].stage, 'tables.add / sync');
  assert.ok(app.logs[0].stack);
});

test('estrutura desconhecida é preservada e o painel continua disponível', async () => {
  const cells = [['Dados existentes']];
  const app = fixture({cells});
  assert.ok((await app.read()).recurringWarning);
  assert.deepEqual(cells, [['Dados existentes']]);
});

test('mês fechado mantém as regras e não recebe lançamentos gerados', async () => {
  const app = fixture({cells:[headers, rule], closed: true});
  assert.equal((await app.read()).recurring.length, 1);
  assert.equal(app.entries.length, 2);
});
