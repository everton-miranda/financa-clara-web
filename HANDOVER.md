# Handover do Controle Financeiro

## 24/11/2026 — diagnóstico e candidato 6.1.0

Base de trabalho: `42e069a` em `main`. As alterações desta etapa foram reunidas no commit “Diagnosticar inicialização do Excel e proteger a migração de recorrências” da branch `manutencao/diagnostico-excel`.

- `app.js` e `index.html`: versão 6.1.0 visível na interface, revisão de URLs dos ativos, diagnóstico local com etapa, build, stack e campos Office.js, e aviso de divergência entre host e diálogo no console. A inicialização de recorrências remove setters de estilo, preserva os dados de abas parciais, recupera tabela compatível que ficou com nome temporário e recusa estruturas desconhecidas. As quatro tabelas principais são lidas em uma operação independente. Geração serializada por runtime, com IDs existentes preservados e geração bloqueada em períodos fechados.
- `tests/recorrentes.test.cjs`: oito testes com dados sintéticos e simulação de Office.js. Validaram versões, inicialização com aba ausente, recuperação parcial, tabela existente, falha isolada da migração, estrutura desconhecida, salário e VA idempotentes, limite de fevereiro, preservação de valor editado e mês fechado.
- `README.md`, `CHANGELOG.md` e este arquivo: atualização, rollback e roteiro de aceite. Os manifestos conservam configuração e identidade.

O commit base trocou a atribuição direta em `format.fill` por `format.fill.color`, mas o HTML conservou `app.js?v=6.0`. A atribuição antiga reproduziu a mesma mensagem em um objeto JavaScript com getter sem setter. O Pages retornou HTTP 200 para HTML e JS, serviu o código do hotfix com o endereço antigo e devolveu `Cache-Control: max-age=600`. Isso sustenta a hipótese de cache, ainda sem confirmação do runtime que apresentou a falha. O stack do erro original no Excel continua pendente. `node --check app.js`, os oito testes e `git diff --check` passaram.

## Aceite antes da publicação

Em uma pasta de trabalho fictícia, abra a compilação candidata no Excel pelo botão da faixa e confira **Versão 6.1.0** também no painel lateral. Confirme lançamento pré-existente e a criação única de `Recorrentes/tblRecorrentes`. Repita com aba parcial e com tabela existente. Se a migração falhar, conserve a cópia e examine no console do host o evento `[Controle Financeiro]` com build, etapa, código e stack. Se a versão estiver antiga, siga o procedimento de cache do README.

Em fevereiro de 2026, cadastre salário de R$ 1.000 no dia 31 e VA de R$ 300 no dia 30, ambos pendentes. Na pasta sintética com despesas de R$ 75 no Cartão VA e R$ 100 na Conta principal, espere duas receitas com data 28/02, saldo VA de R$ 225 e saldo em dinheiro de R$ 900. Atualize duas vezes, edite o salário no Excel, salve, feche e reabra. Confirme IDs únicos e a edição preservada. Filtre Alimentação e dia 10 na análise para obter a despesa de R$ 75, teste também mês vazio, orçamento e fechamento. Registre os resultados antes de publicar em `main`. Após publicação, confirme HTTP 200, versão dos arquivos e versão efetivamente carregada no Excel. Para reverter uma versão publicada, reverta o commit do release no Git, aguarde o Pages e recarregue o suplemento. Preserve a pasta de trabalho e o backup.

## Janela no Windows

| Opção | Comportamento |
| --- | --- |
| Office Dialog atual | Janela separada com comunicação com a pasta aberta. A API oferece largura e altura iniciais. O Office controla título, URL e botões nativos. |
| Painel lateral | Mantém integração e evita a barra do diálogo separado. Ocupa espaço dentro do Excel. |
| Companheiro Windows com WebView2 | Pode desenhar uma janela nativa. Exigiria uma ponte local autenticada e testes de sincronização com a pasta aberta, distribuição e operação sem privilégios. Depende de validação e aprovação da mudança de arquitetura. |

A janela existente e o caminho pelo painel lateral foram preservados no código. A abertura e o fechamento pelo botão, a persistência no XLSX e a navegação de gráficos exigem o aceite desktop. [Microsoft Office Dialog API](https://learn.microsoft.com/en-us/javascript/api/office/office.dialogoptions)

### Verificação do arquivo de teste

A pasta de trabalho fictícia, criada fora do repositório, abriu no Excel desktop 16.0 em modo somente leitura. O Excel enumerou as quatro tabelas e abas esperadas, com a aba Recorrentes ausente. Esse teste verifica o arquivo e seu esquema. O suplemento candidato ainda precisa executar no mesmo Excel para validar o runtime, a migração, a persistência e os gráficos.
