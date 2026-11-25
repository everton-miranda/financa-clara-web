# Controle Financeiro — Excel + HTML (versão 6.1)

Projeto de controle financeiro para Excel Microsoft 365 com interface HTML hospedada no GitHub Pages. O código público **não contém lançamentos nem valores pessoais**.

- Aplicativo: https://everton-miranda.github.io/financa-clara-web/
- Suplemento com botão no Excel: [Controle-Financeiro-BOTAO-EXCEL.xml](./Controle-Financeiro-BOTAO-EXCEL.xml)
- As abas `Lancamentos`, `Orcamentos`, `Categorias` e `Fechamentos` continuam no arquivo Excel do usuário. **Não substitua sua planilha por uma versão em branco se já tiver dados.**

## Novidades da versão 6

- **Receitas fixas:** menu "Receitas fixas" para cadastrar salário, vale-alimentação e outras entradas, com valor, dia, conta/cartão, mês inicial/final e opção de marcar automaticamente como recebido na data. A alternativa segura é manter **Pendente** até conferir o crédito. O saldo em dinheiro é exibido separado do saldo de VA. Não há consulta bancária.
- **Excel automático:** ao usar o menu, o aplicativo cria a aba `Recorrentes` e a tabela `tblRecorrentes` no próprio arquivo, se ainda não existirem. Não exige uma planilha substituta, e permite editar essas regras diretamente no Excel. A receita de cada mês é gerada ao abrir aquela competência, com ID estável para evitar duplicação. Alterações manuais em lançamentos já criados são preservadas. **Exclusão manual de um lançamento recorrente pode fazer com que ele seja recriado ao atualizar o mês; para interromper a geração, exclua ou encerre a regra em Receitas fixas.** Salve a pasta de trabalho no Excel para persistir as alterações.
- **Análise interativa:** menu "Análise interativa", categoria → dia → lista detalhada. No gráfico de fluxo, clique em um mês. Na legenda do gráfico de categorias, clique no nome para filtrar. Os filtros estão vinculados dentro da interface. Não é Power BI, mas oferece navegação similar em análises básicas.
- **Janela HTML maior:** o botão do suplemento abre uma janela a 98% da tela. O título da janela, o endereço do site e os controles nativos são determinados pelo Office; o site não pode esconder o endereço nem inserir botões nativos de minimizar/maximizar.

## Uso

1. Continue abrindo seu arquivo Excel atual.
2. Na faixa de opções, clique em **Abrir Controle Financeiro**.
3. Na janela HTML, acesse **Receitas fixas**, cadastre suas regras e salve. A nova aba `Recorrentes` aparecerá no Excel.
4. Para registrar compras com vale, escolha **Cartão VA** como conta/meio de pagamento.
5. Entre em **Análise interativa** para escolher categoria, dia e ver os lançamentos correspondentes.
6. Salve seu Excel regularmente. Não publique sua pasta de trabalho com dados neste repositório público.

## Custos e limites

GitHub Pages não cobra para este repositório público dentro dos limites da oferta gratuita. O Excel deve estar licenciado. A interface exige internet para carregar e o Office.js se comunica com as tabelas da pasta aberta, sem enviar valores pessoais para o GitHub. O botão e os manifestos existentes permanecem com os mesmos nomes.


## Correção 6.1 — recebimentos mensais recorrentes

- A aba **Recorrentes** é criada **somente ao salvar uma regra de receita fixa**. A abertura do dashboard, mesmo em planilhas antigas, não tenta formatar ou criar tabelas.
- O aplicativo não faz mais a operação de formatação `format.fill` que originou o erro de propriedade somente leitura. Se a criação da nova tabela falhar, outros dados permanecem intactos.
- Ao visualizar a competência, gera um lançamento por regra ativa, identificada por `REC-<ID>-AAAA-MM`. Atualizações sucessivas do painel são serializadas para impedir duplicações.
- **Não gera lançamentos automáticos em meses já fechados**. O histórico de fechamentos prevalece.
- O padrão de pagamentos recorrentes é **Pendente**. A opção de considerar recebido automaticamente gera o registro como **Pago** somente na data do recebimento ou depois; ela não confirma uma entrada bancária real.
- Regra de exemplo: Salário — R$ X — dia 15 — primeiro mês 2026-11 — sem último mês.
- Os dados continuam no próprio Excel. Salve o arquivo após cadastrar/gerar lançamentos.
- A versão é identificada no rodapé do aplicativo como `V6.1`.

**Testes em simulador Office.js:** leitura sem tabela de recorrências, não geração em mês fechado, não criação antecipada no modo Pago automático, serialização de leituras concorrentes e não duplicação do mês. **A abertura no Excel Windows real ainda precisa ser validada pelo usuário.**
