# Finança Clara — interface HTML para Excel

Site estático para o suplemento Office.js. **Não envie dados financeiros ou o arquivo `.xlsx` a este repositório público.**

- Interface: https://everton-miranda.github.io/financa-clara-web/
- Manifesto do painel incorporado: [FinancaClara-PAINEL.xml](./FinancaClara-PAINEL.xml)
- Manifesto do painel lateral: [FinancaClara-LATERAL.xml](./FinancaClara-LATERAL.xml)

## Ativar o GitHub Pages

Em **Settings → Pages**, escolha **Deploy from a branch**, selecione **main** e **/(root)** e pressione **Save**.

## Instalar no Excel

Use **Inserir → Meus Suplementos → Carregar meu suplemento** (se disponível) e selecione o manifesto do painel. Se a opção não aparecer, utilize um catálogo compartilhado confiável autorizado pelo Windows/Office. A interface só acessa a pasta de trabalho quando carregada dentro do Excel.

A pasta de trabalho deve ter as tabelas `tblLancamentos`, `tblCategorias`, `tblOrcamentos` e `tblFechamentos`. Use a planilha corrigida fornecida separadamente pelo ChatGPT.

O GitHub Pages hospeda somente os arquivos HTML, CSS, JavaScript e ícones públicos. Nenhum arquivo de dados pessoais é publicado aqui.
