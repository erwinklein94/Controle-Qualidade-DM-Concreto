# Painel de séries de dormentes — Cavan + Rumo

Site estático pronto para GitHub Pages, com visual inspirado na identidade da Rumo.

## Como publicar

1. Crie um repositório no GitHub.
2. Envie estes arquivos para a raiz do repositório:
   - `index.html`
   - `Painel de series de DM.xlsx`
3. No GitHub, acesse **Settings > Pages**.
4. Em **Build and deployment**, selecione **Deploy from a branch**.
5. Escolha a branch `main` e a pasta `/root`.
6. Salve e abra a URL gerada pelo GitHub Pages.

## Como atualizar os dados

- Substitua o arquivo `Painel de series de DM.xlsx` no repositório por uma planilha atualizada com as mesmas colunas.
- O site tenta ler essa planilha automaticamente.
- Se a planilha não carregar, ele usa os dados embutidos no `index.html` como fallback.
- Também existe um botão **Atualizar com planilha** para carregar um arquivo manualmente no navegador.

## Abas do painel

### Séries de liberação

Aba principal para controlar as séries de produção, ensaio e liberação para carregamento.

### Dash board

Aba para acompanhar a **produção semanal** junto com os **principais indicadores de refugo/reprova**.

- Os filtros de **Data inicial** e **Data final** vêm pré-preenchidos com a última semana disponível na planilha carregada.
- O filtro **Projeto / bitola** permite ver a produção semanal e os refugos de um projeto específico ou de todos.
- A tabela **Produção por projeto no período** mostra a quantidade produzida por cada combinação de projeto + bitola.
- O gráfico **Produção semanal - toda a planilha** soma todos os projetos e mostra a evolução por semana.
- Os cards de qualidade mostram **refugos/reprovas**, **taxa de refugo**, **lotes afetados**, **motivo líder** e **lote crítico** no mesmo recorte do dash board.
- As tabelas **Motivos que mais puxam reprova** e **Lotes críticos para marcar e atacar** ajudam a priorizar ações corretivas.

## Filtros importantes na aba Séries

- Use **Data inicial** para ignorar lotes antigos que já tiveram ensaio de liberação.
- Use **Data final** quando quiser analisar apenas um intervalo específico.
- O filtro **Projeto / bitola** separa automaticamente projetos de mesmo nome em combinações independentes, como `MALHA PAULISTA • BL` e `MALHA PAULISTA • BM`.
- O filtro **Série** permite localizar uma série específica dentro do projeto/bitola selecionado ou em todos os projetos.
- O filtro **Bitola** permite ver somente BL, BM ou itens sem bitola identificada.

## Regra usada no modo “Recalcular pela regra”

- Cada série é fechada quando atinge **2.000 dormentes produzidos** ou **10 lotes produzidos**.
- A contagem é independente por **projeto + bitola**.
- O lote que fecha a meta é destacado como **lote do ensaio**.
- Após o ensaio aprovado, clique em **Marcar ensaio feito e liberar** para indicar que a série está liberada para carregamento em carretas.
- Essa marcação fica salva no navegador do usuário.


## Atualização incluída

- Nova aba **Refugos / Reprovas** no `index.html`.
- Leitura do bloco `REPROVADOS - CAVAN SANTA LUCIA` da planilha.
- Filtros por projeto/bitola, período, motivo comum e lote.
- Indicadores, ranking de motivos, visão por lote e tabela detalhada das ocorrências.
- Dash board atualizado com indicadores de refugo/reprova no mesmo filtro de produção: taxa, motivo líder, lotes afetados e lotes críticos.
- A planilha atualizada deve permanecer no repositório como `Painel de series de DM.xlsx`.


## Atualização do dashboard

A aba Dash board compara produção, refugos/reprovas e taxa de refugo dentro do período e projeto selecionados. A planilha padrão do repositório foi atualizada para `Painel de series de DM.xlsx`.
