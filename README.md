# Painel de séries de dormentes — Cavan

Site estático pronto para GitHub Pages.

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

## Regra usada no modo “Recalcular pela regra”

- Cada série é fechada quando atinge **2.000 dormentes produzidos** ou **10 lotes produzidos**.
- O lote que fecha a meta é destacado como **lote do ensaio**.
- Após o ensaio aprovado, clique em **Marcar ensaio feito e liberar** para indicar que a série está liberada para carregamento em carretas.
- Essa marcação fica salva no navegador do usuário.
