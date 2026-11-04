# BI Gestão CAF em Python

Abra esta pasta no VSCode. Recomendado Python 3.11 ou superior.

No terminal, instale as dependências:

```powershell
python -m pip install -r requirements.txt
```

Execute:

```powershell
python -m streamlit run app.py --server.address 127.0.0.1 --server.port 8502
```

Abra http://127.0.0.1:8502. Caso a porta esteja bloqueada, troque 8502 por 9000 no comando e no endereço. Se o Windows não reconhecer python, substitua por py. Não precisa de Node.js ou Git.

Mantenha o terminal aberto; Ctrl+C encerra o painel. Se a instalação for bloqueada pela rede institucional, solicite à TI a disponibilização das dependências.

## Arquivos

- app.py: aplicação Streamlit em Python.
- dados.py: leitura e filtros dos dados.
- planilha_caf.xlsx: cópia da planilha fornecida.
- requirements.txt: dependências.

## Atualização

Substitua planilha_caf.xlsx na mesma pasta e atualize a página, ou carregue um Excel atualizado pelo menu lateral. O upload prevalece sobre a planilha local durante a sessão. Para voltar à planilha local, remova o upload.

As nove abas ficam disponíveis na navegação. Cada aba inclui todos os campos preenchidos da origem, busca, filtros pertinentes, consulta e CSV. Listas e Resumo são referências, não demandas. A aba Demandas original está vazia. A linha de instruções de Cooperações está disponível mediante opção específica e é excluída dos indicadores por padrão. Possíveis duplicatas não são removidas.

Fórmulas não são recalculadas por openpyxl: o painel utiliza seus resultados salvos. Recalcule e salve no Excel antes de carregar se necessário. Erros de origem são mostrados e contados. O painel não altera nem grava a planilha.

Esta versão Python é independente da versão HTML publicada anteriormente. Não há atualização entre as duas versões.

## Correção visual

Inclui tema claro em .streamlit/config.toml. Extraia todo o pacote, incluindo a pasta .streamlit, e abra a pasta do projeto no VSCode antes de executar. Pare a instância antiga com Ctrl+C e reinicie o comando. A interface também respeita as cores do tema se você escolher tema escuro nas configurações do Streamlit.
