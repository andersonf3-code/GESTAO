"""BI Gestão CAF — execute com python -m streamlit run app.py."""
from datetime import datetime
from pathlib import Path
import pandas as pd
import streamlit as st
from dados import ler_planilha, registros, exibir, filtrar, contar, ERROS

st.set_page_config(page_title='BI Gestão CAF', page_icon='🌱', layout='wide')
st.markdown('''<style>
/* Use as cores do tema ativo: evita texto branco em superfícies claras. */
.block-container{padding-top:2.2rem;padding-bottom:3rem;max-width:1600px}
[data-testid="stSidebar"]{border-right:1px solid var(--secondary-background-color)}
[data-testid="stMetric"]{
 background:var(--secondary-background-color);
 color:var(--text-color);
 border:1px solid color-mix(in srgb,var(--text-color) 18%,transparent);
 border-radius:12px;padding:20px;min-height:125px;
}
[data-testid="stMetricLabel"], [data-testid="stMetricValue"]{color:var(--text-color)}
[data-testid="stMetricValue"]{font-weight:700}
[data-testid="stCaptionContainer"]{color:var(--text-color);opacity:.82}
[data-testid="stSidebar"] [data-testid="stRadio"] label{padding:4px 0}
@media(max-width:700px){.block-container{padding:1.2rem}}
</style>''', unsafe_allow_html=True)

@st.cache_data(show_spinner=False)
def carregar(conteudo):
    return ler_planilha(conteudo)

st.sidebar.title('🌱 CAF · Gestão')
st.sidebar.caption('Compartilhamento de dados')
upload = st.sidebar.file_uploader('Carregar planilha atualizada', type=['xlsx'])
base = Path(__file__).parent / 'planilha_caf.xlsx'
if upload is not None:
    conteudo, nome_arquivo = upload.getvalue(), upload.name
elif base.exists():
    conteudo, nome_arquivo = base.read_bytes(), base.name
else:
    st.info('Carregue a planilha Excel no menu lateral para começar.')
    st.stop()
try:
    abas = carregar(conteudo)
except Exception as erro:
    st.error(f'Não foi possível ler o Excel: {erro}')
    st.stop()
nav = st.sidebar.radio('Navegação', ['Visão geral', *abas.keys()])
st.sidebar.divider()
st.sidebar.caption(f'Fonte: {nome_arquivo}')
st.sidebar.caption('O painel lê o Excel a cada atualização da página. Para atualizar um upload, envie a nova versão.')
st.sidebar.download_button('Baixar Excel de origem', conteudo, file_name=nome_arquivo, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
st.caption('MDA / Agricultura Familiar • BI de gestão')

def grafico(df, coluna, titulo):
    st.subheader(titulo)
    resumo = contar(df, coluna)
    if resumo.empty:
        st.info('Sem registros para este recorte.')
    else:
        st.bar_chart(resumo, x='Categoria', y='Quantidade', horizontal=True, color='#427c5b', width='stretch', height=max(250, min(650, len(resumo)*38)))
        with st.expander('Ver quantidades'):
            st.dataframe(resumo, hide_index=True, width='stretch')

def obter(nome):
    return registros(nome, abas.get(nome, pd.DataFrame()))

def metricas(df):
    campos = df.drop(columns='Linha Excel', errors='ignore')
    erros = int(campos.map(lambda v: str(v) in ERROS).sum().sum())
    cols = st.columns(4)
    cols[0].metric('Registros no recorte', len(df))
    cols[1].metric('Campos da origem', len(campos.columns))
    cols[2].metric('Células com erro', erros)
    cols[3].metric('Linhas com campos vazios', int(campos.isna().any(axis=1).sum()))

if nav == 'Visão geral':
    st.title('Visão geral')
    st.write('Acompanhe as demandas de dados, as solicitações de API e as cooperações institucionais.')
    ident, conecta, coop = obter('Identificáveis'), obter('ConectaGov'), obter('Cooperações')
    cols = st.columns(4)
    cols[0].metric('Demandas identificáveis', len(ident))
    cols[1].metric('Atendidas', int(ident.get('Status', pd.Series(dtype=str)).eq('Atendido').sum()))
    cols[2].metric('Solicitações ConectaGov', len(conecta))
    cols[3].metric('Registros de cooperação', len(coop))
    a,b = st.columns(2)
    with a:
        grafico(ident, 'Status', 'Demandas por status')
        grafico(ident, 'Tipo de requerente', 'Perfil dos requerentes')
    with b:
        grafico(conecta, 'Situação na plataforma', 'ConectaGov por situação')
        grafico(ident, 'Situação do prazo', 'Situação do prazo informada')
    st.subheader('Conteúdo das abas')
    inventario = pd.DataFrame([{'Aba': n, 'Registros': len(registros(n,d)), 'Campos': len(d.drop(columns='Linha Excel',errors='ignore').columns)} for n,d in abas.items()])
    st.dataframe(inventario, hide_index=True, width='stretch')
    st.info('Os totais de cada aba representam unidades diferentes e não devem ser somados. Listas é um dicionário; Resumo contém indicadores da origem. Cooperações exclui apenas a linha de instruções da contagem, sem deduplicar acordos.')
    st.warning('Erros de fórmula e campos vazios são preservados. Datas, dias em aberto e situação dos prazos reproduzem os valores salvos no Excel; o painel não recalcula fórmulas. Caso necessário, recalcule e salve a planilha no Excel antes de carregá-la.')
else:
    st.title(nav)
    original = abas[nav]
    df = registros(nav, original)
    if nav == 'Listas':
        st.info('Opções de preenchimento da planilha. Estas linhas não representam demandas.')
    if nav == 'Resumo':
        st.info('Indicadores salvos na aba Resumo, mantidos conforme a origem.')
    if nav == 'Cooperações':
        st.info('A linha de instruções ACT SIGLA não entra nos indicadores. Os demais registros são mantidos, inclusive possíveis repetições. Textos de vigência não são convertidos automaticamente em datas.')
        if st.checkbox('Mostrar também a linha de instruções'):
            df = original.copy()
    if df.empty:
        st.info('Esta aba ainda não possui registros preenchidos.')
        if len(original.columns):
            st.write('Campos previstos:', ', '.join(c for c in original.columns if c != 'Linha Excel'))
    else:
        st.subheader('Filtros')
        busca = st.text_input('Buscar em todos os campos', placeholder='Processo SEI, instituição, finalidade…')
        candidatas = ['Status','Situação na plataforma','Tipo de requerente','API','Forma de tratamento','Classificação dos dados','Possui metadados?','Contém dados sensíveis','Situação do prazo','Base legal','Situação geral']
        escolhas = {}
        with st.expander('Filtrar por campos', expanded=True):
            colunas = st.columns(3)
            for i,c in enumerate(c for c in candidatas if c in df):
                with colunas[i%3]:
                    opcoes = sorted(df[c].fillna('Não informado').astype(str).unique())
                    escolhas[c] = st.multiselect(c, opcoes, key=f'{nav}_{c}')
        filtrado = filtrar(df, busca, escolhas)
        metricas(filtrado)
        st.caption(f'{len(filtrado)} de {len(df)} registros. Os filtros atualizam indicadores, gráficos, detalhes e exportação.')
        coluna_grafico = next((c for c in candidatas if c in filtrado), None)
        if coluna_grafico:
            grafico(filtrado, coluna_grafico, f'Distribuição por {coluna_grafico}')
        st.subheader('Dados completos')
        st.dataframe(exibir(filtrado), hide_index=True, width='stretch', height=450)
        st.download_button('Exportar recorte em CSV', exibir(filtrado).to_csv(index=False,sep=';').encode('utf-8-sig'), file_name=f'{nav}_filtrado.csv', mime='text/csv')
        if not filtrado.empty:
            st.subheader('Detalhamento sem cortes')
            linhas = filtrado['Linha Excel'].tolist()
            def rotulo(linha):
                r = filtrado.loc[filtrado['Linha Excel'].eq(linha)].iloc[0]
                campos = ['Processo SEI','Processo SEI relacionado','Nome','Análises','Órgão solicitante']
                titulo = next((str(r[c]) for c in campos if c in r and pd.notna(r[c])), 'Registro')
                return f'Linha {linha} • {titulo}'
            selecionada = st.selectbox('Selecione um registro', linhas, format_func=rotulo)
            registro = filtrado.loc[filtrado['Linha Excel'].eq(selecionada)].iloc[0]
            for c,v in registro.items():
                st.text(c)
                st.text('Não informado' if pd.isna(v) else str(v))
