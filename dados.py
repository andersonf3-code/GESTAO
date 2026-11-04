"""Leitura fiel das abas do modelo de gestão CAF, sem alterar o Excel."""
from datetime import date, datetime
from io import BytesIO
import warnings
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

CABECALHOS = {'Demandas': 1, 'Listas': 1, 'Cooperações': 5, 'Resumo': 3}
ERROS = {'#VALUE!', '#REF!', '#DIV/0!', '#NAME?', '#N/A', '#NUM!', '#NULL!', '#SPILL!', '#CALC!'}

def ler_planilha(conteudo: bytes):
    with warnings.catch_warnings():
        warnings.simplefilter('ignore', UserWarning)
        wb = load_workbook(BytesIO(conteudo), data_only=True)
    abas = {}
    for ws in wb:
        linhas = list(ws.values)
        cab = CABECALHOS.get(ws.title, 4)
        if ws.title == 'Demandas' and ws.max_row < 4:
            abas[ws.title] = pd.DataFrame()
            continue
        if len(linhas) < cab:
            abas[ws.title] = pd.DataFrame()
            continue
        n = max((i + 1 for r in linhas[cab-1:] for i, v in enumerate(r) if v is not None), default=0)
        vistos = {}
        colunas = []
        for i in range(n):
            raw = linhas[cab-1][i]
            nome = str(raw).strip() if raw is not None else f'Coluna {get_column_letter(i+1)} (sem título)'
            vistos[nome] = vistos.get(nome, 0) + 1
            colunas.append(nome if vistos[nome] == 1 else f'{nome} ({vistos[nome]})')
        registros = []
        for numero, linha in enumerate(linhas[cab:], start=cab+1):
            if not any(v is not None for v in linha[:n]):
                continue
            valores = [v.strftime('%d/%m/%Y') if isinstance(v, (date, datetime)) else v for v in linha[:n]]
            registros.append({'Linha Excel': numero, **dict(zip(colunas, valores))})
        abas[ws.title] = pd.DataFrame(registros, columns=['Linha Excel', *colunas])
    wb.close()
    return abas

def registros(nome, df):
    if nome == 'Cooperações' and 'Nome' in df:
        return df.loc[df['Nome'].fillna('').astype(str).str.strip().ne('ACT SIGLA')].copy()
    return df.copy()

def exibir(df):
    return df.map(lambda v: '' if pd.isna(v) else str(v))

def filtrar(df, busca='', escolhas=None):
    saida = df.copy()
    for coluna, valores in (escolhas or {}).items():
        if valores:
            saida = saida.loc[saida[coluna].fillna('Não informado').astype(str).isin(valores)]
    if busca.strip():
        texto = exibir(saida)
        mascara = texto.apply(lambda c: c.str.contains(busca.strip(), case=False, regex=False)).any(axis=1)
        saida = saida.loc[mascara]
    return saida

def contar(df, coluna):
    if coluna not in df:
        return pd.DataFrame(columns=['Categoria', 'Quantidade'])
    return df[coluna].fillna('Não informado').replace('', 'Não informado').astype(str).value_counts().rename_axis('Categoria').reset_index(name='Quantidade')
