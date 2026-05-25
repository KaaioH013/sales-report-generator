import pandas as pd

from relatorio_servicos import _produto_label_map


def test_maior_frequencia_vence():
    df = pd.DataFrame({
        'Cod.Material': ['A', 'A', 'A'],
        'Descrição': ['DESC_A', 'DESC_A', 'OUTRA'],
    })
    assert _produto_label_map(df) == {'A': 'DESC_A'}


def test_empate_usa_ultima_descricao_na_ordem():
    df = pd.DataFrame({
        'Cod.Material': ['A', 'A', 'A', 'A'],
        'Descrição': ['UM', 'DOIS', 'UM', 'DOIS'],
    })
    assert _produto_label_map(df)['A'] == 'DOIS'


def test_multiplos_dataframes_concatena_ordem():
    df1 = pd.DataFrame({'Cod.Material': ['1'], 'Descrição': ['Pri']})
    df2 = pd.DataFrame({'Cod.Material': ['1', '1'], 'Descrição': ['Seg', 'Seg']})
    assert _produto_label_map(df1, df2)['1'] == 'Seg'


def test_groupby_cod_material_rename_com_label():
    df = pd.DataFrame({
        'Cod.Material': ['99', '99'],
        'Descrição': ['Certa', 'Errada typo'],
        'Vlr.Total': [10.0, 90.0],
    })
    lmap = _produto_label_map(df)
    agg = df.groupby('Cod.Material', sort=True)['Vlr.Total'].sum()
    labelled = agg.rename(index=lambda c: lmap.get(c, str(c)))
    assert len(labelled) == 1
    assert labelled.index[0] == 'Errada typo'
    assert float(labelled.iloc[0]) == 100.0


def test_ignora_df_sem_colunas_obrigatorias():
    df_ok = pd.DataFrame({'Cod.Material': ['x'], 'Descrição': ['L']})
    df_bad = pd.DataFrame({'Algo': [1]})
    assert _produto_label_map(df_bad, df_ok) == {'x': 'L'}

