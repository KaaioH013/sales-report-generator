# --------------------------------------------------------------------------
# FERRAMENTA: PREPARADOR DE RELATÓRIOS (v1.2 - Preserva Detalhes dos Itens)
# --------------------------------------------------------------------------
# -*- coding: utf-8 -*-

import pandas as pd
import os
from datetime import datetime

# --- CONFIGURAÇÕES ---
MAPEAMENTO_COLUNAS = {
    "relatorio_530": {
        "cliente": "Cliente", "data_pedido": "Dt.Pedido", "status_pedido": "St.",
        "valor_total": "Vlr.Total", "quantidade": "Qtde", "codigo_material": "Cod.Material",
        "descricao_produto": "Descrição", "valor_unitario": "Vlr. Unitário",
        "numero_pedido": "Nro.Pedido", "data_entrega": "Dt.Entrega", "status_item": "St.Item"
    },
    "relatorio_549": {
        "numero_pedido_pv": "PV", "vendedor": "VENDEDOR EXTERNO", "uf": "UF",
        "nome_cliente": "NOME CLIENTE"
    }
}

def preparar_dados(caminho_vendas_530, caminho_faturamento_549):
    """
    Função principal que lê, consolida, limpa e salva os dados, preservando os detalhes dos itens.
    """
    print("Iniciando o processo de preparação de dados...")

    try:
        print(f"Lendo ficheiro de vendas: {os.path.basename(caminho_vendas_530)}")
        df_vendas = pd.read_excel(caminho_vendas_530)
        
        print(f"Lendo ficheiro de faturação: {os.path.basename(caminho_faturamento_549)}")
        df_faturamento = pd.read_excel(caminho_faturamento_549)
        print("Ficheiros lidos com sucesso.\n")

    except Exception as e:
        print(f"\nERRO ao ler os ficheiros Excel. Detalhes: {e}")
        return

    print("Iniciando a transformação e limpeza dos dados...")
    
    col_530 = MAPEAMENTO_COLUNAS['relatorio_530']
    col_549 = MAPEAMENTO_COLUNAS['relatorio_549']

    df_vendas[col_530['numero_pedido']] = df_vendas[col_530['numero_pedido']].astype(str)
    df_faturamento[col_549['numero_pedido_pv']] = df_faturamento[col_549['numero_pedido_pv']].astype(str)

    colunas_info = [col_549['numero_pedido_pv'], col_549['vendedor'], col_549['uf'], col_549['nome_cliente']]
    info_vendedores_unicos = df_faturamento[colunas_info].drop_duplicates(subset=[col_549['numero_pedido_pv']])

    # --- MUDANÇA ESTRUTURAL ---
    # Em vez de agrupar os totais, juntamos as informações a cada linha de item do relatório 530.
    df_final = pd.merge(
        df_vendas, 
        info_vendedores_unicos, 
        left_on=col_530['numero_pedido'], 
        right_on=col_549['numero_pedido_pv'], 
        how='left'
    )
    
    df_final.rename(columns={
        col_549['vendedor']: 'Vendedor',
        col_549['uf']: 'UF',
        col_549['nome_cliente']: 'Nome Cliente Faturamento' # Renomeado para evitar conflito com 'Cliente' do 530
    }, inplace=True)
    
    df_final['Vendedor'] = df_final['Vendedor'].fillna('Vendedor não encontrado')
    df_final['UF'] = df_final['UF'].fillna('N/D')
    
    print("Dados consolidados e limpos.\n")

    print("Verificando a qualidade dos dados...")
    
    # Para o relatório de qualidade, primeiro identificamos os pedidos únicos com problemas
    pedidos_com_problemas_df = df_final[
        (df_final['Vendedor'] == 'Vendedor não encontrado') |
        (df_final['UF'] == 'N/D')
    ][[col_549['numero_pedido_pv'], 'Vendedor', 'UF']].drop_duplicates()
    
    periodo_str = "geral"
    try:
        primeira_data = pd.to_datetime(df_final[col_530['data_pedido']].dropna().iloc[0])
        periodo_str = primeira_data.strftime('%m_%Y')
    except Exception:
        print("Aviso: Não foi possível determinar o período a partir das datas.")

    caminho_saida_qualidade = f'relatorio_qualidade_dados_{periodo_str}.txt'
    
    with open(caminho_saida_qualidade, 'w', encoding='utf-8') as f:
        f.write(f"Relatório de Qualidade de Dados - Período: {periodo_str}\n")
        f.write(f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        f.write("="*50 + "\n\n")

        if pedidos_com_problemas_df.empty:
            f.write("Ótima notícia! Nenhum problema de dados encontrado.\n")
            print("Ótima notícia! Nenhum problema de qualidade de dados foi encontrado.")
        else:
            print(f"AVISO: Foram encontrados {len(pedidos_com_problemas_df)} pedidos com dados em falta.")
            f.write(f"AVISO: Foram encontrados {len(pedidos_com_problemas_df)} pedidos com dados em falta.\n\n")
            f.write(pedidos_com_problemas_df.to_string(index=False))
            print(f"Um relatório detalhado dos problemas foi salvo em: {caminho_saida_qualidade}")

    print("\n")
    caminho_saida_excel = f'dados_consolidados_{periodo_str}.xlsx'
    
    try:
        df_final.to_excel(caminho_saida_excel, index=False)
        print(f"Sucesso! Ficheiro de dados limpos e consolidados foi salvo como:\n{caminho_saida_excel}")
    except Exception as e:
        print(f"\nERRO ao salvar o ficheiro Excel final. Detalhes: {e}")

    print("\nProcesso concluído!")

if __name__ == '__main__':
    print("--- Ferramenta de Preparação de Relatórios de Vendas ---")
    print("Este programa irá juntar os relatórios 530 e 549, limpar os dados e salvá-los num único ficheiro Excel.\n")

    def limpar_caminho(caminho_raw):
        caminho_limpo = caminho_raw.strip()
        if caminho_limpo.startswith('& '):
            caminho_limpo = caminho_limpo[2:]
        caminho_limpo = caminho_limpo.strip("'\"")
        return caminho_limpo

    caminho_raw_530 = input("Por favor, arraste o ficheiro do relatório 530 para esta janela e pressione Enter:\n> ")
    caminho_530 = limpar_caminho(caminho_raw_530)

    caminho_raw_549 = input("\nÓtimo. Agora, arraste o ficheiro do relatório 549 para esta janela e pressione Enter:\n> ")
    caminho_549 = limpar_caminho(caminho_raw_549)
    
    if not os.path.exists(caminho_530) or not os.path.exists(caminho_549):
        print("\nERRO: Um ou ambos os caminhos de ficheiro não são válidos. Por favor, tente novamente.")
        print(f"Caminho 530 interpretado: '{caminho_530}'")
        print(f"Caminho 549 interpretado: '{caminho_549}'")
    else:
        print("\nObrigado. A processar os seus ficheiros...")
        print("-" * 50)
        preparar_dados(caminho_530, caminho_549)
        print("-" * 50)
    
    input("\nPressione Enter para sair.")