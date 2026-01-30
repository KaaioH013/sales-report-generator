# --- INÍCIO DO CÓDIGO ---
# --------------------------------------------------------------------------
# PROGRAMA: GERADOR DE RELATÓRIOS (v19.5 - Finalização Pareto)
# --------------------------------------------------------------------------
# -*- coding: utf-8 -*-

import os
import sys
import matplotlib
matplotlib.use('Agg')
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk, messagebox
from ttkthemes import ThemedTk
import threading
import multiprocessing
import locale
import queue
from datetime import datetime
import pathlib
import pandas as pd
import numpy as np
import tempfile
import shutil

# --- LAZY LOADING ---
import matplotlib.pyplot as plt
import seaborn as sns
from weasyprint import HTML
from matplotlib.ticker import FuncFormatter, PercentFormatter

try:
    import pyi_splash
except ImportError:
    pyi_splash = None

try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil.1252')
    except locale.Error:
        print("AVISO: Locale pt_BR não encontrado.")

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- CONSTANTES GLOBAIS ---
NOME_FICHEIRO_LOGO = "Logo.png"
CORES_GRAFICOS = ['#3498db', '#e74c3c', '#2ecc71', '#f1c40f', '#9b59b6', '#34495e', '#1abc9c', '#d35400', '#c0392b', '#8e44ad']
PALETA_GRAFICOS = sns.color_palette(CORES_GRAFICOS)


def formatar_moeda_str(valor):
    return f'R$ {valor:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')

def obter_caminho_logo_absoluto():
    caminho_abs = resource_path(NOME_FICHEIRO_LOGO)
    if os.path.exists(caminho_abs):
        return pathlib.Path(caminho_abs).as_uri()
    return ""

def _asset_path(assets_dir, filename):
    return os.path.join(assets_dir, filename)

def gerar_mapa_calor(df_vendas_por_estado, log_queue, assets_dir):
    import geopandas
    import pickle
    
    log_queue.put("Iniciando geração do mapa de calor...")
    try:
        caminho_mapa_pkl = resource_path("mapa_brasil_dados.pkl")
        with open(caminho_mapa_pkl, 'rb') as f:
            dados_mapa = pickle.load(f)
        
        gdf_brasil = geopandas.GeoDataFrame(list(dados_mapa.keys()), geometry=list(dados_mapa.values()), columns=['abbrev_state'])
        
        df_vendas_mapa = df_vendas_por_estado.reset_index()
        df_vendas_mapa.rename(columns={'UF': 'abbrev_state', 'Vlr.Total': 'Valor'}, inplace=True)
        
        mapa_com_dados = gdf_brasil.merge(df_vendas_mapa, on='abbrev_state', how='left')
        mapa_com_dados['Valor'] = mapa_com_dados['Valor'].fillna(0)

        fig, ax = plt.subplots(1, 1, figsize=(10, 10))
        mapa_com_dados.plot(column='Valor', cmap='viridis', linewidth=0.8, ax=ax, edgecolor='0.8', legend=True,
                                legend_kwds={'label': "Valor Total de Vendas (R$)", 'orientation': "horizontal", 'shrink': 0.6})
        
        ax.set_axis_off()
        ax.set_title('Distribuição de Vendas por Estado', fontdict={'fontsize': '16', 'fontweight' : '3'})
        
        nome_grafico = 'grafico_mapa_calor.png'
        plt.savefig(_asset_path(assets_dir, nome_grafico), dpi=150)
        plt.close(fig)
        log_queue.put("Mapa de calor gerado com sucesso.")
        return nome_grafico
    except Exception as e:
        log_queue.put(f"ERRO ao gerar mapa de calor: {e}")
        return None

# --- FUNÇÃO 1 (MODIFICADA PARA TAREFA 1) ---
def gerar_relatorio_completo(caminho_dados_consolidados, log_queue, save_path):
    def update_status(message):
        log_queue.put(message)
    assets_dir = None
    try:
        assets_dir = tempfile.mkdtemp(prefix="relatorio_assets_")
        update_status("Iniciando o processo de geração de relatório...")
        update_status(f"Lendo ficheiro de dados consolidados: {os.path.basename(caminho_dados_consolidados)}")
        
        df_final = pd.read_excel(caminho_dados_consolidados)
        update_status("Ficheiro de dados lido com sucesso.")

        try:
            df_final['Dt.Pedido'] = pd.to_datetime(df_final['Dt.Pedido'])
            periodo_label = df_final['Dt.Pedido'].min().strftime('%B de %Y').capitalize()
        except Exception:
            periodo_label = "Período Geral"
        
        update_status(f"Período identificado: {periodo_label}")
        update_status("Iniciando análises...")
        
        analise_vendedor = df_final.groupby('Vendedor').agg(
            Valor_Total_Vendido=('Vlr.Total', 'sum'),
            Num_Pedidos=('Nro.Pedido', 'nunique')
        ).sort_values(by='Valor_Total_Vendido', ascending=False)
        analise_vendedor['Ticket_Medio'] = analise_vendedor['Valor_Total_Vendido'] / analise_vendedor['Num_Pedidos']
        
        vendas_por_estado = df_final[df_final['UF'] != 'N/D'].groupby('UF')['Vlr.Total'].sum().sort_values(ascending=False)
        vendas_diarias = df_final.groupby(df_final['Dt.Pedido'].dt.date)['Vlr.Total'].sum()
        vendas_por_produto_valor = df_final.groupby('Descrição')['Vlr.Total'].sum().sort_values(ascending=False).head(10)
        vendas_por_produto_qtd = df_final.groupby('Descrição')['Qtde'].sum().sort_values(ascending=False).head(10)
        vendas_por_cliente_todos = df_final.groupby('Cliente')['Vlr.Total'].sum().sort_values(ascending=False)
        vendas_por_cliente_top10 = vendas_por_cliente_todos.head(10)

        df_pareto = vendas_por_cliente_todos.to_frame()
        df_pareto.rename(columns={'Vlr.Total': 'Valor_Total'}, inplace=True)
        df_pareto['% Acumulado'] = df_pareto['Valor_Total'].cumsum() / df_pareto['Valor_Total'].sum() * 100
        df_pareto_grafico = df_pareto.head(30)
        update_status("Análises concluídas. Gerando gráficos...")
        
        nome_grafico_pareto_cliente = 'grafico_pareto_cliente.png'
        fig_pareto, ax_pareto = plt.subplots(figsize=(12, 6))
        ax_pareto.bar(df_pareto_grafico.index, df_pareto_grafico['Valor_Total'], color="C0")
        ax_pareto.tick_params(axis='x', rotation=90, labelsize=8)
        ax2_pareto = ax_pareto.twinx()
        ax2_pareto.plot(df_pareto_grafico.index, df_pareto_grafico["% Acumulado"], color="C1", marker="o", ms=5)
        ax2_pareto.yaxis.set_major_formatter(PercentFormatter())
        ax_pareto.set_title("Análise de Pareto por Cliente (Top 30)", fontsize=16)
        fig_pareto.tight_layout()
        plt.savefig(_asset_path(assets_dir, nome_grafico_pareto_cliente), dpi=150)
        plt.close(fig_pareto)
        update_status("Gráfico 'Análise de Pareto' criado.")
        
        nome_grafico_vendas_tempo = 'grafico_vendas_tempo.png'
        fig_tempo, ax_tempo = plt.subplots(figsize=(12, 6))
        vendas_diarias.plot(ax=ax_tempo, marker='o', linestyle='-')
        ax_tempo.set_title(f'Evolução das Vendas - {periodo_label}', fontsize=16)
        fig_tempo.tight_layout()
        plt.savefig(_asset_path(assets_dir, nome_grafico_vendas_tempo), dpi=150)
        plt.close(fig_tempo)
        update_status("Gráfico 'Vendas no Tempo' criado.")

        nome_grafico_vendedor = 'grafico_vendas_vendedor.png'
        fig_vendedor, ax_vendedor = plt.subplots(figsize=(12, 8))
        sns.barplot(y=analise_vendedor.index, x=analise_vendedor['Valor_Total_Vendido'], ax=ax_vendedor, orient='h')
        ax_vendedor.set_title('Vendas por Vendedor', fontsize=16)
        for patch in ax_vendedor.patches:
            width = patch.get_width()
            y = patch.get_y()
            height = patch.get_height()
            ax_vendedor.text(width * 1.01, y + height / 2, f'{formatar_moeda_str(width)}', va='center')
        ax_vendedor.set_xlim(right=ax_vendedor.get_xlim()[1] * 1.25)
        fig_vendedor.tight_layout()
        plt.savefig(_asset_path(assets_dir, nome_grafico_vendedor), dpi=150)
        plt.close(fig_vendedor)
        update_status("Gráfico 'Vendas por Vendedor' criado.")

        nome_grafico_clientes = 'grafico_top_clientes.png'
        fig_cliente, ax_cliente = plt.subplots(figsize=(12, 8))
        sns.barplot(y=vendas_por_cliente_top10.index, x=vendas_por_cliente_top10.values, ax=ax_cliente, orient='h')
        ax_cliente.set_title('Top 10 Clientes por Valor de Compra', fontsize=16)
        for patch in ax_cliente.patches:
            width = patch.get_width()
            y = patch.get_y()
            height = patch.get_height()
            ax_cliente.text(width * 1.01, y + height / 2, f'{formatar_moeda_str(width)}', va='center')
        ax_cliente.set_xlim(right=ax_cliente.get_xlim()[1] * 1.25)
        fig_cliente.tight_layout()
        plt.savefig(_asset_path(assets_dir, nome_grafico_clientes), dpi=150)
        plt.close(fig_cliente)
        update_status("Gráfico 'Top 10 Clientes' criado.")
        
        nome_grafico_produtos_valor = 'grafico_top_produtos_valor.png'
        fig_produto_valor, ax_prod_val = plt.subplots(figsize=(12, 8))
        vendas_por_produto_valor.sort_values().plot(kind='barh', ax=ax_prod_val)
        ax_prod_val.set_title('Top 10 Produtos por Valor de Venda (R$)', fontsize=16)
        for patch in ax_prod_val.patches:
            width = patch.get_width()
            y = patch.get_y()
            height = patch.get_height()
            ax_prod_val.text(width * 1.01, y + height / 2, f'{formatar_moeda_str(width)}', va='center')
        ax_prod_val.set_xlim(right=ax_prod_val.get_xlim()[1] * 1.25)
        fig_produto_valor.tight_layout()
        plt.savefig(_asset_path(assets_dir, nome_grafico_produtos_valor), dpi=150)
        plt.close(fig_produto_valor)
        update_status("Gráfico 'Top 10 Produtos por Valor' criado.")

        nome_grafico_produtos_qtd = 'grafico_top_produtos_qtd.png'
        fig_produto_qtd, ax_prod_qtd = plt.subplots(figsize=(12, 8))
        vendas_por_produto_qtd.sort_values().plot(kind='barh', ax=ax_prod_qtd)
        ax_prod_qtd.set_title('Top 10 Itens por Quantidade Vendida (Unidades)', fontsize=16)
        for patch in ax_prod_qtd.patches:
            width = patch.get_width()
            y = patch.get_y()
            height = patch.get_height()
            ax_prod_qtd.text(width * 1.01, y + height / 2, f'{int(width)} un', va='center')
        ax_prod_qtd.set_xlim(right=ax_prod_qtd.get_xlim()[1] * 1.25)
        fig_produto_qtd.tight_layout()
        plt.savefig(_asset_path(assets_dir, nome_grafico_produtos_qtd), dpi=150)
        plt.close(fig_produto_qtd)
        update_status("Gráfico 'Top 10 Itens por Quantidade' criado.")
        
        nome_grafico_mapa_calor = gerar_mapa_calor(vendas_por_estado, log_queue, assets_dir)
        nome_grafico_pizza_estado = None
        nome_grafico_outros_estados = None
        if not vendas_por_estado.empty:
            top_5_estados = vendas_por_estado.head(5)
            soma_outros = vendas_por_estado.iloc[5:].sum()
            dados_pizza = top_5_estados._append(pd.Series({'Outros': soma_outros})) if soma_outros > 0 else top_5_estados
            
            fig_pizza, ax_pizza = plt.subplots(figsize=(10, 7))
            ax_pizza.pie(dados_pizza, labels=dados_pizza.index, autopct='%1.1f%%', startangle=90)
            ax_pizza.set_title('Distribuição Percentual de Vendas por Estado', fontsize=16)
            ax_pizza.axis('equal')
            nome_grafico_pizza_estado = 'grafico_pizza_estado.png'
            plt.savefig(_asset_path(assets_dir, nome_grafico_pizza_estado), dpi=150)
            plt.close(fig_pizza)
            update_status("Gráfico 'Pizza de Estados' criado.")
            
            outros_estados = vendas_por_estado.iloc[5:]
            if not outros_estados.empty:
                fig_outros, ax_outros = plt.subplots(figsize=(12, 6))
                sns.barplot(y=outros_estados.index, x=outros_estados.values, ax=ax_outros, orient='h')
                ax_outros.set_title('Detalhamento de Vendas - Outros Estados', fontsize=16)
                for patch in ax_outros.patches:
                    width = patch.get_width()
                    y = patch.get_y()
                    height = patch.get_height()
                    ax_outros.text(width * 1.01, y + height / 2, f' {formatar_moeda_str(width)}', va='center')
                ax_outros.set_xlim(right=ax_outros.get_xlim()[1] * 1.25)
                fig_outros.tight_layout()
                nome_grafico_outros_estados = 'grafico_outros_estados.png'
                plt.savefig(_asset_path(assets_dir, nome_grafico_outros_estados), dpi=150)
                plt.close(fig_outros)
                update_status("Gráfico 'Outros Estados' criado.")

        update_status("Montando o relatório PDF...")
        template_path = resource_path('template.html')
        with open(template_path, 'r', encoding='utf-8') as f:
            template_html = f.read()
            
        caminho_logo_final = obter_caminho_logo_absoluto()
        
        total_vendas = df_final['Vlr.Total'].sum()
        num_pedidos = df_final['Nro.Pedido'].nunique()
        ticket_medio = total_vendas / num_pedidos if num_pedidos > 0 else 0
        
        tabela_html = ''
        for vendedor, row in analise_vendedor.iterrows():
            total_formatado = formatar_moeda_str(row['Valor_Total_Vendido'])
            pedidos_formatado = int(row['Num_Pedidos'])
            ticket_medio_formatado = formatar_moeda_str(row['Ticket_Medio'])
            tabela_html += f"<tr><td>{vendedor}</td><td style='text-align:right;'>{total_formatado}</td><td style='text-align:center;'>{pedidos_formatado}</td><td style='text-align:right;'>{ticket_medio_formatado}</td></tr>"

        # --- LÓGICA PARA O RESUMO DE PARETO ---
        clientes_pareto_df = df_pareto[df_pareto['% Acumulado'] <= 80]
        num_clientes_pareto = len(clientes_pareto_df) + 1 
        texto_resumo_pareto = f"Análise de Pareto: Cerca de 80% do valor total de vendas do período está concentrado em <strong>{num_clientes_pareto} clientes</strong>."
        update_status(f"Resumo de Pareto calculado: {num_clientes_pareto} clientes representam 80% das vendas.")
        
        html_final = template_html.replace('{{TITULO_RELATORIO}}', f"Relatório de Vendas - {periodo_label}")
        html_final = html_final.replace('{{DATA_GERACAO}}', datetime.now().strftime('%d/%m/%Y %H:%M:%S'))
        html_final = html_final.replace('{{CAMINHO_LOGO}}', caminho_logo_final)
        html_final = html_final.replace('{{VALOR_TOTAL_VENDAS}}', formatar_moeda_str(total_vendas))
        html_final = html_final.replace('{{NUMERO_PEDIDOS}}', str(num_pedidos))
        html_final = html_final.replace('{{TICKET_MEDIO}}', formatar_moeda_str(ticket_medio))
        html_final = html_final.replace('{{GRAFICO_VENDAS_TEMPO_PATH}}', nome_grafico_vendas_tempo)
        html_final = html_final.replace('{{GRAFICO_VENDEDOR_PATH}}', nome_grafico_vendedor)
        html_final = html_final.replace('{{TABELA_VENDEDORES}}', tabela_html)
        html_final = html_final.replace('{{GRAFICO_PARETO_CLIENTE_PATH}}', nome_grafico_pareto_cliente)
        html_final = html_final.replace('{{GRAFICO_CLIENTES_PATH}}', nome_grafico_clientes)
        html_final = html_final.replace('{{GRAFICO_PRODUTOS_VALOR_PATH}}', nome_grafico_produtos_valor)
        html_final = html_final.replace('{{GRAFICO_PRODUTOS_QTD_PATH}}', nome_grafico_produtos_qtd)
        
        # --- SUBSTITUIÇÃO NO HTML (AGORA ATIVA) ---
        html_final = html_final.replace('{{TEXTO_RESUMO_PARETO}}', texto_resumo_pareto) 

        if nome_grafico_mapa_calor:
            html_final = html_final.replace('{{SECAO_MAPA_CALOR}}', f'<div class="new-page"><h2>Distribuição Geográfica das Vendas</h2><img src="{nome_grafico_mapa_calor}" alt="Mapa de Calor" class="chart"></div>')
        else:
            html_final = html_final.replace('{{SECAO_MAPA_CALOR}}', '')
            
        if nome_grafico_pizza_estado:
            percentual_outros = (soma_outros / vendas_por_estado.sum()) * 100 if vendas_por_estado.sum() > 0 else 0
            lista_estados_outros = ", ".join(outros_estados.index) if 'outros_estados' in locals() and not outros_estados.empty else "Nenhum"
            
            html_final = html_final.replace('{{GRAFICO_PIZZA_ESTADO_PATH}}', nome_grafico_pizza_estado)
            html_final = html_final.replace('{{PERCENTUAL_OUTROS}}', f'{percentual_outros:.1f}')
            html_final = html_final.replace('{{LISTA_ESTADOS_OUTROS}}', lista_estados_outros)
            
            if nome_grafico_outros_estados:
                html_final = html_final.replace('{{GRAFICO_OUTROS_ESTADOS_HTML}}', f'<img src="{nome_grafico_outros_estados}" alt="Detalhamento de Outros Estados" class="chart">')
            else:
                html_final = html_final.replace('{{GRAFICO_OUTROS_ESTADOS_HTML}}', '')
        else:
            html_final = html_final.replace('<img src="{{GRAFICO_PIZZA_ESTADO_PATH}}" alt="Distribuição de Vendas por Estado" class="chart">', '<p><i>Não foram encontrados dados de vendas por estado para gerar este gráfico.</i></p>')
            html_final = html_final.replace('<p>A categoria "Outros" (se houver) representa <strong>{{PERCENTUAL_OUTROS}}%</strong> do total e inclui os seguintes estados: {{LISTA_ESTADOS_OUTROS}}.</p>', '')
            html_final = html_final.replace('{{GRAFICO_OUTROS_ESTADOS_HTML}}', '')

        HTML(string=html_final, base_url=assets_dir).write_pdf(save_path)
        update_status(f"Relatório PDF '{os.path.basename(save_path)}' gerado com sucesso!")
        
    except Exception as e:
        update_status(f"ERRO CRÍTICO: {e}")
        messagebox.showerror("Erro Crítico", f"Ocorreu um erro inesperado:\n{e}")
        raise e
    finally:
        if assets_dir:
            shutil.rmtree(assets_dir, ignore_errors=True)

# --- FUNÇÃO 2 (ORIGINAL E INTACTA) ---
def gerar_relatorio_comparativo(caminho_dados_A, caminho_dados_B, log_queue, save_path):
    def update_status(message):
        log_queue.put(message)
    assets_dir = None
    try:
        assets_dir = tempfile.mkdtemp(prefix="relatorio_assets_")
        update_status("Iniciando relatório comparativo...")
        
        df_A = pd.read_excel(caminho_dados_A)
        df_B = pd.read_excel(caminho_dados_B)
        update_status("Ficheiros de dados lidos com sucesso.")

        df_A['Dt.Pedido'] = pd.to_datetime(df_A['Dt.Pedido'])
        df_B['Dt.Pedido'] = pd.to_datetime(df_B['Dt.Pedido'])
        
        label_A = df_A['Dt.Pedido'].min().strftime('%B/%Y').capitalize()
        label_B = df_B['Dt.Pedido'].min().strftime('%B/%Y').capitalize()
        update_status(f"Comparando Período A ({label_A}) com Período B ({label_B}).")
        
        # --- Análise de KPIs ---
        kpis = {}
        for df, label in [(df_A, 'A'), (df_B, 'B')]:
            vendas = df['Vlr.Total'].sum()
            pedidos = df['Nro.Pedido'].nunique()
            ticket = vendas / pedidos if pedidos > 0 else 0
            kpis[label] = {'vendas': vendas, 'pedidos': pedidos, 'ticket': ticket}

        def calcular_variacao(val_A, val_B):
            if val_A > 0:
                var = ((val_B - val_A) / val_A) * 100
                cor = "positive" if var >= 0 else "negative"
                return f"{var:,.2f}%".replace(',', 'X').replace('.', ',').replace('X', '.'), cor
            return "N/A", ""

        var_ven, cor_ven = calcular_variacao(kpis['A']['vendas'], kpis['B']['vendas'])
        var_ped, cor_ped = calcular_variacao(kpis['A']['pedidos'], kpis['B']['pedidos'])
        var_tic, cor_tic = calcular_variacao(kpis['A']['ticket'], kpis['B']['ticket'])
        
        # --- Análise de Churn ---
        update_status("Iniciando análise de Clientes Inativos (Churn)...")
        clientes_A = set(df_A['Cliente'].unique())
        clientes_B = set(df_B['Cliente'].unique())
        clientes_inativos = clientes_A - clientes_B
        
        tabela_churn_html = ""
        resumo_churn = ""
        if not clientes_inativos:
            tabela_churn_html = "<tr><td colspan='2'>Nenhum cliente importante ficou inativo neste período.</td></tr>"
            resumo_churn = "<p>Nenhum cliente relevante ficou inativo neste período.</p>"
        else:
            df_inativos = df_A[df_A['Cliente'].isin(clientes_inativos)]
            vendas_inativos = df_inativos.groupby('Cliente')['Vlr.Total'].sum().sort_values(ascending=False)
            valor_total_inativo = vendas_inativos.sum()
            
            resumo_churn = f"<p>No total, <strong>{len(clientes_inativos)} clientes</strong> que representavam <strong>{formatar_moeda_str(valor_total_inativo)}</strong> em vendas em {label_A} não compraram em {label_B}.</p>"
            
            for cliente, valor in vendas_inativos.head(15).items():
                tabela_churn_html += f"<tr><td>{cliente}</td><td style='text-align:right;'>{formatar_moeda_str(valor)}</td></tr>"
        update_status(f"Análise de Churn concluída: {len(clientes_inativos)} clientes inativos.")

        # --- Gráficos Comparativos ---
        update_status("Gerando gráficos comparativos...")
        
        def formatar_moeda_simples(valor, pos=None):
            return f'R$ {int(valor):,.0f}'.replace(',', '.')
        
        def criar_grafico_comparativo(df1, df2, label1, label2, coluna_grupo, coluna_valor, titulo, nome_ficheiro):
            dados1 = df1.groupby(coluna_grupo)[coluna_valor].sum().rename(label1)
            dados2 = df2.groupby(coluna_grupo)[coluna_valor].sum().rename(label2)
            
            df_comp = pd.merge(dados1, dados2, on=coluna_grupo, how='outer').fillna(0)
            df_comp['Total'] = df_comp[label1] + df_comp[label2]
            df_comp = df_comp.sort_values(by='Total', ascending=False).head(10).drop(columns=['Total'])
            
            df_grafico = df_comp.reset_index().melt(id_vars=coluna_grupo, var_name='Período', value_name='Vendas')
            
            fig, ax = plt.subplots(figsize=(12, 8))
            sns.barplot(data=df_grafico, y=coluna_grupo, x='Vendas', hue='Período', ax=ax)
            ax.set_title(titulo, fontsize=16)
            ax.xaxis.set_major_formatter(FuncFormatter(formatar_moeda_simples))
            
            for container in ax.containers:
                ax.bar_label(container, fmt=lambda x: formatar_moeda_simples(x) if x > 0 else '', padding=3, fontsize=8)

            ax.set_xlim(right=ax.get_xlim()[1] * 1.3)
            fig.tight_layout()
            plt.savefig(_asset_path(assets_dir, nome_ficheiro), dpi=150)
            plt.close(fig)
            update_status(f"Gráfico '{titulo}' criado.")
            return nome_ficheiro

        nome_grafico_vendedor = criar_grafico_comparativo(df_A, df_B, label_A, label_B, 'Vendedor', 'Vlr.Total', 'Comparativo de Vendas por Vendedor', 'grafico_comp_vendedor.png')
        nome_grafico_produto = criar_grafico_comparativo(df_A, df_B, label_A, label_B, 'Descrição', 'Vlr.Total', 'Comparativo de Top 10 Produtos', 'grafico_comp_produto.png')
        nome_grafico_cliente = criar_grafico_comparativo(df_A, df_B, label_A, label_B, 'Cliente', 'Vlr.Total', 'Comparativo de Top 10 Clientes', 'grafico_comp_cliente.png')

        # Montagem do HTML
        template_path = resource_path('template_comparativo.html')
        with open(template_path, 'r', encoding='utf-8') as f:
            template_html = f.read()
        
        caminho_logo_final = obter_caminho_logo_absoluto()
        
        html_final = template_html.replace('{{TITULO_RELATORIO}}', f'Relatório Comparativo: {label_A} vs {label_B}')
        html_final = html_final.replace('{{DATA_GERACAO}}', datetime.now().strftime('%d/%m/%Y %H:%M:%S'))
        html_final = html_final.replace('{{CAMINHO_LOGO}}', caminho_logo_final)
        html_final = html_final.replace('{{LABEL_A}}', label_A)
        html_final = html_final.replace('{{LABEL_B}}', label_B)

        html_final = html_final.replace('{{VALOR_TOTAL_A}}', formatar_moeda_str(kpis['A']['vendas']))
        html_final = html_final.replace('{{VALOR_TOTAL_B}}', formatar_moeda_str(kpis['B']['vendas']))
        html_final = html_final.replace('{{VARIACAO_VALOR}}', var_ven)
        html_final = html_final.replace('{{COR_VARIACAO_VALOR}}', cor_ven)
        
        html_final = html_final.replace('{{NUM_PEDIDOS_A}}', str(kpis['A']['pedidos']))
        html_final = html_final.replace('{{NUM_PEDIDOS_B}}', str(kpis['B']['pedidos']))
        html_final = html_final.replace('{{VARIACAO_PEDIDOS}}', var_ped)
        html_final = html_final.replace('{{COR_VARIACAO_PEDIDOS}}', cor_ped)
        
        html_final = html_final.replace('{{TICKET_MEDIO_A}}', formatar_moeda_str(kpis['A']['ticket']))
        html_final = html_final.replace('{{TICKET_MEDIO_B}}', formatar_moeda_str(kpis['B']['ticket']))
        html_final = html_final.replace('{{VARIACAO_TICKET}}', var_tic)
        html_final = html_final.replace('{{COR_VARIACAO_TICKET}}', cor_tic)

        html_final = html_final.replace('{{GRAFICO_COMP_VENDEDOR_PATH}}', nome_grafico_vendedor)
        html_final = html_final.replace('{{GRAFICO_COMP_PRODUTO_PATH}}', nome_grafico_produto)
        html_final = html_final.replace('{{GRAFICO_COMP_CLIENTE_PATH}}', nome_grafico_cliente)
        html_final = html_final.replace('{{RESUMO_CHURN}}', resumo_churn)
        html_final = html_final.replace('{{TABELA_CHURN}}', tabela_churn_html)

        HTML(string=html_final, base_url=assets_dir).write_pdf(save_path)
        update_status(f"Relatório Comparativo '{os.path.basename(save_path)}' gerado com sucesso!")
        
    except Exception as e:
        update_status(f"ERRO CRÍTICO: {e}")
        messagebox.showerror("Erro Crítico", f"Ocorreu um erro inesperado:\n{e}")
        raise e
    finally:
        if assets_dir:
            shutil.rmtree(assets_dir, ignore_errors=True)

# --- FUNÇÃO 3 (ORIGINAL E INTACTA) ---
def gerar_relatorio_anual(caminho_dados_anuais, log_queue, save_path):
    def update_status(message):
        log_queue.put(message)
    assets_dir = None
    try:
        assets_dir = tempfile.mkdtemp(prefix="relatorio_assets_")
        update_status("Iniciando Análise Anual...")
        df = pd.read_excel(caminho_dados_anuais)
        update_status("Ficheiro de dados anuais lido com sucesso.")

        df['Dt.Pedido'] = pd.to_datetime(df['Dt.Pedido'])
        
        venda_total = df['Vlr.Total'].sum()
        pedidos_unicos = df['Nro.Pedido'].nunique()
        ticket_medio = venda_total / pedidos_unicos if pedidos_unicos > 0 else 0
        unidades_vendidas = df['Qtde'].sum()
        
        update_status("KPIs anuais calculados. Gerando gráficos...")
        
        def formatar_inteiro(valor, pos=None):
            return f'R$ {int(valor):,.0f}'.replace(',', '.')

        # Gráfico de Evolução Mensal
        vendas_mensais = df.set_index('Dt.Pedido')['Vlr.Total'].resample('M').sum()
        fig_mensal, ax_mensal = plt.subplots(figsize=(12, 6))
        vendas_mensais.index = vendas_mensais.index.strftime('%b')
        sns.barplot(x=vendas_mensais.index, y=vendas_mensais.values, ax=ax_mensal, color='royalblue')
        ax_mensal.set_title('Evolução Mensal de Vendas', fontsize=16)
        for container in ax_mensal.containers:
            ax_mensal.bar_label(container, fmt=lambda x: formatar_inteiro(x) if x > 0 else '', padding=3)
        fig_mensal.tight_layout()
        nome_grafico_mensal = 'grafico_evolucao_mensal.png'
        plt.savefig(_asset_path(assets_dir, nome_grafico_mensal), dpi=150)
        plt.close(fig_mensal)
        update_status("Gráfico 'Evolução Mensal' criado.")

        # Gráfico de Vendas Trimestral
        vendas_trimestrais = df.set_index('Dt.Pedido')['Vlr.Total'].resample('Q').sum()
        fig_trimestral, ax_trimestral = plt.subplots(figsize=(8, 8))
        
        def autopct_format_anual(pct):
            total = vendas_trimestrais.sum()
            val = int(round(pct * total / 100.0))
            return f'{pct:.1f}%\n({formatar_inteiro(val)})'
            
        ax_trimestral.pie(vendas_trimestrais, labels=[f"T{i.quarter}" for i in vendas_trimestrais.index], autopct=autopct_format_anual, startangle=90)
        ax_trimestral.set_title('Distribuição das Vendas por Trimestre', fontsize=16)
        ax_trimestral.axis('equal')
        nome_grafico_trimestral = 'grafico_vendas_trimestral.png'
        plt.savefig(_asset_path(assets_dir, nome_grafico_trimestral), dpi=150)
        plt.close(fig_trimestral)
        update_status("Gráfico 'Vendas por Trimestre' criado.")

        # Gráficos de Rankings Anuais
        top_produtos = df.groupby('Descrição')['Vlr.Total'].sum().nlargest(10)
        top_produtos_qtd = df.groupby('Descrição')['Qtde'].sum().nlargest(10)
        top_clientes = df.groupby('Cliente')['Vlr.Total'].sum().nlargest(10)

        fig_prods, ax_prods = plt.subplots(figsize=(12, 8))
        top_produtos.sort_values().plot(kind='barh', ax=ax_prods)
        ax_prods.set_title('Top 10 Produtos do Ano (por Valor de Venda)', fontsize=16)
        for patch in ax_prods.patches:
            width = patch.get_width()
            y = patch.get_y()
            height = patch.get_height()
            ax_prods.text(width * 1.01, y + height / 2, f'{formatar_moeda_str(width)}', va='center')
        ax_prods.set_xlim(right=ax_prods.get_xlim()[1] * 1.25)
        fig_prods.tight_layout()
        nome_grafico_top_produtos = 'grafico_top_produtos.png'
        plt.savefig(_asset_path(assets_dir, nome_grafico_top_produtos), dpi=150)
        plt.close(fig_prods)
        update_status("Gráfico 'Top 10 Produtos por Valor' criado.")

        fig_prods_qtd, ax_prods_qtd = plt.subplots(figsize=(12, 8))
        top_produtos_qtd.sort_values().plot(kind='barh', ax=ax_prods_qtd)
        ax_prods_qtd.set_title('Top 10 Produtos do Ano (por Quantidade Vendida)', fontsize=16)
        for patch in ax_prods_qtd.patches:
            width = patch.get_width()
            y = patch.get_y()
            height = patch.get_height()
            ax_prods_qtd.text(width * 1.01, y + height / 2, f'{int(width)} un', va='center')
        ax_prods_qtd.set_xlim(right=ax_prods_qtd.get_xlim()[1] * 1.25)
        fig_prods_qtd.tight_layout()
        nome_grafico_top_produtos_qtd = 'grafico_top_produtos_qtd.png'
        plt.savefig(_asset_path(assets_dir, nome_grafico_top_produtos_qtd), dpi=150)
        plt.close(fig_prods_qtd)
        update_status("Gráfico 'Top 10 Produtos por Qtd' criado.")

        fig_clientes, ax_clientes = plt.subplots(figsize=(12, 8))
        top_clientes.sort_values().plot(kind='barh', ax=ax_clientes)
        ax_clientes.set_title('Top 10 Clientes do Ano (por Valor de Compra)', fontsize=16)
        for patch in ax_clientes.patches:
            width = patch.get_width()
            y = patch.get_y()
            height = patch.get_height()
            ax_clientes.text(width * 1.01, y + height / 2, f'{formatar_moeda_str(width)}', va='center')
        ax_clientes.set_xlim(right=ax_clientes.get_xlim()[1] * 1.25)
        fig_clientes.tight_layout()
        nome_grafico_top_clientes = 'grafico_top_clientes.png'
        plt.savefig(_asset_path(assets_dir, nome_grafico_top_clientes), dpi=150)
        plt.close(fig_clientes)
        update_status("Gráfico 'Top 10 Clientes' criado.")

        update_status("Montando o relatório PDF anual...")
        template_path = resource_path('template_anual.html')
        with open(template_path, 'r', encoding='utf-8') as f:
            template_html = f.read()

        caminho_logo_final = obter_caminho_logo_absoluto()

        html_final = template_html.replace('{{DATA_GERACAO}}', datetime.now().strftime('%d/%m/%Y %H:%M:%S'))
        html_final = html_final.replace('{{CAMINHO_LOGO}}', caminho_logo_final)
        html_final = html_final.replace('{{VENDA_TOTAL_ANUAL}}', formatar_moeda_str(venda_total))
        html_final = html_final.replace('{{NUM_PEDIDOS_ANUAL}}', str(pedidos_unicos))
        html_final = html_final.replace('{{TICKET_MEDIO_ANUAL}}', formatar_moeda_str(ticket_medio))
        html_final = html_final.replace('{{UNIDADES_VENDIDAS_ANUAL}}', str(int(unidades_vendidas)))
        
        html_final = html_final.replace('{{GRAFICO_EVOLUCAO_MENSAL_PATH}}', nome_grafico_mensal)
        html_final = html_final.replace('{{GRAFICO_VENDAS_TRIMESTRAL_PATH}}', nome_grafico_trimestral)
        html_final = html_final.replace('{{GRAFICO_TOP_PRODUTOS_PATH}}', nome_grafico_top_produtos)
        html_final = html_final.replace('{{GRAFICO_TOP_PRODUTOS_QTD_ANUAL_PATH}}', nome_grafico_top_produtos_qtd)
        html_final = html_final.replace('{{GRAFICO_TOP_CLIENTES_PATH}}', nome_grafico_top_clientes)

        HTML(string=html_final, base_url=assets_dir).write_pdf(save_path)
        update_status(f"Relatório Anual '{os.path.basename(save_path)}' gerado com sucesso!")

    except Exception as e:
        update_status(f"ERRO CRÍTICO na Análise Anual: {e}")
        messagebox.showerror("Erro Crítico", f"Ocorreu um erro inesperado:\n{e}")
        raise e
    finally:
        if assets_dir:
            shutil.rmtree(assets_dir, ignore_errors=True)
        
# --- FUNÇÃO 4 (ORIGINAL E INTACTA) ---
def gerar_relatorio_comparativo_anual(caminho_ano_A, caminho_ano_B, log_queue, save_path):
    def update_status(message):
        log_queue.put(message)
    assets_dir = None
    try:
        assets_dir = tempfile.mkdtemp(prefix="relatorio_assets_")
        update_status("--- INICIANDO RELATÓRIO COMPARATIVO ANUAL ---")
        
        update_status(f"Lendo dados do Ano A: {os.path.basename(caminho_ano_A)}")
        df_A = pd.read_excel(caminho_ano_A)
        df_A['Dt.Pedido'] = pd.to_datetime(df_A['Dt.Pedido'])

        update_status(f"Lendo dados do Ano B: {os.path.basename(caminho_ano_B)}")
        df_B = pd.read_excel(caminho_ano_B)
        df_B['Dt.Pedido'] = pd.to_datetime(df_B['Dt.Pedido'])

        ano_A = df_A['Dt.Pedido'].dt.year.iloc[0]
        ano_B = df_B['Dt.Pedido'].dt.year.iloc[0]
        update_status(f"Anos identificados: {ano_A} vs {ano_B}")

        update_status("Calculando KPIs principais...")
        venda_total_A = df_A['Vlr.Total'].sum()
        pedidos_A = df_A['Nro.Pedido'].nunique()
        ticket_medio_A = venda_total_A / pedidos_A if pedidos_A > 0 else 0
        
        venda_total_B = df_B['Vlr.Total'].sum()
        pedidos_B = df_B['Nro.Pedido'].nunique()
        ticket_medio_B = venda_total_B / pedidos_B if pedidos_B > 0 else 0
        
        def calcular_variacao_anual(val_A, val_B):
            if val_A > 0:
                var = ((val_B - val_A) / val_A) * 100
                cor = "positive" if var >= 0 else "negative"
                sinal = "+" if var >= 0 else ""
                return f"{sinal}{var:.2f}%", cor
            return "N/A", ""

        var_venda, cor_venda = calcular_variacao_anual(venda_total_A, venda_total_B)
        var_pedidos, cor_pedidos = calcular_variacao_anual(pedidos_A, pedidos_B)
        var_ticket, cor_ticket = calcular_variacao_anual(ticket_medio_A, ticket_medio_B)
        update_status("KPIs calculados. Gerando gráficos...")
        
        # Gráfico 1: Evolução Mensal Comparativa
        df_A['Mes'] = df_A['Dt.Pedido'].dt.month
        vendas_mensais_A = df_A.groupby('Mes')['Vlr.Total'].sum().reindex(range(1, 13), fill_value=0)
        df_B['Mes'] = df_B['Dt.Pedido'].dt.month
        vendas_mensais_B = df_B.groupby('Mes')['Vlr.Total'].sum().reindex(range(1, 13), fill_value=0)
        nomes_meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
        
        fig, ax = plt.subplots(figsize=(12, 6))
        bar_width = 0.4
        index = np.arange(len(nomes_meses))
        ax.bar(index - bar_width/2, vendas_mensais_A, bar_width, label=f'{ano_A}', color=CORES_GRAFICOS[0])
        ax.bar(index + bar_width/2, vendas_mensais_B, bar_width, label=f'{ano_B}', color=CORES_GRAFICOS[1])
        ax.set_title(f'Comparativo de Vendas Mensais ({ano_A} vs {ano_B})', fontsize=16)
        ax.set_xticks(index)
        ax.set_xticklabels(nomes_meses)
        ax.yaxis.set_major_formatter(FuncFormatter(lambda val, pos: formatar_moeda_str(val)))
        ax.legend()
        fig.tight_layout()
        nome_grafico_evolucao = 'grafico_comp_anual_mensal.png'
        plt.savefig(_asset_path(assets_dir, nome_grafico_evolucao), dpi=150)
        plt.close(fig)
        update_status("Gráfico 'Comparativo Mensal' criado.")

        # Gráfico 2: Comparativo Top 10 Produtos
        top_produtos_B = df_B.groupby('Descrição')['Vlr.Total'].sum().nlargest(10)
        top_produtos_A = df_A[df_A['Descrição'].isin(top_produtos_B.index)].groupby('Descrição')['Vlr.Total'].sum().reindex(top_produtos_B.index, fill_value=0)
        df_comp_produtos = pd.DataFrame({f'{ano_A}': top_produtos_A, f'{ano_B}': top_produtos_B}).sort_values(by=f'{ano_B}', ascending=True)
        
        fig, ax = plt.subplots(figsize=(12, 8))
        df_comp_produtos.plot(kind='barh', ax=ax, color=[CORES_GRAFICOS[0], CORES_GRAFICOS[1]])
        ax.set_title(f'Comparativo de Vendas: Top 10 Produtos de {ano_B}', fontsize=16)
        ax.xaxis.set_major_formatter(FuncFormatter(lambda val, pos: formatar_moeda_str(val)))
        fig.tight_layout()
        nome_grafico_produtos = 'grafico_comp_anual_produtos.png'
        plt.savefig(_asset_path(assets_dir, nome_grafico_produtos), dpi=150)
        plt.close(fig)
        update_status("Gráfico 'Comparativo Top Produtos' criado.")
        
        # Gráfico 3: Comparativo Top 10 Clientes
        top_clientes_B = df_B.groupby('Cliente')['Vlr.Total'].sum().nlargest(10)
        top_clientes_A = df_A[df_A['Cliente'].isin(top_clientes_B.index)].groupby('Cliente')['Vlr.Total'].sum().reindex(top_clientes_B.index, fill_value=0)
        df_comp_clientes = pd.DataFrame({f'{ano_A}': top_clientes_A, f'{ano_B}': top_clientes_B}).sort_values(by=f'{ano_B}', ascending=True)
        
        fig, ax = plt.subplots(figsize=(12, 8))
        df_comp_clientes.plot(kind='barh', ax=ax, color=[CORES_GRAFICOS[0], CORES_GRAFICOS[1]])
        ax.set_title(f'Comparativo de Compras: Top 10 Clientes de {ano_B}', fontsize=16)
        ax.xaxis.set_major_formatter(FuncFormatter(lambda val, pos: formatar_moeda_str(val)))
        fig.tight_layout()
        nome_grafico_clientes = 'grafico_comp_anual_clientes.png'
        plt.savefig(_asset_path(assets_dir, nome_grafico_clientes), dpi=150)
        plt.close(fig)
        update_status("Gráfico 'Comparativo Top Clientes' criado.")

        update_status("Montando relatório PDF...")
        template_path = resource_path('template_comparativo_anual.html')
        with open(template_path, 'r', encoding='utf-8') as f:
            template_html = f.read()

        html_final = template_html
        html_final = html_final.replace('{{DATA_GERACAO}}', datetime.now().strftime('%d/%m/%Y %H:%M:%S'))
        html_final = html_final.replace('{{CAMINHO_LOGO}}', obter_caminho_logo_absoluto())
        html_final = html_final.replace('{{ANO_A}}', str(ano_A))
        html_final = html_final.replace('{{ANO_B}}', str(ano_B))
        html_final = html_final.replace('{{VENDA_A}}', formatar_moeda_str(venda_total_A))
        html_final = html_final.replace('{{VENDA_B}}', formatar_moeda_str(venda_total_B))
        html_final = html_final.replace('{{VARIACAO_VENDA}}', var_venda)
        html_final = html_final.replace('{{COR_VARIACAO_VENDA}}', cor_venda)
        html_final = html_final.replace('{{PEDIDOS_A}}', str(int(pedidos_A)))
        html_final = html_final.replace('{{PEDIDOS_B}}', str(int(pedidos_B)))
        html_final = html_final.replace('{{VARIACAO_PEDIDOS}}', var_pedidos)
        html_final = html_final.replace('{{COR_VARIACAO_PEDIDOS}}', cor_pedidos)
        html_final = html_final.replace('{{TICKET_A}}', formatar_moeda_str(ticket_medio_A))
        html_final = html_final.replace('{{TICKET_B}}', formatar_moeda_str(ticket_medio_B))
        html_final = html_final.replace('{{VARIACAO_TICKET}}', var_ticket)
        html_final = html_final.replace('{{COR_VARIACAO_TICKET}}', cor_ticket)
        html_final = html_final.replace('{{GRAFICO_EVOLUCAO_MENSAL_COMP_PATH}}', nome_grafico_evolucao)
        html_final = html_final.replace('{{GRAFICO_TOP_PRODUTOS_COMP_PATH}}', nome_grafico_produtos)
        html_final = html_final.replace('{{GRAFICO_TOP_CLIENTES_COMP_PATH}}', nome_grafico_clientes)

        HTML(string=html_final, base_url=assets_dir).write_pdf(save_path)
        update_status(f"Relatório Comparativo Anual '{os.path.basename(save_path)}' gerado com sucesso!")

    except Exception as e:
        update_status(f"ERRO CRÍTICO no Comparativo Anual: {e}")
        messagebox.showerror("Erro Crítico", f"Ocorreu um erro inesperado no Comparativo Anual:\n{e}")
        raise e
    finally:
        if assets_dir:
            shutil.rmtree(assets_dir, ignore_errors=True)

# --- CLASSE DA APLICAÇÃO (ATUALIZADA) ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Relatórios v19.5")
        self.root.state('zoomed') 
        
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(pady=10, padx=10, fill="both", expand=True)
        
        self.tab1 = ttk.Frame(self.notebook, padding="10")
        self.tab2 = ttk.Frame(self.notebook, padding="10")
        self.tab3 = ttk.Frame(self.notebook, padding="10")
        self.tab4 = ttk.Frame(self.notebook, padding="10") # Nova Aba
        
        self.notebook.add(self.tab1, text=' Relatório de Período ')
        self.notebook.add(self.tab2, text=' Comparativo & Clientes Inativos ')
        self.notebook.add(self.tab3, text=' Análise Anual ')
        self.notebook.add(self.tab4, text=' Comparativo Anual ') # Nova Aba
        
        status_frame = ttk.LabelFrame(root, text=" Status do Processamento ", padding="10")
        status_frame.pack(padx=10, pady=(0, 10), fill='x')
        self.status_box = scrolledtext.ScrolledText(status_frame, height=10, state='disabled', font=("Consolas", 9))
        self.status_box.pack(fill='x', expand=True)
        
        self.create_tab1_widgets()
        self.create_tab2_widgets()
        self.create_tab3_widgets()
        self.create_tab4_widgets() # Nova Aba
        
        self.log_queue = queue.Queue()
        self.root.after(100, self.process_log_queue)
        self.update_status("Bem-vindo ao Gerador de Relatórios v19.5!")

    def process_log_queue(self):
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.update_status(message)
        except queue.Empty:
            pass
        self.root.after(100, self.process_log_queue)
        
    def create_tab1_widgets(self):
        self.path_consolidado = tk.StringVar()
        frame = ttk.LabelFrame(self.tab1, text=" 1. Selecionar Ficheiro de Dados ", padding="10")
        frame.pack(fill='x', expand=True, pady=(0, 10))
        
        ttk.Label(frame, text="Ficheiro de Dados:").grid(row=0, column=0, sticky='w', padx=5, pady=8)
        ttk.Entry(frame, textvariable=self.path_consolidado, width=80).grid(row=0, column=1, padx=5, pady=8)
        ttk.Button(frame, text="Procurar...", command=lambda: self.select_file_for_var(self.path_consolidado)).grid(row=0, column=2, padx=5, pady=8)
        
        button_frame = ttk.Frame(self.tab1)
        button_frame.pack(pady=20, padx=10, fill='x')
        self.generate_button_A = ttk.Button(button_frame, text="GERAR RELATÓRIO DE PERÍODO", command=self.start_report_thread, padding="10")
        self.generate_button_A.pack(side='left', expand=True, fill='x', padx=(0, 5))
        self.clear_button_A = ttk.Button(button_frame, text="Limpar", command=lambda: self.path_consolidado.set(""), padding="10")
        self.clear_button_A.pack(side='right', padx=(5, 0))

    def create_tab2_widgets(self):
        self.path_consolidado_A = tk.StringVar()
        self.path_consolidado_B = tk.StringVar()
        
        frame_A = ttk.LabelFrame(self.tab2, text=" 1. Selecionar Período A (o mais antigo) ", padding="10")
        frame_A.pack(fill='x', expand=True, pady=(0, 10))
        ttk.Label(frame_A, text="Ficheiro de Dados A:").grid(row=0, column=0, sticky='w', padx=5, pady=8)
        ttk.Entry(frame_A, textvariable=self.path_consolidado_A, width=80).grid(row=0, column=1, padx=5, pady=8)
        ttk.Button(frame_A, text="Procurar...", command=lambda: self.select_file_for_var(self.path_consolidado_A)).grid(row=0, column=2, padx=5, pady=8)

        frame_B = ttk.LabelFrame(self.tab2, text=" 2. Selecionar Período B (o mais recente) ", padding="10")
        frame_B.pack(fill='x', expand=True, pady=(0, 10))
        ttk.Label(frame_B, text="Ficheiro de Dados B:").grid(row=0, column=0, sticky='w', padx=5, pady=8)
        ttk.Entry(frame_B, textvariable=self.path_consolidado_B, width=80).grid(row=0, column=1, padx=5, pady=8)
        ttk.Button(frame_B, text="Procurar...", command=lambda: self.select_file_for_var(self.path_consolidado_B)).grid(row=0, column=2, padx=5, pady=8)

        button_frame = ttk.Frame(self.tab2)
        button_frame.pack(pady=20, padx=10, fill='x')
        self.generate_button_B = ttk.Button(button_frame, text="GERAR ANÁLISE COMPARATIVA", command=self.start_churn_report_thread, padding="10")
        self.generate_button_B.pack(side='left', expand=True, fill='x', padx=(0, 5))
        self.clear_button_B = ttk.Button(button_frame, text="Limpar", command=lambda: [self.path_consolidado_A.set(""), self.path_consolidado_B.set("")], padding="10")
        self.clear_button_B.pack(side='right', padx=(5, 0))

    def create_tab3_widgets(self):
        self.path_anual = tk.StringVar()
        frame = ttk.LabelFrame(self.tab3, text=" 1. Selecionar Ficheiro de Dados Anuais ", padding="10")
        frame.pack(fill='x', expand=True, pady=(0, 10))
        
        ttk.Label(frame, text="Ficheiro de Dados Anual:").grid(row=0, column=0, sticky='w', padx=5, pady=8)
        ttk.Entry(frame, textvariable=self.path_anual, width=80).grid(row=0, column=1, padx=5, pady=8)
        ttk.Button(frame, text="Procurar...", command=lambda: self.select_file_for_var(self.path_anual)).grid(row=0, column=2, padx=5, pady=8)
        
        button_frame = ttk.Frame(self.tab3)
        button_frame.pack(pady=20, padx=10, fill='x')
        self.generate_button_C = ttk.Button(button_frame, text="GERAR ANÁLISE ANUAL", command=self.start_annual_report_thread, padding="10")
        self.generate_button_C.pack(side='left', expand=True, fill='x', padx=(0, 5))
        self.clear_button_C = ttk.Button(button_frame, text="Limpar", command=lambda: self.path_anual.set(""), padding="10")
        self.clear_button_C.pack(side='right', padx=(5, 0))

    def create_tab4_widgets(self):
        self.path_comp_anual_A = tk.StringVar()
        self.path_comp_anual_B = tk.StringVar()
        
        frame_A = ttk.LabelFrame(self.tab4, text=" 1. Selecionar Ano A (o mais antigo) ", padding="10")
        frame_A.pack(fill='x', expand=True, pady=(0, 10))
        ttk.Label(frame_A, text="Ficheiro Ano A:").grid(row=0, column=0, sticky='w', padx=5, pady=8)
        ttk.Entry(frame_A, textvariable=self.path_comp_anual_A, width=80).grid(row=0, column=1, padx=5, pady=8)
        ttk.Button(frame_A, text="Procurar...", command=lambda: self.select_file_for_var(self.path_comp_anual_A)).grid(row=0, column=2, padx=5, pady=8)

        frame_B = ttk.LabelFrame(self.tab4, text=" 2. Selecionar Ano B (o mais recente) ", padding="10")
        frame_B.pack(fill='x', expand=True, pady=(0, 10))
        ttk.Label(frame_B, text="Ficheiro Ano B:").grid(row=0, column=0, sticky='w', padx=5, pady=8)
        ttk.Entry(frame_B, textvariable=self.path_comp_anual_B, width=80).grid(row=0, column=1, padx=5, pady=8)
        ttk.Button(frame_B, text="Procurar...", command=lambda: self.select_file_for_var(self.path_comp_anual_B)).grid(row=0, column=2, padx=5, pady=8)

        button_frame = ttk.Frame(self.tab4)
        button_frame.pack(pady=20, padx=10, fill='x')
        self.generate_button_D = ttk.Button(button_frame, text="GERAR COMPARATIVO ANUAL", command=self.start_annual_comp_report_thread, padding="10")
        self.generate_button_D.pack(side='left', expand=True, fill='x', padx=(0, 5))
        self.clear_button_D = ttk.Button(button_frame, text="Limpar", command=lambda: [self.path_comp_anual_A.set(""), self.path_comp_anual_B.set("")], padding="10")
        self.clear_button_D.pack(side='right', padx=(5, 0))

    def select_file_for_var(self, path_variable):
        filepath = filedialog.askopenfilename(
            title="Selecione o ficheiro de dados consolidados",
            filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            path_variable.set(filepath)

    def update_status(self, message):
        self.status_box.config(state='normal')
        self.status_box.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.status_box.see(tk.END)
        self.status_box.config(state='disabled')
        self.root.update_idletasks()

    def report_finished(self, success=True):
        self.generate_button_A.config(state='normal')
        self.generate_button_B.config(state='normal')
        self.generate_button_C.config(state='normal')
        self.generate_button_D.config(state='normal') # Re-habilita novo botão
        if success:
            self.log_queue.put(">>>> PROCESSO CONCLUÍDO COM SUCESSO! <<<<")
        else:
            self.log_queue.put(">>>> PROCESSO INTERROMPIDO POR ERRO. <<<<")

    def start_report_thread(self):
        path_consolidado = self.path_consolidado.get()
        if not path_consolidado:
            messagebox.showerror("Ficheiro em Falta", "Por favor, selecione o ficheiro de dados consolidados.")
            return
        
        default_filename = f"Relatorio_Vendas_{os.path.basename(path_consolidado).replace('dados_consolidados_', '').replace('.xlsx', '')}.pdf"
        
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
            title="Salvar Relatório de Período Único", initialfile=default_filename)
        if not save_path:
            self.update_status("Operação de salvar cancelada pelo usuário.")
            return

        self.prepare_to_run()
        thread = threading.Thread(target=self.run_report, args=(path_consolidado, save_path))
        thread.start()
        
    def start_churn_report_thread(self):
        path_A = self.path_consolidado_A.get()
        path_B = self.path_consolidado_B.get()
        if not path_A or not path_B:
            messagebox.showerror("Ficheiros em Falta", "Por favor, selecione os dois ficheiros de dados para a comparação.")
            return
        
        default_filename = "Relatorio_Comparativo.pdf"
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
            title="Salvar Análise Comparativa", initialfile=default_filename)
        if not save_path:
            self.update_status("Operação de salvar cancelada pelo usuário.")
            return

        self.prepare_to_run()
        thread = threading.Thread(target=self.run_churn_report, args=(path_A, path_B, save_path))
        thread.start()

    def start_annual_report_thread(self):
        path_anual = self.path_anual.get()
        if not path_anual:
            messagebox.showerror("Ficheiro em Falta", "Por favor, selecione o ficheiro de dados anuais.")
            return
        
        default_filename = "Relatorio_Analise_Anual.pdf"
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
            title="Salvar Análise Anual", initialfile=default_filename)
        if not save_path:
            self.update_status("Operação de salvar cancelada pelo usuário.")
            return

        self.prepare_to_run()
        thread = threading.Thread(target=self.run_annual_report, args=(path_anual, save_path))
        thread.start()

    def start_annual_comp_report_thread(self):
        path_A = self.path_comp_anual_A.get()
        path_B = self.path_comp_anual_B.get()
        if not path_A or not path_B:
            messagebox.showerror("Ficheiros em Falta", "Por favor, selecione os dois ficheiros de dados anuais para a comparação.")
            return
        
        default_filename = "Relatorio_Comparativo_Anual.pdf"
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
            title="Salvar Comparativo Anual", initialfile=default_filename)
        if not save_path:
            self.update_status("Operação de salvar cancelada pelo usuário.")
            return

        self.prepare_to_run()
        thread = threading.Thread(target=self.run_annual_comp_report, args=(path_A, path_B, save_path))
        thread.start()

    def prepare_to_run(self):
        self.generate_button_A.config(state='disabled')
        self.generate_button_B.config(state='disabled')
        self.generate_button_C.config(state='disabled')
        self.generate_button_D.config(state='disabled') # Desabilita novo botão
        self.status_box.config(state='normal')
        self.status_box.delete('1.0', tk.END)
        self.status_box.config(state='disabled')
        
    def run_report(self, path_consolidado, save_path):
        try:
            gerar_relatorio_completo(path_consolidado, self.log_queue, save_path)
            self.report_finished(success=True)
        except Exception as e:
            self.report_finished(success=False)
            print(f"Erro capturado pela thread da GUI: {e}")
            self.log_queue.put(f"ERRO CRÍTICO: {e}")
            
    def run_churn_report(self, path_A, path_B, save_path):
        try:
            gerar_relatorio_comparativo(path_A, path_B, self.log_queue, save_path)
            self.report_finished(success=True)
        except Exception as e:
            self.report_finished(success=False)
            print(f"Erro capturado pela thread da GUI: {e}")
            self.log_queue.put(f"ERRO CRÍTICO: {e}")
            
    def run_annual_report(self, path_anual, save_path):
        try:
            gerar_relatorio_anual(path_anual, self.log_queue, save_path)
            self.report_finished(success=True)
        except Exception as e:
            self.report_finished(success=False)
            print(f"Erro capturado pela thread da GUI: {e}")
            self.log_queue.put(f"ERRO CRÍTICO: {e}")

    def run_annual_comp_report(self, path_A, path_B, save_path):
        try:
            gerar_relatorio_comparativo_anual(path_A, path_B, self.log_queue, save_path)
            self.report_finished(success=True)
        except Exception as e:
            self.report_finished(success=False)
            print(f"Erro capturado pela thread da GUI: {e}")
            self.log_queue.put(f"ERRO CRÍTICO: {e}")

if __name__ == "__main__":
    log_file_path = "crash_log.txt"
    try:
        multiprocessing.freeze_support()
        root = ThemedTk(theme="arc")
        app = App(root)
        
        if pyi_splash:
            pyi_splash.close()
            
        root.mainloop()

    except Exception as e:
        import traceback
        with open(log_file_path, 'w', encoding='utf-8') as f:
            f.write(f"Ocorreu um erro fatal:\n")
            f.write(str(e) + "\n\n")
            f.write(traceback.format_exc())
# --- FIM DO CÓDIGO ---