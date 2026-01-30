# converter_mapa.py
import geopandas
import pickle
import os

print("Iniciando a conversão do arquivo de mapa...")

# Caminho do arquivo de mapa complexo (fonte)
caminho_mapa_gpkg = os.path.join('geo_data', 'br_states.gpkg')

# Nome do arquivo de mapa simplificado que vamos criar (destino)
caminho_mapa_pkl = 'mapa_brasil_dados.pkl'

try:
    # Verifica se o arquivo de origem existe
    if not os.path.exists(caminho_mapa_gpkg):
        print(f"ERRO: Arquivo de origem '{caminho_mapa_gpkg}' não encontrado.")
        print("Por favor, execute o script 'baixar_mapa.py' primeiro.")
    else:
        # Lê o arquivo de mapa complexo com geopandas
        print(f"Lendo dados de '{caminho_mapa_gpkg}'...")
        gdf = geopandas.read_file(caminho_mapa_gpkg)
        
        # Cria um dicionário mais simples, contendo apenas o essencial: sigla e geometria
        dados_mapa = {row['abbrev_state']: row['geometry'] for index, row in gdf.iterrows()}
        
        # Salva o dicionário em um arquivo .pkl (formato simples)
        with open(caminho_mapa_pkl, 'wb') as f:
            pickle.dump(dados_mapa, f)
            
        print(f"\nSUCESSO! O mapa foi convertido e salvo como '{caminho_mapa_pkl}'.")
        print("Este arquivo agora será usado pelo programa principal.")

except Exception as e:
    print(f"\nOcorreu um erro durante a conversão: {e}")