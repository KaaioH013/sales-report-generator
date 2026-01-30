# baixar_mapa.py
import geobr
import os
import shutil

# Define o nome da pasta para os dados geográficos
data_folder = 'geo_data'

# Se a pasta já existir, apague-a para garantir dados limpos
if os.path.exists(data_folder):
    print(f"Removendo a pasta '{data_folder}' antiga...")
    shutil.rmtree(data_folder)

# Cria a pasta
os.makedirs(data_folder)
print(f"Pasta '{data_folder}' criada com sucesso.")

try:
    print("\nBaixando o arquivo de mapa dos estados do Brasil...")
    print("(Isso pode levar um momento e requer conexão com a internet)")
    
    # Baixa os dados dos estados
    gdf_brasil = geobr.read_state(year=2020, simplified=True)
    
    # Define o caminho completo do arquivo a ser salvo
    file_path = os.path.join(data_folder, 'br_states.gpkg')
    
    # Salva os dados em um arquivo no formato GeoPackage
    gdf_brasil.to_file(file_path, driver='GPKG')

    print(f"\nSUCESSO! O mapa foi baixado e salvo em: '{file_path}'")
    print("Você só precisa executar este script uma vez.")

except Exception as e:
    print(f"\nERRO: Não foi possível baixar o arquivo do mapa.")
    print(f"Detalhe do erro: {e}")
    print("Verifique sua conexão com a internet e tente novamente.")