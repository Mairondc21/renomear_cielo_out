import pandas as pd
from pathlib import Path
import os

def process_files(file_base_path, file_nets_path, output_path):
    # Carregar os arquivos em dataframes
    arq_base = pd.read_excel(file_base_path, sheet_name='bd')
    arq_nets = pd.read_excel(file_nets_path, sheet_name='pBI')
    
    # Definir as colunas para verificar e renomear
    cols_to_check = ["MOTNEU_1", "MOTNEU_2", "MOTNEU_3", "MOTDET_1", "MOTDET_2", "MOTDET_3", "MOTPRO_1", "MOTPRO_2", "MOTPRO_3"]

    # Iterar sobre as linhas do 'arq_nets' para verificar e substituir valores em 'arq_base'
    for _, row in arq_nets.iterrows():
        label = row["Label original"]
        net_value = row["Net"]

        # Verificar e renomear valores correspondentes em 'arq_base'
        for col in cols_to_check:
            arq_base[col] = arq_base[col].replace(label, net_value)
    
    # Salvar o dataframe modificado em um novo arquivo Excel
    arq_base.to_excel(output_path, index=False)

def renomear_audios(base: pd.ExcelFile, caminho_audios: Path):
    df_base = pd.read_excel(base, sheet_name="FINAL")
    arquivos = os.listdir(caminho_audios)

    for arquivo in arquivos:
        if arquivo.endswith(".mp3"):
            caminho_subarquivos = os.path.join(caminho_audios,arquivo)
            nome_completo = arquivo.split("_")[1]
            nome = os.path.splitext(nome_completo)[0]
        else:
            continue    
        try:
            for index, row in df_base.astype(str).iterrows():
                # Verificar se o nome está em uma das três colunas
                if nome == row['ID_ONDA']:
                    
                    novo_nome = f"{row['Atributo']}\{row['SEG']}\{row['FINAL']}"
                    novo_caminho = os.path.join(caminho_audios,novo_nome)

                    os.rename(caminho_subarquivos,novo_caminho)            
        except Exception as e:
            print(f"Unexpected {e=}, {type(e)=}")


bd_8 = "./dados/bd_8.xlsx"
nets = "./dados/nets.xlsx"
caminho_audios = r'C:\Users\mairon.costa\OneDrive - Expertise Inteligência e Pesquisa de Mercado\expertise_mairon\2024\Cielo_NPS_cortes\outubro\audios'

renomear_audios(bd_8,caminho_audios)
