import os
from pathlib import Path
import pandas as pd
import shutil

pasta_atual = Path(__file__).parent
pasta_nfce = Path('C:/NFCE')
pasta_destino = Path('C:/NFCE/Arquivos_XML') 
planilha_coo = pd.read_excel(pasta_atual / 'planilha_COO.xlsx', dtype=str) 
coos = planilha_coo['COO'].tolist()
serie = "789" 

def formatar_numero():
    coos_formatado = [coo.zfill(9) for coo in coos]
    serie_formatada = serie.zfill(3)
    arquivo_procurado = list()
    for coo in coos_formatado:
        coo_serie = serie_formatada + coo 
        arquivo_procurado.append(coo_serie)
    return arquivo_procurado

arquivos_procurados = formatar_numero()

for arquivo_procurado in arquivos_procurados:
    for arquivo in pasta_nfce.glob('**/*'):
        if arquivo.is_file and arquivo_procurado in arquivo.name:
            if pasta_destino.exists():
                shutil.copy(arquivo, pasta_destino)
                print(f'Arquivo {arquivo} adicionado')
            else:
                pasta_destino.mkdir(parents=True, exist_ok=True)   
                shutil.copy(arquivo, pasta_destino)
        

        

