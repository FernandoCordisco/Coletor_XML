import os
from pathlib import Path
import pandas as pd
import shutil
import sys

if hasattr(sys, '_MEIPASS'):
    pasta_atual = Path(sys._MEIPASS)
else:
    pasta_atual = Path(__file__).parent
    
pasta_atual = Path(__file__).parent
pasta_nfce = Path('C:/NFCE')
pasta_destino = Path('C:/NFCE/Arquivos_XML') 
planilha= pd.read_excel(pasta_atual / 'planilha_COO.xlsx', dtype=str) 
coos = planilha['COO'].tolist()
serie = planilha['Serie'].iloc[0]

def formatar_numero():
    coos_formatado = [coo.zfill(9) for coo in coos]
    serie_formatada = serie.zfill(3)
    arquivo_procurado = list()
    for coo in coos_formatado:
        coo_serie = serie_formatada + coo 
        arquivo_procurado.append(coo_serie)
    return arquivo_procurado

def copiar_arquivos():
    for arquivo in pasta_nfce.glob('**/*'):
        if arquivo.is_file():
            if pasta_destino in arquivo.parents:
                continue
            for procurado in arquivos_procurados_set:
                if procurado in arquivo.name:
                    subpasta = arquivo.parent.name
                    caminho_destino = pasta_destino / subpasta / arquivo.name
                    caminho_destino.parent.mkdir(parents=True, exist_ok=True)
                    if caminho_destino.exists():
                        print(f'Arquivo {arquivo} já existe na pasta')
                    else:
                        shutil.copy(arquivo, caminho_destino)
                        print(f'Arquivo {arquivo} adicionado')
                    break

arquivos_procurados = formatar_numero()
arquivos_procurados_set = set(arquivos_procurados)
copiar_arquivos()
            
        

        

