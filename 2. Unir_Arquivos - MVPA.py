import os
import pandas as pd

pasta = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.2. Bases Tratadas'  # Substitua pelo caminho da pasta desejada
dados_concatenados = pd.DataFrame()
# Loop para percorrer os arquivos na pasta
for arquivo in os.listdir(pasta):
    caminho_arquivo = os.path.join(pasta, arquivo)
    if os.path.isfile(caminho_arquivo) and arquivo.endswith('.xlsx'):
        dados_planilha = pd.read_excel(caminho_arquivo)
        dados_concatenados = pd.concat([dados_concatenados, dados_planilha])

caminho_destino = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.3. Consolidado das bases\consolidado.xlsx'  # Substitua pelo caminho desejado
# Salvar os dados concatenados em uma única planilha no arquivo Excel no caminho especificado
dados_concatenados.to_excel(caminho_destino, sheet_name='Dados Consolidados', index=False)
