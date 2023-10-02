import subprocess
import time

def executar_script(caminho_arquivo, tempo_espera):
    subprocess.run(['python', caminho_arquivo])
    time.sleep(tempo_espera)

# Lista de caminhos dos arquivos Python a serem executados
caminhos_arquivos = [
    r'C:\Users\Victor\PycharmProjects\Unir os arquivos\1. Carregar_Arquivos.py',
    r'C:\Users\Victor\PycharmProjects\Unir os arquivos\2. Unir_Arquivos.py',
    r'C:\Users\Victor\PycharmProjects\Unir os arquivos\3. Carregar_Arquivos_BD.py'
]

# Tempo de espera entre a execução de cada script
tempo_espera = 10

# Executa cada arquivo Python e aguarda o tempo de espera
for caminho_arquivo in caminhos_arquivos:
    executar_script(caminho_arquivo, tempo_espera)