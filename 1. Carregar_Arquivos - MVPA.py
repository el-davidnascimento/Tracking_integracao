import pandas as pd
import datetime as date

########################################################################### PANAMERICANO ######################################################
# 03. Checklist Integração - Horário Integral

caminho_arquivo = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.1. Bases Cruas\03. Checklist Integração - Horário Integral.xlsx'
nome_aba = 'FORMULÁRIO'
dados_excel = pd.read_excel(caminho_arquivo, header=2, sheet_name=nome_aba)

# Obter a data atual
data_atual = date.datetime.today()


# Criar uma nova coluna com a data atual
dados_excel['DATA'] = data_atual
dados_excel['SETOR'] = 'Turno Integral'
dados_excel['AREA'] = 'OP. Escola'
dados_excel['UNIDADE'] = 'MVPA'

# Seleciona as colunas desejadas
colunas_fixas = ['FINANCEIRO', 'PEDAGÓGICO', 'RH', 'JURÍDICO', 'MARKETING', 'ATENDIMENTO', 'DADOS', 'OPERACIONAL', 'TI', 'EXPANSÃO', 'M&A', 'FORNECEDOR']  # Insira as letras das colunas desejadas aqui

# Substitui os valores "X" por "SUBSTITUIR" nas colunas fixas
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('X', 'SIM').fillna('NÃO')
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('x', 'SIM').fillna('NÃO')

# Exclui as linhas com valores em branco na segunda coluna
dados_excel = dados_excel.dropna(subset=[dados_excel.columns[1]])


colunas = dados_excel.columns
caminho_destino = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.2. Bases Tratadas\Colunas_Integral.xlsx'
dados_excel.to_excel(caminho_destino,index=False)



# 03. Checklist Integração - Setor Pessoal

caminho_arquivo = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.1. Bases Cruas\03. Checklist Integração - Setor Pessoal.xlsx'
nome_aba = 'FORMULÁRIO'
dados_excel = pd.read_excel(caminho_arquivo, header=2, sheet_name=nome_aba)

# Obter a data atual
data_atual = date.datetime.today()


# Criar uma nova coluna com a data atual
dados_excel['DATA'] = data_atual
dados_excel['SETOR'] = 'Setor Pessoal'
dados_excel['AREA'] = 'OP. Escola'
dados_excel['UNIDADE'] = 'MVPA'

# Seleciona as colunas desejadas
colunas_fixas = ['FINANCEIRO', 'PEDAGÓGICO', 'RH', 'JURÍDICO', 'MARKETING', 'ATENDIMENTO', 'DADOS', 'OPERACIONAL', 'TI', 'EXPANSÃO', 'M&A', 'FORNECEDOR']  # Insira as letras das colunas desejadas aqui

# Substitui os valores "X" por "SUBSTITUIR" nas colunas fixas
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('X', 'SIM').fillna('NÃO')
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('x', 'SIM').fillna('NÃO')

# Exclui as linhas com valores em branco na segunda coluna
dados_excel = dados_excel.dropna(subset=[dados_excel.columns[1]])


colunas = dados_excel.columns
caminho_destino = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.2. Bases Tratadas\Colunas_Gente.xlsx'
dados_excel.to_excel(caminho_destino,index=False)


# 03. Checklist Integração - Almoxarifado

caminho_arquivo = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.1. Bases Cruas\03. Checklist Integração - Almoxarifado.xlsx'
nome_aba = 'FORMULÁRIO'
dados_excel = pd.read_excel(caminho_arquivo, header=2, sheet_name=nome_aba)

# Obter a data atual
data_atual = date.datetime.today()


# Criar uma nova coluna com a data atual
dados_excel['DATA'] = data_atual
dados_excel['SETOR'] = 'Almoxarifado'
dados_excel['AREA'] = 'OP. Escola'
dados_excel['UNIDADE'] = 'MVPA'

# Seleciona as colunas desejadas
colunas_fixas = ['FINANCEIRO', 'PEDAGÓGICO', 'RH', 'JURÍDICO', 'MARKETING', 'ATENDIMENTO', 'DADOS', 'OPERACIONAL', 'TI', 'EXPANSÃO', 'M&A', 'FORNECEDOR']  # Insira as letras das colunas desejadas aqui

# Substitui os valores "X" por "SUBSTITUIR" nas colunas fixas
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('X', 'SIM').fillna('NÃO')
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('x', 'SIM').fillna('NÃO')

# Exclui as linhas com valores em branco na segunda coluna
dados_excel = dados_excel.dropna(subset=[dados_excel.columns[1]])


colunas = dados_excel.columns
caminho_destino = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.2. Bases Tratadas\Colunas_almoxarifado.xlsx'
dados_excel.to_excel(caminho_destino,index=False)


# 03. Checklist Integração - Cozinha e Cantina

caminho_arquivo = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.1. Bases Cruas\03. Checklist Integração - Cozinha e Cantina.xlsx'
nome_aba = 'FORMULÁRIO'
dados_excel = pd.read_excel(caminho_arquivo, header=2, sheet_name=nome_aba)

# Obter a data atual
data_atual = date.datetime.today()


# Criar uma nova coluna com a data atual
dados_excel['DATA'] = data_atual
dados_excel['SETOR'] = 'Cozinha e Cantina'
dados_excel['AREA'] = 'OP. Escola'
dados_excel['UNIDADE'] = 'MVPA'

# Seleciona as colunas desejadas
colunas_fixas = ['FINANCEIRO', 'PEDAGÓGICO', 'RH', 'JURÍDICO', 'MARKETING', 'ATENDIMENTO', 'DADOS', 'OPERACIONAL', 'TI', 'EXPANSÃO', 'M&A', 'FORNECEDOR']  # Insira as letras das colunas desejadas aqui

# Substitui os valores "X" por "SUBSTITUIR" nas colunas fixas
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('X', 'SIM').fillna('NÃO')
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('x', 'SIM').fillna('NÃO')

# Exclui as linhas com valores em branco na segunda coluna
dados_excel = dados_excel.dropna(subset=[dados_excel.columns[1]])


colunas = dados_excel.columns
caminho_destino = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.2. Bases Tratadas\Colunas_cozinha.xlsx'
dados_excel.to_excel(caminho_destino,index=False)


# 03. Checklist Integração - Dados

caminho_arquivo = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.1. Bases Cruas\03. Checklist Integração - Dados.xlsx'
nome_aba = 'FORMULÁRIO'
dados_excel = pd.read_excel(caminho_arquivo, header=2, sheet_name=nome_aba)

# Obter a data atual
data_atual = date.datetime.today()


# Criar uma nova coluna com a data atual
dados_excel['DATA'] = data_atual
dados_excel['SETOR'] = 'Dados'
dados_excel['AREA'] = 'CSC'
dados_excel['UNIDADE'] = 'MVPA'

# Seleciona as colunas desejadas
colunas_fixas = ['FINANCEIRO', 'PEDAGÓGICO', 'RH', 'JURÍDICO', 'MARKETING', 'ATENDIMENTO', 'DADOS', 'OPERACIONAL', 'TI', 'EXPANSÃO', 'M&A', 'FORNECEDOR']  # Insira as letras das colunas desejadas aqui

# Substitui os valores "X" por "SUBSTITUIR" nas colunas fixas
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('X', 'SIM').fillna('NÃO')
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('x', 'SIM').fillna('NÃO')

# Exclui as linhas com valores em branco na segunda coluna
dados_excel = dados_excel.dropna(subset=[dados_excel.columns[1]])


colunas = dados_excel.columns
caminho_destino = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.2. Bases Tratadas\Colunas_dados.xlsx'
dados_excel.to_excel(caminho_destino,index=False)


# 03. Checklist Integração - Financeiro

caminho_arquivo = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.1. Bases Cruas\03. Checklist Integração - Financeiro.xlsx'
nome_aba = 'FORMULÁRIO'
dados_excel = pd.read_excel(caminho_arquivo, header=2, sheet_name=nome_aba)

# Obter a data atual
data_atual = date.datetime.today()


# Criar uma nova coluna com a data atual
dados_excel['DATA'] = data_atual
dados_excel['SETOR'] = dados_excel['Macro']
dados_excel['AREA'] = 'CSC'
dados_excel['UNIDADE'] = 'MVPA'

# Seleciona as colunas desejadas
colunas_fixas = ['FINANCEIRO', 'PEDAGÓGICO', 'RH', 'JURÍDICO', 'MARKETING', 'ATENDIMENTO', 'DADOS', 'OPERACIONAL', 'TI', 'EXPANSÃO', 'M&A', 'FORNECEDOR']  # Insira as letras das colunas desejadas aqui

# Substitui os valores "X" por "SUBSTITUIR" nas colunas fixas
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('X', 'SIM').fillna('NÃO')
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('x', 'SIM').fillna('NÃO')

# Exclui as linhas com valores em branco na segunda coluna
dados_excel = dados_excel.dropna(subset=[dados_excel.columns[1]])


colunas = dados_excel.columns
caminho_destino = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.2. Bases Tratadas\Colunas_financeiro.xlsx'
dados_excel.to_excel(caminho_destino,index=False)


# 03. Checklist Integração - Infraestrutura

caminho_arquivo = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.1. Bases Cruas\03. Checklist Integração - Infraestrutura.xlsx'
nome_aba = 'FORMULÁRIO'
dados_excel = pd.read_excel(caminho_arquivo, header=2, sheet_name=nome_aba)

# Obter a data atual
data_atual = date.datetime.today()


# Criar uma nova coluna com a data atual
dados_excel['DATA'] = data_atual
dados_excel['SETOR'] = 'Infraestrutura'
dados_excel['AREA'] = 'CSC'
dados_excel['UNIDADE'] = 'MVPA'

# Seleciona as colunas desejadas
colunas_fixas = ['FINANCEIRO', 'PEDAGÓGICO', 'RH', 'JURÍDICO', 'MARKETING', 'ATENDIMENTO', 'DADOS', 'OPERACIONAL', 'TI', 'EXPANSÃO', 'M&A', 'FORNECEDOR']  # Insira as letras das colunas desejadas aqui

# Substitui os valores "X" por "SUBSTITUIR" nas colunas fixas
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('X', 'SIM').fillna('NÃO')
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('x', 'SIM').fillna('NÃO')

# Exclui as linhas com valores em branco na segunda coluna
dados_excel = dados_excel.dropna(subset=[dados_excel.columns[1]])


colunas = dados_excel.columns
caminho_destino = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.2. Bases Tratadas\Colunas_Infraestrutura.xlsx'
dados_excel.to_excel(caminho_destino,index=False)


# 03. Checklist Integração - Marketing

caminho_arquivo = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.1. Bases Cruas\03. Checklist Integração - Marketing.xlsx'
nome_aba = 'FORMULÁRIO'
dados_excel = pd.read_excel(caminho_arquivo, header=2, sheet_name=nome_aba)

# Obter a data atual
data_atual = date.datetime.today()


# Criar uma nova coluna com a data atual
dados_excel['DATA'] = data_atual
dados_excel['SETOR'] = 'Marketing'
dados_excel['AREA'] = 'CSC'
dados_excel['UNIDADE'] = 'MVPA'

# Seleciona as colunas desejadas
colunas_fixas = ['FINANCEIRO', 'PEDAGÓGICO', 'RH', 'JURÍDICO', 'MARKETING', 'ATENDIMENTO', 'DADOS', 'OPERACIONAL', 'TI', 'EXPANSÃO', 'M&A', 'FORNECEDOR']  # Insira as letras das colunas desejadas aqui

# Substitui os valores "X" por "SUBSTITUIR" nas colunas fixas
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('X', 'SIM').fillna('NÃO')
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('x', 'SIM').fillna('NÃO')

# Exclui as linhas com valores em branco na segunda coluna
dados_excel = dados_excel.dropna(subset=[dados_excel.columns[1]])


colunas = dados_excel.columns
caminho_destino = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.2. Bases Tratadas\Colunas_Marketing.xlsx'
dados_excel.to_excel(caminho_destino,index=False)


# 03. Checklist Integração - Pedagógico

caminho_arquivo = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.1. Bases Cruas\03. Checklist Integração - Pedagógico.xlsx'
nome_aba = 'FORMULÁRIO'
dados_excel = pd.read_excel(caminho_arquivo, header=2, sheet_name=nome_aba)

# Obter a data atual
data_atual = date.datetime.today()


# Criar uma nova coluna com a data atual
dados_excel['DATA'] = data_atual
dados_excel['SETOR'] = 'Pedagógico'
dados_excel['AREA'] = 'OP. Escola'
dados_excel['UNIDADE'] = 'MVPA'

# Seleciona as colunas desejadas
colunas_fixas = ['FINANCEIRO', 'PEDAGÓGICO', 'RH', 'JURÍDICO', 'MARKETING', 'ATENDIMENTO', 'DADOS', 'OPERACIONAL', 'TI', 'EXPANSÃO', 'M&A', 'FORNECEDOR']  # Insira as letras das colunas desejadas aqui

# Substitui os valores "X" por "SUBSTITUIR" nas colunas fixas
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('X', 'SIM').fillna('NÃO')
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('x', 'SIM').fillna('NÃO')

# Exclui as linhas com valores em branco na segunda coluna
dados_excel = dados_excel.dropna(subset=[dados_excel.columns[1]])


colunas = dados_excel.columns
caminho_destino = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.2. Bases Tratadas\Colunas_Pedagógico.xlsx'
dados_excel.to_excel(caminho_destino,index=False)



# 03. Checklist Integração - Portaria

caminho_arquivo = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.1. Bases Cruas\03. Checklist Integração - Portaria.xlsx'
nome_aba = 'FORMULÁRIO'
dados_excel = pd.read_excel(caminho_arquivo, header=2, sheet_name=nome_aba)

# Obter a data atual
data_atual = date.datetime.today()

# Criar uma nova coluna com a data atual
dados_excel['DATA'] = data_atual
dados_excel['SETOR'] = 'Portaria'
dados_excel['AREA'] = 'OP. Escola'
dados_excel['UNIDADE'] = 'MVPA'

# Seleciona as colunas desejadas
colunas_fixas = ['FINANCEIRO', 'PEDAGÓGICO', 'RH', 'JURÍDICO', 'MARKETING', 'ATENDIMENTO', 'DADOS', 'OPERACIONAL', 'TI', 'EXPANSÃO', 'M&A', 'FORNECEDOR']  # Insira as letras das colunas desejadas aqui

# Substitui os valores "X" por "SUBSTITUIR" nas colunas fixas
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('X', 'SIM').fillna('NÃO')
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('x', 'SIM').fillna('NÃO')

# Exclui as linhas com valores em branco na segunda coluna
dados_excel = dados_excel.dropna(subset=[dados_excel.columns[1]])


colunas = dados_excel.columns
caminho_destino = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.2. Bases Tratadas\Colunas_Portaria.xlsx'
dados_excel.to_excel(caminho_destino,index=False)


# 03. Checklist Integração - Ti

caminho_arquivo = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.1. Bases Cruas\03. Checklist Integração - Ti.xlsx'
nome_aba = 'FORMULÁRIO'
dados_excel = pd.read_excel(caminho_arquivo, header=2, sheet_name=nome_aba)

# Obter a data atual
data_atual = date.datetime.today()


# Criar uma nova coluna com a data atual
dados_excel['DATA'] = data_atual
dados_excel['SETOR'] = 'TI'
dados_excel['AREA'] = 'CSC'
dados_excel['UNIDADE'] = 'MVPA'

# Seleciona as colunas desejadas
colunas_fixas = ['FINANCEIRO', 'PEDAGÓGICO', 'RH', 'JURÍDICO', 'MARKETING', 'ATENDIMENTO', 'DADOS', 'OPERACIONAL', 'TI', 'EXPANSÃO', 'M&A', 'FORNECEDOR']  # Insira as letras das colunas desejadas aqui

# Substitui os valores "X" por "SUBSTITUIR" nas colunas fixas
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('X', 'SIM').fillna('NÃO')
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('x', 'SIM').fillna('NÃO')

# Exclui as linhas com valores em branco na segunda coluna
dados_excel = dados_excel.dropna(subset=[dados_excel.columns[1]])


colunas = dados_excel.columns
caminho_destino = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.2. Bases Tratadas\Colunas_Ti.xlsx'
dados_excel.to_excel(caminho_destino,index=False)


# 03. Checklist Integração Facilities

caminho_arquivo = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.1. Bases Cruas\03. Checklist Integração Facilities.xlsx'
nome_aba = 'FORMULÁRIO'
dados_excel = pd.read_excel(caminho_arquivo, header=2, sheet_name=nome_aba)

# Obter a data atual
data_atual = date.datetime.today()


# Criar uma nova coluna com a data atual
dados_excel['DATA'] = data_atual
dados_excel['SETOR'] = 'Facilities'
dados_excel['AREA'] = 'OP. Escola'
dados_excel['UNIDADE'] = 'MVPA'

# Seleciona as colunas desejadas
colunas_fixas = ['FINANCEIRO', 'PEDAGÓGICO', 'RH', 'JURÍDICO', 'MARKETING', 'ATENDIMENTO', 'DADOS', 'OPERACIONAL', 'TI', 'EXPANSÃO', 'M&A', 'FORNECEDOR']  # Insira as letras das colunas desejadas aqui

# Substitui os valores "X" por "SUBSTITUIR" nas colunas fixas
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('X', 'SIM').fillna('NÃO')
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('x', 'SIM').fillna('NÃO')

# Exclui as linhas com valores em branco na segunda coluna
dados_excel = dados_excel.dropna(subset=[dados_excel.columns[1]])


colunas = dados_excel.columns
caminho_destino = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.2. Bases Tratadas\Colunas_Facilities.xlsx'
dados_excel.to_excel(caminho_destino,index=False)



# 03. Checklist Integração Lojinha

caminho_arquivo = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.1. Bases Cruas\03. Checklist Integração Lojinha.xlsx'
nome_aba = 'FORMULÁRIO'
dados_excel = pd.read_excel(caminho_arquivo, header=2, sheet_name=nome_aba)

# Obter a data atual
data_atual = date.datetime.today()


# Criar uma nova coluna com a data atual
dados_excel['DATA'] = data_atual
dados_excel['SETOR'] = 'Lojinha'
dados_excel['AREA'] = 'OP. Escola'
dados_excel['UNIDADE'] = 'MVPA'

# Seleciona as colunas desejadas
colunas_fixas = ['FINANCEIRO', 'PEDAGÓGICO', 'RH', 'JURÍDICO', 'MARKETING', 'ATENDIMENTO', 'DADOS', 'OPERACIONAL', 'TI', 'EXPANSÃO', 'M&A', 'FORNECEDOR']  # Insira as letras das colunas desejadas aqui

# Substitui os valores "X" por "SUBSTITUIR" nas colunas fixas
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('X', 'SIM').fillna('NÃO')
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('x', 'SIM').fillna('NÃO')

# Exclui as linhas com valores em branco na segunda coluna
dados_excel = dados_excel.dropna(subset=[dados_excel.columns[1]])


colunas = dados_excel.columns
caminho_destino = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.2. Bases Tratadas\Colunas_Lojinha.xlsx'
dados_excel.to_excel(caminho_destino,index=False)


# 03. Checklist Integração Operações

caminho_arquivo = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.1. Bases Cruas\03. Checklist Integração Operações.xlsx'
nome_aba = 'FORMULÁRIO'
dados_excel = pd.read_excel(caminho_arquivo, header=2, sheet_name=nome_aba)

# Obter a data atual
data_atual = date.datetime.today()

# Criar uma nova coluna com a data atual
dados_excel['DATA'] = data_atual
dados_excel['SETOR'] = 'Operações'
dados_excel['AREA'] = 'OP. Escola'
dados_excel['UNIDADE'] = 'MVPA'

# Seleciona as colunas desejadas
colunas_fixas = ['FINANCEIRO', 'PEDAGÓGICO', 'RH', 'JURÍDICO', 'MARKETING', 'ATENDIMENTO', 'DADOS', 'OPERACIONAL', 'TI', 'EXPANSÃO', 'M&A', 'FORNECEDOR']  # Insira as letras das colunas desejadas aqui

# Substitui os valores "X" por "SUBSTITUIR" nas colunas fixas
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('X', 'SIM').fillna('NÃO')
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('x', 'SIM').fillna('NÃO')

# Exclui as linhas com valores em branco na segunda coluna
dados_excel = dados_excel.dropna(subset=[dados_excel.columns[1]])


colunas = dados_excel.columns
caminho_destino = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.2. Bases Tratadas\Colunas_Operações.xlsx'
dados_excel.to_excel(caminho_destino,index=False)


# 03. Checklist Integração Recepção

caminho_arquivo = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.1. Bases Cruas\03. Checklist Integração Recepção.xlsx'
nome_aba = 'FORMULÁRIO'
dados_excel = pd.read_excel(caminho_arquivo, header=2, sheet_name=nome_aba)

# Obter a data atual
data_atual = date.datetime.today()


# Criar uma nova coluna com a data atual
dados_excel['DATA'] = data_atual
dados_excel['SETOR'] = 'Recepção'
dados_excel['AREA'] = 'OP. Escola'
dados_excel['UNIDADE'] = 'MVPA'

# Seleciona as colunas desejadas
colunas_fixas = ['FINANCEIRO', 'PEDAGÓGICO', 'RH', 'JURÍDICO', 'MARKETING', 'ATENDIMENTO', 'DADOS', 'OPERACIONAL', 'TI', 'EXPANSÃO', 'M&A', 'FORNECEDOR']  # Insira as letras das colunas desejadas aqui

# Substitui os valores "X" por "SUBSTITUIR" nas colunas fixas
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('X', 'SIM').fillna('NÃO')
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('x', 'SIM').fillna('NÃO')

# Exclui as linhas com valores em branco na segunda coluna
dados_excel = dados_excel.dropna(subset=[dados_excel.columns[1]])


colunas = dados_excel.columns
caminho_destino = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.2. Bases Tratadas\Colunas_Recepção.xlsx'
dados_excel.to_excel(caminho_destino,index=False)


# 03. Checklist Integração Regulatório

caminho_arquivo = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.1. Bases Cruas\03. Checklist Integração Regulatório.xlsx'
nome_aba = 'FORMULÁRIO'
dados_excel = pd.read_excel(caminho_arquivo, header=2, sheet_name=nome_aba)

# Obter a data atual
data_atual = date.datetime.today()


# Criar uma nova coluna com a data atual
dados_excel['DATA'] = data_atual
dados_excel['SETOR'] = 'Regulatório'
dados_excel['AREA'] = 'OP. Escola'
dados_excel['UNIDADE'] = 'MVPA'

# Seleciona as colunas desejadas
colunas_fixas = ['FINANCEIRO', 'PEDAGÓGICO', 'RH', 'JURÍDICO', 'MARKETING', 'ATENDIMENTO', 'DADOS', 'OPERACIONAL', 'TI', 'EXPANSÃO', 'M&A', 'FORNECEDOR']  # Insira as letras das colunas desejadas aqui

# Substitui os valores "X" por "SUBSTITUIR" nas colunas fixas
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('X', 'SIM').fillna('NÃO')
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('x', 'SIM').fillna('NÃO')

# Exclui as linhas com valores em branco na segunda coluna
dados_excel = dados_excel.dropna(subset=[dados_excel.columns[1]])


colunas = dados_excel.columns
caminho_destino = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.2. Bases Tratadas\Colunas_Regulatorio.xlsx'
dados_excel.to_excel(caminho_destino,index=False)


# 03. Checklist Integração Secretaria

caminho_arquivo = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.1. Bases Cruas\03. Checklist Integração Secretaria.xlsx'
nome_aba = 'FORMULÁRIO'
dados_excel = pd.read_excel(caminho_arquivo, header=2, sheet_name=nome_aba)

# Obter a data atual
data_atual = date.datetime.today()


# Criar uma nova coluna com a data atual
dados_excel['DATA'] = data_atual
dados_excel['SETOR'] = 'Secretaria'
dados_excel['AREA'] = 'OP. Escola'
dados_excel['UNIDADE'] = 'MVPA'

# Seleciona as colunas desejadas
colunas_fixas = ['FINANCEIRO', 'PEDAGÓGICO', 'RH', 'JURÍDICO', 'MARKETING', 'ATENDIMENTO', 'DADOS', 'OPERACIONAL', 'TI', 'EXPANSÃO', 'M&A', 'FORNECEDOR']  # Insira as letras das colunas desejadas aqui

# Substitui os valores "X" por "SUBSTITUIR" nas colunas fixas
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('X', 'SIM').fillna('NÃO')
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('x', 'SIM').fillna('NÃO')

# Exclui as linhas com valores em branco na segunda coluna
dados_excel = dados_excel.dropna(subset=[dados_excel.columns[1]])


colunas = dados_excel.columns
caminho_destino = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.2. Bases Tratadas\Colunas_Secretaria.xlsx'
dados_excel.to_excel(caminho_destino,index=False)


# 03. Checklist Integração Serviços Gerais

caminho_arquivo = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.1. Bases Cruas\03. Checklist Integração Serviços Gerais.xlsx'
nome_aba = 'FORMULÁRIO'
dados_excel = pd.read_excel(caminho_arquivo, header=2, sheet_name=nome_aba)

# Obter a data atual
data_atual = date.datetime.today()


# Criar uma nova coluna com a data atual
dados_excel['DATA'] = data_atual
dados_excel['SETOR'] = 'Serviços Gerais'
dados_excel['AREA'] = 'OP. Escola'
dados_excel['UNIDADE'] = 'MVPA'

# Seleciona as colunas desejadas
colunas_fixas = ['FINANCEIRO', 'PEDAGÓGICO', 'RH', 'JURÍDICO', 'MARKETING', 'ATENDIMENTO', 'DADOS', 'OPERACIONAL', 'TI', 'EXPANSÃO', 'M&A', 'FORNECEDOR']  # Insira as letras das colunas desejadas aqui

# Substitui os valores "X" por "SUBSTITUIR" nas colunas fixas
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('X', 'SIM').fillna('NÃO')
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('x', 'SIM').fillna('NÃO')

# Exclui as linhas com valores em branco na segunda coluna
dados_excel = dados_excel.dropna(subset=[dados_excel.columns[1]])


colunas = dados_excel.columns
caminho_destino = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.2. Bases Tratadas\Colunas_Serviços.xlsx'
dados_excel.to_excel(caminho_destino,index=False)


# 03. Plano de Integração Atendimento

caminho_arquivo = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.1. Bases Cruas\03. Plano de Integração Atendimento.xlsx'
nome_aba = 'FORMULÁRIO'
dados_excel = pd.read_excel(caminho_arquivo, header=2, sheet_name=nome_aba)

# Obter a data atual
data_atual = date.datetime.today()

# Criar uma nova coluna com a data atual
dados_excel['DATA'] = data_atual
dados_excel['SETOR'] = 'Atendimento'
dados_excel['AREA'] = 'CSC'
dados_excel['UNIDADE'] = 'MVPA'

# Seleciona as colunas desejadas
colunas_fixas = ['FINANCEIRO', 'PEDAGÓGICO', 'RH', 'JURÍDICO', 'MARKETING', 'ATENDIMENTO', 'DADOS', 'OPERACIONAL', 'TI', 'EXPANSÃO', 'M&A', 'FORNECEDOR']  # Insira as letras das colunas desejadas aqui

# Substitui os valores "X" por "SUBSTITUIR" nas colunas fixas
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('X', 'SIM').fillna('NÃO')
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('x', 'SIM').fillna('NÃO')
dados_excel[colunas_fixas] = dados_excel[colunas_fixas].replace('TALLOS/CRM', 'SIM').fillna('NÃO')


# Exclui as linhas com valores em branco na segunda coluna
dados_excel = dados_excel.dropna(subset=[dados_excel.columns[1]])


colunas = dados_excel.columns
caminho_destino = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.3. Integração Panamericano\8.3.3. Bases\8.3.3.2. Bases Tratadas\Colunas_Atendimento.xlsx'
dados_excel.to_excel(caminho_destino,index=False)