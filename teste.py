import pandas as pd


#ler documentos execel e manipular eles
df = pd.read_excel('fruta.xlsx') #lê o arquivo da planilha
#print(df['cor']) -> lê somente as linhas relacionadas a esta coluna 'cor'. aqui os arquivos vem no estilo de listas

#criando dicionários com os arquivos da planilha
info = {
    'Coluna1': df['coluna1'][0],
    'Coluna2': df['coluna2'][0]
}

print(info['Coluna1'] + info['Coluna2'])

