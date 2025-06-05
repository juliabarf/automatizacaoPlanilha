import pandas as pd
import xlsxwriter

# Criar um DataFrame de exemplo
data = {'Coluna 1': ['Valor 1', 'Valor 2', 'Valor 3'],
        'Coluna 2': [1, 2, 3],
        'Coluna 3': [4, 5, 6],
        'Coluna 4': [7, 8, 9],}
df = pd.DataFrame(data)

# Cria um ExcelWriter
writer = pd.ExcelWriter('saida.xlsx', engine='xlsxwriter')

# Abre uma planilha
df.to_excel(writer, sheet_name='Planilha1', index=False)

# Obtem o workbook e a planilha
workbook = writer.book
worksheet = writer.sheets['Planilha1']

# Ajustar a largura das colunas
# A largura é especificada em pixels
worksheet.set_column('A:A', 20) # Coluna A com 20 pixels de largura
worksheet.set_column('B:B', 10) # Coluna B com 10 pixels de largura

worksheet.merged_cells('C5:D5')
# Assume que 'df' é o seu DataFrame e 'coluna_a_converter' é a coluna que você quer converter

# Fechar o ExcelWriter
writer.close()