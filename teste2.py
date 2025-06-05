import openpyxl

# Carrega a planilha Excel
workbook = openpyxl.load_workbook("tabela.xlsx")

# Define a aba
sheet = workbook["Planilha1"]  # Substitua "Planilha1" pelo nome da sua aba

# Mescla as células A1, B1 e C1

sheet.merge_cells("A1:A2")

# Salva as alterações na planilha
workbook.save("minha_planilha_mesclada.xlsx")