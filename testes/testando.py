import openpyxl
from openpyxl.chart import BarChart, Reference

# Carregar a planilha existente
wb = openpyxl.load_workbook("testeEXCEL.xlsx")
ws = wb.active  # Ou use wb['Nome_da_Aba'] para selecionar uma aba específica

# Adicionar dados à planilha (opcional)
# Se você já tiver dados, pode pular esta parte
data = [
    ['Mês', 'Vendas'],
    ['Janeiro', 30],
    ['Fevereiro', 45],
    ['Março', 25],
    ['Abril', 50],
]

# Adicionando dados a partir da linha 1, coluna 1
for row in data:
    ws.append(row)

# Criar um gráfico de barras
chart = BarChart()
chart.title = "Vendas Mensais"
chart.x_axis.title = "Mês"
chart.y_axis.title = "Vendas"

# Definir os dados do gráfico
data_reference = Reference(ws, min_col=2, min_row=1, max_col=2, max_row=5)
categories_reference = Reference(ws, min_col=1, min_row=2, max_row=5)

chart.add_data(data_reference, titles_from_data=True)
chart.set_categories(categories_reference)

# Adicionar o gráfico à planilha
ws.add_chart(chart, "D7")  # Posição onde o gráfico será inserido

# Salvar a planilha
wb.save("dados_existentes.xlsx")
