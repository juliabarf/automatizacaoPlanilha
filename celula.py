import pandas as pd

colunas = {
    'Profundidade': self._profundidade,
    'Porosity (%)': self.porosidade(),
    'Porosity Decimal': self.porosidadeDec(),
    'Permeability (mD)': self._permeabilidade,
    'RQI': self.rqi(),
    'PHI(Z)': self.phi(),
    'FZI': self.fzi()
}
dfColunas = pd.DataFrame(colunas)

writer = pd.ExcelWriter('tabela.xlsx', engine='xlsxwriter')
dfColunas.to_excel(writer, sheet_name='Planilha1', index=False)

workbook = writer.book
worksheet = writer.sheets['Planilha1']

# Definindo o formato com centralização horizontal e vertical
cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})

# Aplicando o formato às células preenchidas
for col_num, value in enumerate(dfColunas.columns.values):
    # A primeira linha (0) é o cabeçalho
    worksheet.write(0, col_num, value, cell_format)
    for row in range(1, len(dfColunas) + 1):
        worksheet.write(row, col_num, dfColunas.iloc[row-1, col_num], cell_format)

# Definindo largura das colunas
worksheet.set_column('A:A', 20)
worksheet.set_column('B:B', 20)
worksheet.set_column('C:C', 20)
worksheet.set_column('D:D', 20)
worksheet.set_column('E:E', 15)
worksheet.set_column('F:F', 15)
worksheet.set_column('G:G', 15)

writer.close()
