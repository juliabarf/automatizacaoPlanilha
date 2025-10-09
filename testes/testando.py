file_path = self.nomeTabela + 'Alterada.xlsx'

with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
    dfColunas.to_excel(writer, sheet_name='Planilha1', index=False)

    workbook = writer.book
    worksheet = writer.sheets['Planilha1']

    cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
    decimal_format = workbook.add_format({'num_format': '0.000', 'align': 'center', 'valign': 'vcenter'})
    decimal_format3 = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
    float_format = workbook.add_format({'num_format': '0.############', 'align': 'center', 'valign': 'vcenter'})

    colunas_com_decimal = ['Porosity Decimal', 'Profundidade', 'Permeability (mD)', 'Porosity (%)', 'PHI(Z)']
    coluna3_dec = ['RQI', 'FZI']

    for col_num, value in enumerate(dfColunas.columns.values):
        worksheet.write(0, col_num, value, cell_format)
        for row in range(1, len(dfColunas) + 1):
            valor = dfColunas.iloc[row - 1, col_num]
            if value in colunas_com_decimal:
                worksheet.write(row, col_num, valor, decimal_format)
            elif value in coluna3_dec:
                worksheet.write(row, col_num, valor, decimal_format3)
            else:
                worksheet.write(row, col_num, valor, cell_format)

    for col_num, value in enumerate(dfColunas.columns.values):
        if value in ['FZI', 'RQI', 'PHI(Z)', 'GHE']:
            worksheet.write(0, col_num, value, workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFFF99'}))
        elif value in ['Profundidade', 'Porosity (%)', 'Porosity Decimal', 'Permeability (mD)']:
            worksheet.write(0, col_num, value, workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFCCCC'}))
        else:
            worksheet.write(0, col_num, value, cell_format)

    worksheet.set_column('A:G', 20)
