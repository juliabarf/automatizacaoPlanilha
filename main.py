import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import math
from decimal import Decimal, getcontext
from openpyxl import Workbook, load_workbook
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, PatternFill
import numpy as np
import os

# -----------------------------
def permeabilit(phi, fzi):
    phi = np.clip(phi, 1e-6, 0.99)
    phi_e = phi / (1 - phi)
    k = phi * ((fzi * phi_e) / 0.0314) ** 2
    return k

class AutomatizacaoPlanilha:
    def __init__(self, df, nomeTabela):
        self._df = df.copy()
        self.nomeTabela = nomeTabela

        # Limpa nomes de colunas
        self._df.columns = self._df.columns.str.strip()

        convert_dic = {'Prof. (m)': float}
        self._df = self._df.astype(convert_dic)
        self._profundidade = self._df['Prof. (m)']
        self._porosidade = self._df['Porosidade (%)']
        self._permeabilidade = self._df['Perm Abs Longitud (mD)']

    def porosidade(self):
        return list(self._porosidade)

    def porosidadeDec(self):
        return [round(p / 100, 3) for p in self._porosidade]

    getcontext().prec = 28  # define alta precisão
    def porosidadeDec(self):
        # Não usar round aqui!
        return [Decimal(str(p)) / Decimal("100") for p in self._porosidade]

    def rqi(self):
        resultado = []
        for i in range(len(self._df)):
            try:
                permeabilidade = Decimal(str(self._permeabilidade[i]))
                porosidade_dec = self.porosidadeDec()[i]
                if permeabilidade == 0 or porosidade_dec == 0:
                    resultado.append(Decimal("0"))
                else:
                    rqi = Decimal("0.0314") * (permeabilidade / porosidade_dec).sqrt()
                    resultado.append(rqi)
            except Exception:
                resultado.append(Decimal("0"))
        return resultado

    def phi(self):
        resultado = []
        for i in range(len(self._df)):
            try:
                porosidade = self.porosidadeDec()[i]
                if porosidade <= 0 or porosidade >= 1:
                    resultado.append(Decimal("0"))
                else:
                    phi = porosidade / (Decimal("1") - porosidade)
                    resultado.append(phi)
            except Exception as e:
                print(f"Erro no índice {i}: {e}")
                resultado.append(Decimal("0"))
        return resultado

    def fzi(self):
        phi = self.phi()
        rqi = self.rqi()
        return [(r / p if p != 0 else Decimal("0")) for r, p in zip(rqi, phi)]

    def ghe(self):
        fzi = self.fzi()
        resultado = []
        for valor in fzi:
            try:
                if valor < 0.0938:
                    resultado.append(0)
                elif valor < 0.1875:
                    resultado.append(1)
                elif valor < 0.3750:
                    resultado.append(2)
                elif valor < 0.7500:
                    resultado.append(3)
                elif valor < 1.5000:
                    resultado.append(4)
                elif valor < 3.0000:
                    resultado.append(5)
                elif valor < 6.0000:
                    resultado.append(6)
                elif valor < 12.0000:
                    resultado.append(7)
                elif valor < 24.0000:
                    resultado.append(8)
                elif valor < 48.0000:
                    resultado.append(9)
                else:
                    resultado.append(10)
            except:
                resultado.append(0)
        return resultado

    def criaPlanilha(self):
        colunas = {
            'Profundidade': self._profundidade,
            'Porosity (%)': self.porosidade(),
            'Porosity Decimal': self.porosidadeDec(),
            'Permeability (mD)': self._permeabilidade,
            'RQI': self.rqi(),
            'PHI(Z)': self.phi(),
            'FZI': self.fzi(),
            'GHE': self.ghe()
        }
        dfColunas = pd.DataFrame(colunas).fillna(0)

        file_path = self.nomeTabela + 'Alterada.xlsx'

        # Cria ou abre o workbook
        if os.path.exists(file_path):
            workbook = load_workbook(file_path)
        else:
            workbook = Workbook()
            if "Sheet" in workbook.sheetnames:
                workbook.remove(workbook["Sheet"])

        # ---- Planilha 1: Dados
        if "Planilha1" in workbook.sheetnames:
            sheet1 = workbook["Planilha1"]
            # Limpa a planilha existente
            for row in sheet1.iter_rows():
                for cell in row:
                    cell.value = None
        else:
            sheet1 = workbook.create_sheet("Planilha1")

        # Escreve cabeçalhos e dados
        headers = list(dfColunas.columns)
        for col_num, header in enumerate(headers):
            cell = sheet1.cell(row=1, column=col_num+1, value=header)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            if header in ['FZI', 'RQI', 'PHI(Z)', 'GHE']:
                cell.fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
            elif header in ['Profundidade', 'Porosity (%)', 'Porosity Decimal', 'Permeability (mD)']:
                cell.fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

        for row_num, row in enumerate(dfColunas.itertuples(index=False), start=2):
            for col_num, value in enumerate(row):
                cell = sheet1.cell(row=row_num, column=col_num+1, value=float(value) if isinstance(value, Decimal) else value)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                if headers[col_num] in ['Porosity Decimal', 'Profundidade', 'Permeability (mD)', 'Porosity (%)', 'PHI(Z)']:
                    cell.number_format = '0.000'
                elif headers[col_num] in ['RQI', 'FZI']:
                    cell.number_format = '0.############'

        # Ajusta largura das colunas
        for col in range(1, len(headers)+1):
            sheet1.column_dimensions[get_column_letter(col)].width = 20

        # ---- Prepara dados para gráfico
        porosidade_dec = [float(p) for p in dfColunas['Porosity Decimal']]
        permeability = [float(p) for p in dfColunas['Permeability (mD)']]
        ghe = [float(g) for g in dfColunas['GHE']]

        # ---- Planilha 2: Faixas
        fzi_values = [48, 24, 12, 6, 3, 1.5, 0.75, 0.375, 0.1875, 0.0938]
        ghe_labels = list(range(10, 0, -1))
        phi = np.linspace(0.01, 0.5, 300)

        # Calcula faixas
        data = []
        for i in range(len(fzi_values)):
            k = permeabilit(phi, fzi_values[i])
            data.append(k)

        if "Faixas" in workbook.sheetnames:
            sheet_faixas = workbook["Faixas"]
            # Limpa a planilha existente
            for row in sheet_faixas.iter_rows():
                for cell in row:
                    cell.value = None
        else:
            sheet_faixas = workbook.create_sheet("Faixas")

        # Escreve cabeçalhos
        sheet_faixas.cell(row=1, column=1, value="Porosity")
        for i, label in enumerate(ghe_labels):
            sheet_faixas.cell(row=1, column=i+2, value=f"GHE_{label}")

        # Escreve dados
        for row, p in enumerate(phi, start=2):
            sheet_faixas.cell(row=row, column=1, value=p)
            for col, k_array in enumerate(data, start=2):
                sheet_faixas.cell(row=row, column=col, value=k_array[row-2])

        # ---- Cria gráfico no Excel usando openpyxl
        chart = ScatterChart()
        chart.title = "Global Hydraulic Elements (GHE)"
        chart.x_axis.title = "Porosity (decimal)"
        chart.y_axis.title = "Permeability (mD)"
        chart.y_axis.scaling.logBase = 10
        chart.legend.position = "r"
        chart.width = 15  # Aproximadamente 800 pixels
        chart.height = 10  # Aproximadamente 500 pixels

        # Adiciona faixas coloridas
        colors = [
            "FF0000", "FF4500", "FFA500", "FFD700", "ADFF2F",
            "00FA9A", "00CED1", "1E90FF", "8A2BE2", "FF69B4"
        ]

        for i in range(len(fzi_values) - 1):
            xvalues = Reference(sheet_faixas, min_col=1, min_row=2, max_row=301)
            yvalues = Reference(sheet_faixas, min_col=i+2, min_row=2, max_row=301)
            series = Series(yvalues, xvalues, title=f"GHE {ghe_labels[i]}")
            series.graphicalProperties.line.solidFill = colors[i]
            chart.append(series)

        # Adiciona pontos experimentais
        xvalues_exp = Reference(sheet1, min_col=3, min_row=2, max_row=len(porosidade_dec)+1)  # Coluna Porosity Decimal
        yvalues_exp = Reference(sheet1, min_col=4, min_row=2, max_row=len(permeability)+1)  # Coluna Permeability
        series_exp = Series(yvalues_exp, xvalues_exp, title="teste")
        series_exp.marker.symbol = "circle"
        series_exp.marker.size = 6
        series_exp.marker.graphicalProperties.solidFill = "000000"
        series_exp.graphicalProperties.line.noFill = True
        chart.append(series_exp)

        # Insere gráfico na planilha principal
        sheet1.add_chart(chart, "I2")

        # Salva o arquivo
        workbook.save(file_path)

    # O método criGrafico não é mais necessário, pois foi integrado


class Aplicativo:
    def __init__(self, master=None):
        def selecionar_arquivo():
            arquivo = filedialog.askopenfilename(
                title="Selecione o arquivo .xlsx",
                filetypes=[("Planilhas Excel", "*.xlsx")]
            )
            print(arquivo)

            if arquivo:
                preview = pd.read_excel(arquivo, header=None)
                colunas_necessarias = ['Prof. (m)', 'Porosidade (%)', 'Perm Abs Longitud (mD)']
                header_row = None

                for i, row in preview.iterrows():
                    valores_validos = [str(v).strip() for v in row.values if pd.notna(v)]
                    if all(col in valores_validos for col in colunas_necessarias):
                        header_row = i
                        break

                if header_row is not None:
                    df = pd.read_excel(arquivo, header=header_row)
                    df.columns = df.columns.str.strip()

                    if all(col in df.columns for col in colunas_necessarias):
                        nomeTabela = arquivo.split('.')[0]
                        AutomatizacaoPlanilha(df, nomeTabela).criaPlanilha()
                        messagebox.showinfo('Sucesso', 'Sua planilha foi criada com sucesso!')
                    else:
                        cols_faltando = [col for col in colunas_necessarias if col not in df.columns]
                        messagebox.showerror('Erro', f'A planilha está faltando as colunas: {", ".join(cols_faltando)}')
                else:
                    messagebox.showerror('Erro', 'Não foi possível encontrar as colunas esperadas na planilha.')

        self.primeiroContainer = tk.Frame(master, pady=10, padx=100)
        self.primeiroContainer.pack()

        self.segundoContainer = tk.Frame(master, pady=10)
        self.segundoContainer.pack()

        self.terceiroContainer = tk.Frame(master, pady=5)
        self.terceiroContainer.pack()

        self.quartoContainer = tk.Frame(master, pady=10)
        self.quartoContainer.pack()

        self.titulo = tk.Label(
        self.primeiroContainer,text='Antes de selecionar o arquivo, observe\nse a planilha contém as seguintes colunas:\n\nProf. (m)\nPorosidade (%)\nPerm Abs Longitud (m)')
        self.titulo.pack()

        self.btnArquivo = tk.Button(self.segundoContainer, text='Selecionar arquivo', width=25, command=selecionar_arquivo)
        self.btnArquivo.pack()

        #self.btnDocx = tk.Button(self.quartoContainer, text='Converter tabela .docx em planilha .xls', width=40, command=selecionar_e_converter)
        #self.btnDocx.pack()


# Executa o programa
root = tk.Tk()
root.title('Planilhas Petrofisica Lagesed')
app = Aplicativo(root)
root.mainloop()
