
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import math
from decimal import Decimal
from decimal import Decimal, getcontext
import xlsxwriter
from converte import selecionar_e_converter
from graficoCurvas import principal

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
        writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
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
        writer.close()

        porosidade_dec = []
        for k in range(len(colunas['Porosity Decimal'])):
            # converte cada valor Decimal para float
            porosidade_dec.append(float(colunas['Porosity Decimal'][k]))

        print(porosidade_dec)

        permeability = []
        for f in range(len(colunas['Permeability (mD)'])):
            # converte cada valor Decimal para float
            permeability.append(float(colunas['Permeability (mD)'][f]))

        print(permeability)

        ghe = []
        for f in range(len(colunas['GHE'])):
            ghe.append(float(colunas['GHE'][f]))

        print(ghe)
        principal(porosidade_dec, permeability, ghe, file_path)

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