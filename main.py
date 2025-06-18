import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import math
import xlsxwriter


class AutomatizacaoPlanilha:
    def __init__(self, arquivo):
        # lê a planilha com o caminho da tabela fornecido pela pessoa e coloca o conteúdo das colunas em variáveis
        df = pd.read_excel(arquivo)
        nomeTabela = arquivo.split('.')[0]
        self.nomeTabela = nomeTabela
        self._df = df
        convert_dic = {'Prof. (m)': float}
        self._df = self._df.astype(convert_dic)
        self._profundidade = df['Prof. (m)']
        self._porosidade = df['Porosidade (%)']
        self._permeabilidade = df['Perm Abs Longitud (mD)']

    def porosidade(self):
        lista_porosidade = []
        for i in range(len(self._porosidade)):
            lista_porosidade.append(self._porosidade[i])
        return lista_porosidade

    def porosidadeDec(self):
        listaPorosidadeDec = []
        for i in range(len(self._porosidade)):
            colunaPorosidadeDec = self._porosidade[i] / 100
            listaPorosidadeDec.append(round(colunaPorosidadeDec, 3))
        return listaPorosidadeDec

    def rqi(self):
        listaRQI = []
        for i in range(len(self._df)):
            try:
                permeabilidade = float(self._permeabilidade[i])
                porosidade_dec = float(self.porosidadeDec()[i])

                if pd.isna(permeabilidade) or pd.isna(porosidade_dec) or permeabilidade == 0 or porosidade_dec == 0:
                    listaRQI.append(0)
                else:
                    colunaRQI = 0.0314 * math.sqrt(permeabilidade / porosidade_dec)
                    listaRQI.append((colunaRQI))

            except (ValueError, TypeError, ZeroDivisionError):
                listaRQI.append(0)
        return listaRQI

    def phi(self):
        listaPHI = []
        for i in range(len(self._df)):
            try:
                porosidade = float(self._porosidade[i])
                if pd.isna(porosidade) or porosidade == 0 or porosidade == 100:
                    listaPHI.append(0)
                else:
                    phi = porosidade / (100 - porosidade) * 100
                    listaPHI.append(round(phi))
            except (ValueError, TypeError, ZeroDivisionError):
                listaPHI.append(0)
        return listaPHI

    # depois que conseguir os outros resultados
    def fzi(self):
        phi = self.phi()
        rqi = self.rqi()
        listaFZI = []

        for i in range(len(self._df)):
            try:
                valor_phi = float(phi[i])
                valor_rqi = float(rqi[i])

                if pd.isna(valor_phi) or valor_phi == 0:
                    listaFZI.append(0)
                else:
                    fzi = (valor_rqi / valor_phi) * 100
                    listaFZI.append(round(fzi, 4))
            except (ValueError, TypeError, ZeroDivisionError, IndexError):
                listaFZI.append(0)
        return listaFZI

    def ghe(self):
        fzi = self.fzi()
        listaGHE = []

        for i in range(len(self._df)):
            try:
                valor = fzi[i]
                if valor < 0.0938:
                    listaGHE.append(0)
                elif valor >= 0.0938 and valor < 0.1875:
                    listaGHE.append(1)
                elif valor >= 0.1875 and valor < 0.375:
                    listaGHE.append(2)
                elif valor >= 0.375 and valor < 0.75:
                    listaGHE.append(3)
                elif valor >= 0.75 and valor < 1.5:
                    listaGHE.append(4)
                elif valor >= 1.5 and valor < 3.0:
                    listaGHE.append(5)
                elif valor >= 3.0 and valor < 6.0:
                    listaGHE.append(6)
                elif valor >= 6.0 and valor < 12.0:
                    listaGHE.append(7)
                elif valor >= 12.0 and valor < 24.0:
                    listaGHE.append(8)
                elif valor >= 24.0 and valor < 48.0:
                    listaGHE.append(9)
                elif valor >= 48.0:
                    listaGHE.append(10)
                else:
                    listaGHE.append(0)
            except (IndexError, ValueError, TypeError):
                listaGHE.append(0)

        return listaGHE

    def criaPlanilha(self):
        # self._df['RQI'] = self._porosidade1 * 100
        # for i in range(len(self._df)):
        #    self._df['RQI'][i] = self.rqi()[i]

        colunas = {
            'Profundidade': self._profundidade,
            'Porosity (%)': self.porosidade(),
            'Porosity Decimal': self.porosidadeDec(),
            'Permeability (mD)': self._permeabilidade,
            'RQI': self.rqi(),
            'PHI(Z)': [f"{round(p, 2)}%" for p in self.phi()],
            'FZI': self.fzi(),
            'GHE': self.ghe()
        }
        dfColunas = pd.DataFrame(colunas)
        dfColunas = dfColunas.fillna(0)  # substitui células vazias por 0

        writer = pd.ExcelWriter(self.nomeTabela+'Alterada.xlsx', engine='xlsxwriter')
        dfColunas.to_excel(writer, sheet_name='Planilha1', index=False)


        workbook = writer.book
        worksheet = writer.sheets['Planilha1']

        # Define formatação para células
        cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
        decimal_format = workbook.add_format({'num_format': '0.000', 'align': 'center', 'valign': 'vcenter'})
        decimal_format3 = workbook.add_format({'num_format': '0.0000', 'align': 'center', 'valign': 'vcenter'})
        decimal_format4 = workbook.add_format({'num_format': '0', 'align': 'center', 'valign': 'vcenter'})

        # Lista de colunas que devem receber formatação com 3 e 2 casas decimais
        colunas_com_decimal = ['Porosity Decimal', 'Profundidade', 'Permeability (mD)', 'Porosity (%)']
        coluna3_dec = ['RQI','FZI']
        colunaPHI = ['PHI(Z)']

        # Escreve os dados com a formatação apropriada
        for col_num, value in enumerate(dfColunas.columns.values):
            worksheet.write(0, col_num, value, cell_format)
            for row in range(1, len(dfColunas) + 1):
                valor = dfColunas.iloc[row - 1, col_num]

                if value in colunas_com_decimal:
                    worksheet.write(row, col_num, valor, decimal_format)
                elif value in coluna3_dec:
                    worksheet.write(row, col_num, valor, decimal_format3)
                elif value in colunaPHI:
                    worksheet.write(row, col_num, valor, decimal_format4)

                else:
                    worksheet.write(row, col_num, valor, cell_format)

        # Cabeçalhos coloridos ( "FZI", "RQI", "")
        for col_num, value in enumerate(dfColunas.columns.values):
            if value == 'FZI' or value == 'RQI' or value == 'PHI(Z)' or value == 'GHE':
                worksheet.write(0, col_num, value,
                                workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFFF99'}))

            elif value == 'Profundidade' or value == 'Porosity (%)' or value == 'Porosity Decimal' or value == 'Permeability (mD)':
                worksheet.write(0, col_num, value,
                                workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFCCCC'}))
            else:
                worksheet.write(0, col_num, value, cell_format)

        #colore as células
        #style = self._profundidade.style.set_properties(**{'background-color': 'red', 'color': 'white'})

        # Define largura das colunas
        worksheet.set_column('A:A', 20)
        worksheet.set_column('B:B', 20)
        worksheet.set_column('C:C', 20)
        worksheet.set_column('D:D', 20)
        worksheet.set_column('E:E', 15)
        worksheet.set_column('F:F', 15)
        worksheet.set_column('G:G', 15)

        writer.close()




class Aplicativo:
    def __init__(self, master = None):
        def selecionar_arquivo():
            arquivo = filedialog.askopenfilename()
            if arquivo:
                if arquivo.lower().endswith('.xlsx'):
                    try:
                        # Lê a planilha
                        df = pd.read_excel(arquivo)

                        # Colunas para verificar
                        colunas_necessarias = ['Prof. (m)', 'Porosidade (%)', 'Perm Abs Longitud (mD)']

                        # Verifica se todas as colunas existem na planilha
                        if all(col in df.columns for col in colunas_necessarias):
                            messagebox.showinfo('Processo concluído.', 'Sua planilha foi criada com sucesso!')
                            print("Arquivo criado com sucesso!")
                            AutomatizacaoPlanilha(arquivo).criaPlanilha()
                        else:
                            cols_faltando = [col for col in colunas_necessarias if col not in df.columns]
                            messagebox.showerror('Erro',
                                                 f'A planilha está faltando as colunas: {", ".join(cols_faltando)}')
                    except Exception as e:
                        messagebox.showerror('Erro', f'Falha ao ler o arquivo:\n{str(e)}')
                else:
                    messagebox.showerror('Erro', 'Por favor, selecione um arquivo com extensão .xlsx.')

        #container para o texto de atenção
        self.primeiroContainer = tk.Frame(master)
        self.primeiroContainer['pady'] = 10
        self.primeiroContainer['padx'] = 100
        self.primeiroContainer.pack()

        #container para o botão
        self.segundoContainer = tk.Frame(master)
        self.segundoContainer['pady'] = 10
        self.segundoContainer.pack()

        self.terceiroContainer = tk.Frame(master)
        self.terceiroContainer['pady'] = 10
        self.terceiroContainer.pack()

        #texto de atenção
        self.titulo = tk.Label(self.primeiroContainer, text=' Antes de selecionar o arquivo, observe\n se a planilha contém as seguintes colunas:\n \nProf. (m) \nPorosidade (%) \nPerm Abs Longitud (md)')
        self.titulo.pack()

        #botão - arquivo
        self.btnArquivo = tk.Button(self.segundoContainer, text='Selecionar arquivo', width='25', command=selecionar_arquivo)
        self.btnArquivo.pack()




#cria a janela principal
#janela = tk.Tk()
#janela.title("Selecionar Arquivo")

#cria o botão
#botao_selecionar = tk.Button(janela, text="Selecionar Arquivo", command=selecionar_arquivo)
#botao_selecionar.pack()

#executa
# a janela
root = tk.Tk()
root.title('Planilhas Lagesed')
#root.iconbitmap('iconApp.ico')
app = Aplicativo(root)
root.mainloop()

#AutomatizacaoPlanilha('tabelaDoc.xlsx').ghe()