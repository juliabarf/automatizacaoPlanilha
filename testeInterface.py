import tkinter as tk
from tkinter import LEFT
from tkinter import filedialog
import pandas as pd
import xlsxwriter
import math


class AutomatizacaoPlanilha:
    def __init__(self, arquivo):
        # lê a planilha com o caminho da tabela fornecido pela pessoa e coloca o conteúdo das colunas em variáveis
        df = pd.read_excel(arquivo)
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
        # 0,0314 * raiz(Permeabilidade/Porosidade)
        listaRQI = []

        for i in range(len(self._df)):
            colunaRQI = 0.0314 * (math.sqrt(self._permeabilidade[i] / self.porosidadeDec()[i]))
            listaRQI.append(colunaRQI)
        return (listaRQI)

    def phi(self):
        # porosidade/(100-porosidade)
        listaPHI = []

        for i in range(len(self._df)):
            phi = self._porosidade[i] / (100 - self._porosidade[i]) * 100
            listaPHI.append(round(phi))
        return (listaPHI)

    # depois que conseguir os outros resultados
    def fzi(self):
        # rqi/phi
        phi = self.phi()
        rqi = self.rqi()
        listaFZI = []

        for i in range(len(self._df)):
            fzi = (rqi[i] / phi[i]) * 100
            listaFZI.append(fzi)
        return (listaFZI)

    def litofaceis(self):
        pass

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
            'PHI(Z)': self.phi(),
            'FZI': self.fzi()
        }
        dfColunas = pd.DataFrame(colunas)

        writer = pd.ExcelWriter('tabela.xlsx', engine='xlsxwriter')
        dfColunas.to_excel(writer, sheet_name='Planilha1', index=False)

        workbook = writer.book
        worksheet = writer.sheets['Planilha1']

        # Define a centralização das células na horizontal e vertical
        cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})

        # Aplica o formato nas células da planilha
        for col_num, value in enumerate(dfColunas.columns.values):
            worksheet.write(0, col_num, value, cell_format)
            for row in range(1, len(dfColunas) + 1):
                worksheet.write(row, col_num, dfColunas.iloc[row - 1, col_num], cell_format)

        # Define largura das colunas
        worksheet.set_column('A:A', 20)
        worksheet.set_column('B:B', 20)
        worksheet.set_column('C:C', 20)
        worksheet.set_column('D:D', 20)
        worksheet.set_column('E:E', 15)
        worksheet.set_column('F:F', 15)
        worksheet.set_column('G:G', 15)

        writer.close()


#chama esta função para o botão
def selecionar_arquivo():
    """
    Função que abre uma caixa de diálogo para selecionar um arquivo
    e imprime o seu caminho.
    """
    arquivo = filedialog.askopenfilename()
    if arquivo:
        print("Arquivo selecionado:", arquivo)
        AutomatizacaoPlanilha(arquivo).criaPlanilha()


class Aplicativo:
    def __init__(self, master = None):
        #container para o texto de atenção
        self.primeiroContainer = tk.Frame(master)
        self.primeiroContainer['pady'] = 10
        self.primeiroContainer['padx'] = 100
        self.primeiroContainer.pack()

        #container para o botão
        self.segundoContainer = tk.Frame(master)
        self.segundoContainer['pady'] = 10
        self.segundoContainer.pack()

        """
        self.terceiroContainer = tk.Frame(master)
        self.terceiroContainer['pady'] = 10
        self.terceiroContainer.pack(anchor='w')
        """

        #texto de atenção
        self.titulo = tk.Label(self.primeiroContainer, text=' Antes de selecionar o arquivo, observe\n se a planilha contém as seguintes colunas:\n \nProf. (m) \nPorosidade (%) \nPerm Abs Longitud (md)')
        self.titulo.pack()

        """
        #texto parte 2
        self.titulo2 = tk.Label(self.terceiroContainer, text='Prof.')
        self.titulo2.pack(anchor='w')
        """

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
app = Aplicativo(root)
root.mainloop()