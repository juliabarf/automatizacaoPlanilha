import pandas as pd
import math
#primeira coluna -> profundidade
#segunda coluna -> porosidade(decimal) = 0.192
#terceira coluna -> permeabilidade(mD)
#ao invés de criar uma nova aba, eu posso só criar novas colunas na mesma planilha
#agora preciso criar novas tabelas para colocar os resultados dos def que eu criei para poder fazer
#verificar o arredondamento da coluna PHI


#classse que faz as operações de cada coluna
class AutomatizacaoPlanilha:
    def __init__(self, caminho):
        #lê a planilha com o caminho da tabela fornecido pela pessoa e coloca o conteúdo das colunas em variáveis
        df = pd.read_excel(caminho)
        self._df = df
        self._profundidade = df['Prof. (m)']
        self._porosidade = df['Porosidade (%)']
        self._permeabilidade = df['Perm Abs Longitud (mD)']

    def porosidade(self):
        lista_porosidade = []
        for i in range(len(self._porosidade)):
            lista_porosidade.append(round(self._porosidade[i],3))
        return lista_porosidade

    def porosidadeDec(self):
        listaPorosidadeDec = []
        for i in range(len(self._porosidade)):
            colunaPorosidadeDec = self._porosidade[i]/100
            listaPorosidadeDec.append(round(colunaPorosidadeDec,3))
        return listaPorosidadeDec

    def rqi(self):
        #0,0314 * raiz(Permeabilidade/Porosidade)
        listaRQI = []

        for i in range(len(self._df)):
            colunaRQI = 0.0314 * (math.sqrt(self._permeabilidade[i]/self.porosidadeDec()[i]))
            listaRQI.append(colunaRQI)
        return(listaRQI)

    def phi(self):
        #porosidade/(100-porosidade)
        listaPHI = []

        for i in range(len(self._df)):
            phi = self._porosidade[i]/(100 - self._porosidade[i]) * 100
            listaPHI.append(round(phi))
        return(listaPHI)

    #depois que conseguir os outros resultados
    def fzi(self):
        #rqi/phi
        phi = self.phi()
        rqi = self.rqi()
        listaFZI = []

        for i in range(len(self._df)):
            fzi = (rqi[i] / phi[i]) * 100
            listaFZI.append(fzi)
        return(listaFZI)

    def litofaceis(self):
        pass

    def criaPlanilha(self):
        #self._df['RQI'] = self._porosidade1 * 100
        # for i in range(len(self._df)):
        #    self._df['RQI'][i] = self.rqi()[i]
        colunas = {
            'profundidade': self._profundidade,
            'porosidade (%)': self.porosidade(),
            'porosidade decimal': self.porosidadeDec(),
            'permeabilidade': self._permeabilidade,
            'RQI': self.rqi(),
            'PHI': self.phi(),
            'FZI': self.fzi()
        }
        dfColunas = pd.DataFrame(colunas)
        print(dfColunas)
        dfColunas.to_excel('tabela.xlsx', index=False)
        return colunas


teste = AutomatizacaoPlanilha('tabelaDoc.xlsx')
teste.criaPlanilha()


