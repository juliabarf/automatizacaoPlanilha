import pandas as pd
import math
#primeira coluna -> profundidade
#segunda coluna -> porosidade(decimal) = 0.192
#terceira coluna -> permeabilidade(mD)
#ao invés de criar uma nova aba, eu posso só criar novas colunas na mesma planilha
#agora preciso criar novas tabelas para colocar os resultados dos def que eu criei para poder fazer

#classse que faz as operações de cada coluna
class AutomatizacaoPlanilha:
    def __init__(self, caminho):
        df = pd.read_excel(caminho)
        self._df = df
        self._profundidade = df['Profundidade']
        self._porosidade = df['Porosity (%)']
        self._porosidade1 = df['Porosity Decimal']
        self._permeabilidade = df['Permeability (mD)']

    def teste(self):
        dados = {
            'profundidade': self._profundidade,
            'porosidade': self._porosidade,
            'porosidade1': self._porosidade1,
            'permeabilidade':  self._permeabilidade,
        }
        lista = []
        for i in range(len(self._df)):
            lista.append(self._profundidade[i])
        print(lista)

    def rqi(self):
        #0,0314 * raiz(Permeabilidade/Porosidade)
        listaRQI = []

        for i in range(len(self._df)):
            colunaRQI = 0.0314 * (math.sqrt(self._permeabilidade[i]/self._porosidade1[i]))
            listaRQI.append(colunaRQI)
        print(listaRQI)

    def phi(self):
        #porosidade/(100-porosidade)
        listaPHI = []

        for i in range(len(self._df)):
            phi = self._porosidade[i]/(100 - self._porosidade[i]) * 100
            listaPHI.append(phi)
        print(listaPHI)

    #depois que conseguir os outros resultados
    def fzi(self):
        #rqi/phi
        phi = AutomatizacaoPlanilha.phi(self)
        rqi = AutomatizacaoPlanilha.rqi(self)

        print(phi, rqi)

    def litofaceis(self):
        pass

teste = AutomatizacaoPlanilha('fruta.xlsx')
teste.fzi()


