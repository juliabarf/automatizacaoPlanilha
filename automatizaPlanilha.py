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
        self.df = df
        self.profundidade = df['Profundidade']
        self.porosidade = df['Porosity (%)']
        self.porosidade1 = df['Porosity Decimal']
        self.permeabilidade = df['Permeability (mD)']


    def teste(self):
        dados = {
            'profundidade': self.profundidade,
            'porosidade': self.porosidade,
            'porosidade1': self.porosidade1,
            'permeabilidade':  self.permeabilidade,
        }
        print(dados)
    def rqi(self):
        #0,0314 * raiz(Permeabilidade/Porosidade)
        colunaRQI = 0.0314 * (math.sqrt(self.permeabilidade[0]/self.porosidade1[0]))
        print(colunaRQI)

    def phi(self):
        #porosidade/(100-porosidade)
        phi = self.porosidade[0]/(100 - self.porosidade[0])
        print(phi*100) #colocar apenas os números antes da vírgula. esse é porcentagem



    #depois que conseguir os outros resultados
    def fzi(self):
        #rqi/phi
        pass

    def litofaceis(self):
        pass

teste = AutomatizacaoPlanilha('fruta.xlsx')
teste.teste()


