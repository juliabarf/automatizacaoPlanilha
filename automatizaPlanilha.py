import pandas as pd
import math
#primeira coluna -> profundidade
#segunda coluna -> porosidade(decimal) = 0.192
#terceira coluna -> permeabilidade(mD)


#classse que faz as operações de cada coluna
class AutomatizacaoPlanilha:
    def __init__(self, planilha):
        df = pd.read_excel(planilha)
        self.df = df
        self.profundidade = df['Profundidade']
        self.porosidade = df['Porosity (%)']
        self.porosidade1 = df['Porosity Decimal']
        self.permeabilidade = df['Permeability (mD)']




    def teste(self, teste1):
        planilha = {
            'profundidade': 1,
            'porosidade': 2,
            'porosidade1': "testando",
            'permeabilidade': "tesstando",
        }
        planilha['profundidade'] = teste1
        print(planilha)
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
teste.teste('oba')


