import pandas as pd

base = pd.read_excel('tabela.xlsx')
profundidade = base['Profundidade']
style = profundidade.style.set_properties(**{'background-color': 'red', 'color': 'white'})

print(base)