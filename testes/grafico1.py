import numpy as np
import pandas as pd

# Função para calcular permeabilidade
def permeability(phi, fzi):
    phi = np.clip(phi, 1e-6, 0.99)
    phi_e = phi / (1 - phi)
    k = phi * ((fzi * phi_e) / 0.0314) ** 2
    return k

# Porosidade (decimal)
phi = np.linspace(0.01, 0.5, 50)  # 50 pontos de exemplo

# Valores FZI correspondentes a cada GHE
fzi_values = [48, 24, 12, 6, 3, 1.5, 0.75, 0.375, 0.1875, 0.0938]
ghe_labels = [f"GHE {i}" for i in range(10, 0, -1)]  # GHE 10 até 1

# Criar dicionário para DataFrame
data = {"Porosity": phi}
for ghe, fzi in zip(ghe_labels, fzi_values):
    data[ghe] = permeability(phi, fzi)

# Criar DataFrame
df = pd.DataFrame(data)

# Salvar em Excel
file_path = "Tabela_GHE.xlsx"
with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='GHE_Table', index=False)
    worksheet = writer.sheets['GHE_Table']
    worksheet.set_column("A:K", 15)  # Ajustar largura das colunas

print(f"✅ Tabela gerada com sucesso: {file_path}")
