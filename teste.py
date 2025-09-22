import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import xlsxwriter

# -----------------------------
# Função para calcular k a partir de phi e FZI
def permeability(phi, fzi):
    phi_e = phi / (1 - phi)          # transformação da porosidade
    k = phi * ((fzi * phi_e) / 0.0314) ** 2
    return k

# -----------------------------
# Tabela FZI ↔ GHE
fzi_values = [48, 24, 12, 6, 3, 1.5, 0.75, 0.375, 0.1875, 0.0938]
ghe_labels = list(range(10, 0, -1))  # GHE 10 até GHE 1

# Porosidade (decimal)
phi = np.linspace(0.01, 0.5, 300)

# Cores (do verde ao vermelho)
colors = plt.cm.jet(np.linspace(0, 1, len(fzi_values)))

# -----------------------------
# Criar gráfico
plt.figure(figsize=(9,7))

# Preencher faixas entre curvas
for i in range(len(fzi_values)-1):
    k1 = permeability(phi, fzi_values[i])       # curva superior
    k2 = permeability(phi, fzi_values[i+1])     # curva inferior
    plt.fill_between(phi, k1, k2, color=colors[i], alpha=0.7, label=f"GHE {ghe_labels[i]}")

# -----------------------------
# Pontos experimentais fictícios
phi_points = [0.08, 0.12, 0.18, 0.25, 0.35]
fzi_points = [6, 12, 3, 1.5, 0.75]
k_points = [permeability(phi_points[i], fzi_points[i]) for i in range(len(phi_points))]

plt.scatter(phi_points, k_points, color="black", marker="o", s=30, label="Pontos FZI")

# -----------------------------
# Escalas logarítmicas
plt.yscale("log")
plt.xscale("linear")

# Labels e título
plt.xlabel("Porosity (decimal)")
plt.ylabel("Permeability (mD)")
plt.title("Global Hydraulic Elements (GHE)")

# Legenda
plt.legend(title="GHE", bbox_to_anchor=(1.05, 1), loc='upper left')

plt.grid(True, which="both", ls="--", lw=0.5)
plt.tight_layout()

# -----------------------------
# Salvar gráfico em imagem temporária
img_file = "ghe_plot.png"
plt.savefig(img_file, dpi=300)
plt.close()

# -----------------------------
# Criar planilha Excel com os dados e gráfico
file_path = "GHE_test.xlsx"

# Criar dataframe com pontos fictícios
df = pd.DataFrame({
    "Porosity": phi_points,
    "FZI": fzi_points,
    "Permeability (mD)": np.round(k_points, 2)
})

with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
    df.to_excel(writer, sheet_name="Dados", index=False)

    # Inserir imagem do gráfico em outra aba
    workbook = writer.book
    worksheet = workbook.add_worksheet("Gráfico")
    worksheet.insert_image("B2", img_file, {"x_scale": 0.8, "y_scale": 0.8})

print(f"✅ Planilha gerada com sucesso: {file_path}")
