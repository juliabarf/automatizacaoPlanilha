import numpy as np
import pandas as pd
import xlsxwriter
import os

# -----------------------------
def permeabilit(phi, fzi):
    phi = np.clip(phi, 1e-6, 0.99)
    phi_e = phi / (1 - phi)
    k = phi * ((fzi * phi_e) / 0.0314) ** 2
    return k

# -----------------------------
def principal(porosidade_dec, permeability, ghe, file_path):
    fzi_values = [48, 24, 12, 6, 3, 1.5, 0.75, 0.375, 0.1875, 0.0938]
    ghe_labels = list(range(10, 0, -1))
    phi = np.linspace(0.01, 0.5, 300)

    # Calcula faixas
    data = []
    for i in range(len(fzi_values)):
        k = permeabilit(phi, fzi_values[i])
        data.append(k)

    # Cria arquivo Excel
    with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
        workbook = writer.book

        # ---- Planilha 1: Dados
        df = pd.DataFrame({"Porosity (decimal)": porosidade_dec, "Permeability (mD)": permeability})
        df.to_excel(writer, sheet_name="Planilha1", index=False)
        worksheet = writer.sheets["Planilha1"]

        # ---- Planilha 2: Faixas
        faixa_df = pd.DataFrame({"Porosity": phi})
        for i, k in enumerate(data):
            faixa_df[f"GHE_{ghe_labels[i]}"] = k
        faixa_df.to_excel(writer, sheet_name="Faixas", index=False)
        ws_faixas = writer.sheets["Faixas"]

        # ---- Cria gráfico no Excel
        chart = workbook.add_chart({"type": "scatter", "subtype": "smooth"})

        # Adiciona faixas coloridas
        colors = [
            "#FF0000", "#FF4500", "#FFA500", "#FFD700", "#ADFF2F",
            "#00FA9A", "#00CED1", "#1E90FF", "#8A2BE2", "#FF69B4"
        ]

        for i in range(len(fzi_values) - 1):
            chart.add_series({
                "name": f"GHE {ghe_labels[i]}",
                "categories": ["Faixas", 1, 0, 300, 0],  # coluna da Porosidade
                "values": ["Faixas", 1, i+1, 300, i+1],  # coluna de K
                "line": {"color": colors[i]},
            })

        # Adiciona pontos experimentais
        chart.add_series({
            "name": "teste",
            "categories": ["Planilha1", 1, 0, len(porosidade_dec), 0],
            "values": ["Planilha1", 1, 1, len(permeability), 1],
            "marker": {"type": "circle", "size": 6, "border": {"color": "black"}, "fill": {"color": "black"}},
            "line": {"none": True},
        })

        # Configurações do gráfico
        chart.set_title({"name": "Global Hydraulic Elements (GHE)"})
        chart.set_x_axis({"name": "Porosity (decimal)", "min": 0, "max": 0.5})
        chart.set_y_axis({"name": "Permeability (mD)", "log_base": 10})
        chart.set_legend({"position": "right"})
        chart.set_size({"width": 800, "height": 500})

        # Insere gráfico na planilha principal
        worksheet.insert_chart("E2", chart)

# -----------------------------
# Exemplo de uso
phi_points = [0.05, 0.12, 0.2, 0.15, 0.3, 0.4]
fzi_points = [3, 3, 4, 2, 5, 6]
