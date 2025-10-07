import numpy as np
import matplotlib.pyplot as plt
import pandas as pd

# -----------------------------
# Função para calcular k a partir de phi e FZI
def permeability(phi, fzi):
    phi_e = phi / (1 - phi)          # transformação da porosidade
    k = phi * ((fzi * phi_e) / 0.0314) ** 2
    return k

# -----------------------------
def principal(phi_points, fzi_points, file_path):
    # Tabela FZI ↔ GHE
    fzi_values = [48, 24, 12, 6, 3, 1.5, 0.75, 0.375, 0.1875, 0.0938]
    ghe_labels = list(range(10, 0, -1))  # GHE 10 até GHE 1

    # Porosidade (decimal)
    phi = np.linspace(0.01, 0.5, 300)

    # Cores (do verde ao vermelho)
    colors = plt.cm.jet(np.linspace(0, 1, len(fzi_values)))

    # -----------------------------
    # Criar gráfico
    plt.figure(figsize=(9, 7))

    # Preencher faixas entre curvas
    for i in range(len(fzi_values) - 1):
        k1 = permeability(phi, fzi_values[i])       # curva superior
        k2 = permeability(phi, fzi_values[i + 1])   # curva inferior
        plt.fill_between(phi, k1, k2, color=colors[i], alpha=0.7, label=f"GHE {ghe_labels[i]}")

    # -----------------------------
    # Pontos experimentais
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

    # Salvar gráfico como imagem
    img_file = "ghe_plot.png"
    plt.savefig(img_file, dpi=300)
    plt.close()

    # -----------------------------
    # Criar DataFrame com os pontos
    dfColunas = pd.DataFrame({
        "Phi Points": phi_points,
        "FZI Points": fzi_points,
        "K Points": k_points
    })

    # -------------------------
    # Exemplo de dados adicionais (para gráfico por GHE)
    data = {
        "Porosidade (%)": [5.3, 12, 20.5, 15, 30, 60],
        "Permeabilidade (mD)": [0.035, 1.2, 8.7, 2.1, 50, 20],
        "GHE": [3, 3, 4, 5, 4, 6]
    }
    df = pd.DataFrame(data)

    # -------------------------
    # Salvar no Excel com gráfico e imagem
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        # Aba 1: pontos FZI
        dfColunas.to_excel(writer, sheet_name='Pontos FZI', index=False)
        worksheet1 = writer.sheets['Pontos FZI']
        worksheet1.set_column("A:C", 20)

        # Aba 2: dados com GHE
        df.to_excel(writer, sheet_name="Dados", index=False)
        worksheet = writer.sheets["Dados"]

        workbook = writer.book

        # Criar gráfico de dispersão
        chart = workbook.add_chart({"type": "scatter"})

        # Adicionar série por GHE
        unique_ghe = df["GHE"].unique()
        for ghe in unique_ghe:
            ghe_data = df[df["GHE"] == ghe]
            start_row = ghe_data.index.min() + 1  # +1 por causa do cabeçalho
            end_row = ghe_data.index.max() + 1

            chart.add_series({
                "name":       f"GHE {ghe}",
                "categories": ["Dados", start_row, 0, end_row, 0],  # Porosidade
                "values":     ["Dados", start_row, 1, end_row, 1],  # Permeabilidade
                "marker": {"type": "circle", "size": 7},
            })

        # Eixos do gráfico
        chart.set_x_axis({"name": "Porosidade (%)", "label_position": "low"})
        chart.set_y_axis({
            "name": "Permeabilidade (mD)",
            "log_base": 10,
            "min": 0.01,
            "max": 10000,
        })
        chart.set_title({"name": "Gráfico por GHE"})
        chart.set_legend({"position": "bottom"})

        # Inserir gráfico na aba de dados
        worksheet.insert_chart("E2", chart)

        # Aba 3: gráfico GHE como imagem
        worksheet2 = workbook.add_worksheet("Gráfico GHE")
        worksheet2.insert_image("B2", img_file, {"x_scale": 0.8, "y_scale": 0.8})

    print(f"✅ Planilha gerada com sucesso: {file_path}")
