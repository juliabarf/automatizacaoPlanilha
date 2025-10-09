import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import os

# -----------------------------
def permeabilit(phi, fzi):
    phi = np.clip(phi, 1e-6, 0.99)
    phi_e = phi / (1 - phi)
    k = phi * ((fzi * phi_e) / 0.0314) ** 2
    return k

# -----------------------------
def principal(porosidade_dec, permeability, ghe,  file_path):
    # Tabela FZI ↔ GHE
    fzi_values = [48, 24, 12, 6, 3, 1.5, 0.75, 0.375, 0.1875, 0.0938]
    ghe_labels = list(range(10, 0, -1))

    phi = np.linspace(0.01, 0.5, 300)
    colors = plt.cm.jet(np.linspace(0, 1, len(fzi_values)))

    fig, ax = plt.subplots(figsize=(9, 7))

    # Preencher faixas de GHE
    for i in range(len(fzi_values) - 1):
        k1 = permeabilit(phi, fzi_values[i])
        k2 = permeabilit(phi, fzi_values[i + 1])
        ax.fill_between(phi, k1, k2, color=colors[i], alpha=0.7, label=f"GHE {ghe_labels[i]}")

    # Pontos experimentais
    k_points = [permeabilit(porosidade_dec[i], permeability[i]) for i in range(len(porosidade_dec))]
    sc = ax.scatter(porosidade_dec, k_points, color="black", s=50)

    ax.set_yscale("log")
    ax.set_xscale("linear")
    ax.set_xlabel("Porosity (decimal)")
    ax.set_ylabel("Permeability (mD)")
    ax.set_title("Global Hydraulic Elements (GHE)")
    ax.grid(True, which="both", ls="--", lw=0.5)

    # DataFrame para armazenar legendas clicadas
    df_clicks = pd.DataFrame(columns=["Porosity (decimal)", "Permeability (mD)", "K"])
    annotations = []
    # -----------------------------
    # Função para adicionar legenda ao clicar
    def on_click(event):
        nonlocal df_clicks, annotations

        if event.inaxes == ax:
            distances = np.sqrt((np.array(porosidade_dec) - event.xdata) ** 2 +
                                (np.log10(np.array(k_points)) - np.log10(event.ydata)) ** 2)
            min_index = np.argmin(distances)

            if distances[min_index] < 0.02:
                # Verifica se já existe anotação nesse ponto
                existing = [ann for ann, idx in annotations if idx == min_index]

                if existing:
                    # Remover anotação e entrada do DataFrame
                    existing[0].remove()
                    annotations = [(ann, idx) for ann, idx in annotations if idx != min_index]
                    df_clicks = df_clicks[df_clicks["Porosity (decimal)"] != porosidade_dec[min_index]]
                    print(f"❌ Anotação removida: Porosidade={porosidade_dec[min_index]:.2f}")
                else:
                    # Adicionar anotação no gráfico
                    ann = ax.annotate(
                        f"SÉRIE 'GHE' \nPonto Porosity '{porosidade_dec[min_index]:.2f}'\n( {porosidade_dec[min_index]:.2f},{permeability[min_index]})",
                        (porosidade_dec[min_index], k_points[min_index]),
                        textcoords="offset points", xytext=(10, 10),
                        arrowprops=dict(arrowstyle="->", color='red'),
                        fontsize=9, color='blue'
                    )
                    annotations.append((ann, min_index))
                    print(f"✅ Anotação adicionada: Porosidade={porosidade_dec[min_index]:.2f} Permeability={permeability[min_index]:.2f}")

                    # Adicionar ao DataFrame
                    new_row = {
                        "Porosity (decimal)": porosidade_dec[min_index],
                        "Permeability (mD)": permeability[min_index],
                        "K": k_points[min_index]
                    }
                    df_clicks = pd.concat([df_clicks, pd.DataFrame([new_row])], ignore_index=True)

                # Atualizar o gráfico
                fig.canvas.draw()

                # Salvar/atualizar Excel
                with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                    # Aba original
                    df = pd.DataFrame({
                        "Porosity (decimal)": porosidade_dec,
                        "Permeability (mD)": permeability,
                        "K Points": k_points
                    })
                    df.to_excel(writer, sheet_name='Planilha1', index=False)

                    # Aba de pontos clicados
                    df_clicks.to_excel(writer, sheet_name='Pontos_Clicados', index=False)

                    # Inserir gráfico
                    img_file = "ghe_plot.png"
                    plt.savefig(img_file, dpi=300)
                    workbook = writer.book
                    worksheet2 = workbook.add_worksheet("Gráfico")
                    worksheet2.insert_image("B2", img_file, {"x_scale": 0.8, "y_scale": 0.8})

    fig.canvas.mpl_connect("button_press_event", on_click)
    plt.show()

# -----------------------------
# Exemplo de uso
phi_points = [0.05, 0.12, 0.2, 0.15, 0.3, 0.6]
fzi_points = [3, 3, 4, 2, 5, 6]
principal(phi_points, fzi_points, "GHE_Excel.xlsx")
