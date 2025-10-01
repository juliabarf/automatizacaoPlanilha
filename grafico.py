import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import os

# -----------------------------
def permeability(phi, fzi):
    phi = np.clip(phi, 1e-6, 0.99)
    phi_e = phi / (1 - phi)
    k = phi * ((fzi * phi_e) / 0.0314) ** 2
    return k

# -----------------------------
def principal(phi_points, fzi_points, file_path):
    # Tabela FZI ↔ GHE
    fzi_values = [48, 24, 12, 6, 3, 1.5, 0.75, 0.375, 0.1875, 0.0938]
    ghe_labels = list(range(10, 0, -1))

    phi = np.linspace(0.01, 0.5, 300)
    colors = plt.cm.jet(np.linspace(0, 1, len(fzi_values)))

    fig, ax = plt.subplots(figsize=(9, 7))

    # Preencher faixas de GHE
    for i in range(len(fzi_values) - 1):
        k1 = permeability(phi, fzi_values[i])
        k2 = permeability(phi, fzi_values[i + 1])
        ax.fill_between(phi, k1, k2, color=colors[i], alpha=0.7, label=f"GHE {ghe_labels[i]}")

    # Pontos experimentais
    k_points = [permeability(phi_points[i], fzi_points[i]) for i in range(len(phi_points))]
    sc = ax.scatter(phi_points, k_points, color="black", s=50)

    ax.set_yscale("log")
    ax.set_xscale("linear")
    ax.set_xlabel("Porosity (decimal)")
    ax.set_ylabel("Permeability (mD)")
    ax.set_title("Global Hydraulic Elements (GHE)")
    ax.grid(True, which="both", ls="--", lw=0.5)

    # DataFrame para armazenar legendas clicadas
    df_clicks = pd.DataFrame(columns=["Phi", "FZI", "K"])

    # -----------------------------
    # Função para adicionar legenda ao clicar
    def on_click(event):
        nonlocal df_clicks
        if event.inaxes == ax:
            distances = np.sqrt((np.array(phi_points) - event.xdata)**2 +
                                (np.log10(np.array(k_points)) - np.log10(event.ydata))**2)
            min_index = np.argmin(distances)
            if distances[min_index] < 0.02:
                # Adicionar anotação no gráfico
                ax.annotate(f"Phi={phi_points[min_index]:.2f}\nFZI={fzi_points[min_index]}",
                            (phi_points[min_index], k_points[min_index]),
                            textcoords="offset points", xytext=(10,10),
                            arrowprops=dict(arrowstyle="->", color='red'),
                            fontsize=9, color='blue')
                fig.canvas.draw()

                # Adicionar ao DataFrame
                new_row = {
                    "Phi": phi_points[min_index],
                    "FZI": fzi_points[min_index],
                    "K": k_points[min_index]
                }
                df_clicks = pd.concat([df_clicks, pd.DataFrame([new_row])], ignore_index=True)

                # Salvar/atualizar Excel
                with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                    # Aba original
                    df = pd.DataFrame({
                        "Phi Points": phi_points,
                        "FZI Points": fzi_points,
                        "K Points": k_points
                    })
                    df.to_excel(writer, sheet_name='Planilha1', index=False)

                    # Aba de pontos clicados
                    df_clicks.to_excel(writer, sheet_name='Pontos_Clicados', index=False)

                    # Gráfico como imagem
                    img_file = "ghe_plot.png"
                    plt.savefig(img_file, dpi=300)
                    workbook = writer.book
                    worksheet2 = workbook.add_worksheet("Gráfico")
                    worksheet2.insert_image("B2", img_file, {"x_scale": 0.8, "y_scale": 0.8})

                print(f"✅ Ponto clicado registrado: Phi={phi_points[min_index]:.2f}, FZI={fzi_points[min_index]}")

    fig.canvas.mpl_connect("button_press_event", on_click)

    plt.show()  # Mostrar interativo

# -----------------------------
# Exemplo de uso
phi_points = [0.05, 0.12, 0.2, 0.15, 0.3, 0.6]
fzi_points = [3, 3, 4, 2, 5, 6]
principal(phi_points, fzi_points, "GHE_Excel.xlsx")
