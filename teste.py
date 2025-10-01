import pandas as pd

# -------------------------
# Exemplo de dados
# -------------------------
data = {
    "Porosidade (%)": [5.3, 12, 20.5, 15, 30,60],
    "Permeabilidade (mD)": [0.035, 1.2, 8.7, 2.1, 50,20],
    "GHE": [3, 3, 4, 5, 4,6]
}
df = pd.DataFrame(data)

# -------------------------
# Salvar no Excel com gráfico
# -------------------------
output_file = "grafico_ghe.xlsx"

with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    df.to_excel(writer, sheet_name="Dados", index=False)

    # Pegar workbook e worksheet
    workbook  = writer.book
    worksheet = writer.sheets["Dados"]

    # Criar gráfico de dispersão (só pontos, sem linhas)
    chart = workbook.add_chart({"type": "scatter"})

    # Pegar lista de GHE únicos
    unique_ghe = df["GHE"].unique()

    # Adicionar cada GHE como série no gráfico
    for ghe in unique_ghe:
        ghe_data = df[df["GHE"] == ghe]
        start_row = ghe_data.index.min() + 1  # +1 por causa do header
        end_row   = ghe_data.index.max() + 1

        chart.add_series({
            "name":       f"GHE {ghe}",
            "categories": ["Dados", start_row, 0, end_row, 0],  # Porosidade (%)
            "values":     ["Dados", start_row, 1, end_row, 1],  # Permeabilidade (mD)
            "marker": {"type": "circle", "size": 7},
        })

    # Eixos
    chart.set_x_axis({"name": "Porosidade (%)", "label_position": "low"})
    chart.set_y_axis({
        "name": "Permeabilidade (mD)",
        "log_base": 10,
        "min": 0.01,
        "max": 10000,
    })

    chart.set_title({"name": "Gráfico por GHE"})
    chart.set_legend({"position": "bottom"})

    # Inserir gráfico na planilha
    worksheet.insert_chart("E2", chart)

print(f"Arquivo Excel criado: {output_file}")
