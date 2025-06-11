import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
from openpyxl import Workbook
import re

def formatar_numero_brasileiro(valor):
    """Converte números no formato brasileiro (1.234,56) para float"""
    if isinstance(valor, str):
        valor = valor.strip()
        # Verifica se é um número no formato brasileiro (com vírgula decimal)
        if re.match(r'^-?\d{1,3}(?:\.\d{3})*,\d+$', valor):
            try:
                return float(valor.replace('.', '').replace(',', '.'))
            except:
                return valor
        # Verifica se já está no formato americano (ponto decimal)
        elif re.match(r'^-?\d+\.\d+$', valor):
            try:
                return float(valor)
            except:
                return valor
    return valor

def docx_para_planilha(docx_path, xlsx_path):
    try:
        doc = Document(docx_path)
        wb = Workbook()
        ws = wb.active

        linha = 1
        for tabela in doc.tables:
            for row in tabela.rows:
                colunas = [cell.text for cell in row.cells]  # Removido .strip() inicial
                for coluna, valor in enumerate(colunas, start=1):
                    # Remove espaços extras apenas no final da conversão
                    if isinstance(valor, str):
                        valor = valor.strip()
                    # Converte valores numéricos
                    valor_formatado = formatar_numero_brasileiro(valor)
                    ws.cell(row=linha, column=coluna, value=valor_formatado)
                linha += 1
            linha += 1  # Linha em branco entre tabelas (opcional)

        wb.save(xlsx_path)
        messagebox.showinfo('Sucesso', f'Tabela salva em:\n{xlsx_path}')
    except Exception as e:
        messagebox.showerror('Erro', f'Erro ao converter o arquivo:\n{str(e)}')

def selecionar_e_converter():
    docx_path = filedialog.askopenfilename(
        title="Selecione o arquivo .docx",
        filetypes=[("Documentos Word", "*.docx")]
    )
    if docx_path:
        xlsx_path = filedialog.asksaveasfilename(
            title="Salvar como",
            defaultextension=".xlsx",
            filetypes=[("Planilhas Excel", "*.xlsx")]
        )
        if xlsx_path:
            docx_para_planilha(docx_path, xlsx_path)

# Configuração da janela (descomente para usar)
# janela = tk.Tk()
# janela.title("Conversor de Tabela (.docx → .xlsx)")
# janela.geometry("400x200")
# botao = tk.Button(janela, text="Selecionar e Converter Tabela", command=selecionar_e_converter, font=("Arial", 12))
# botao.pack(pady=60)
# janela.mainloop()