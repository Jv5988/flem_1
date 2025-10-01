import os
import pandas as pd
from PyPDF2 import PdfReader

pasta_principal = r"C:\Users\ausna\VS CODE - CODS\py_flem_1" #Personalizar caminho

dados_por_pasta = {}

for root, dirs, files in os.walk(pasta_principal):
    nome_pasta = os.path.basename(root)
    
    if nome_pasta == os.path.basename(pasta_principal):
        continue

    registros = []
    for arquivo in files:
        if arquivo.lower().endswith(".pdf"):
            caminho = os.path.join(root, arquivo)

            nome_arquivo = os.path.splitext(arquivo)[0]
            partes = nome_arquivo.split("_")
            
            if len(partes) >= 2:
                nome = " ".join(partes[:-1]).upper()
                matricula = partes[-1].zfill(5)
            else:
                nome = nome_arquivo.upper()
                matricula = "00000"

            try:
                leitor = PdfReader(caminho)
                num_paginas = len(leitor.pages)
            except Exception:
                num_paginas = 0

            registros.append([nome, matricula, num_paginas])

    if registros:
        dados_por_pasta[nome_pasta] = pd.DataFrame(
            registros, columns=["Nome", "Matrícula", "N° Páginas"]
        )

saida = os.path.join(pasta_principal, "relatorio_funcionarios.xlsx")
with pd.ExcelWriter(saida, engine="openpyxl") as writer:
    for pasta, df in dados_por_pasta.items():
        aba = pasta[:31]
        df.to_excel(writer, sheet_name=aba, index=False)

print(f"✅ Planilha criada em: {saida}")
