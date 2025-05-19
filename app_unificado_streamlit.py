# Reescrever o conteúdo correto do app_unificado_streamlit.py

conteudo_completo = """
import streamlit as st
import pandas as pd
import fitz
from PyPDF2 import PdfReader
from openpyxl import load_workbook
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl.styles import Alignment, Border, Side
import os
import zipfile

st.set_page_config(page_title="Super App - ARP", layout="wide")

st.title("📊 Super App para Geração de Atas de Registro de Preços")
menu = st.sidebar.radio("Etapas", ["1️⃣ PDF ➡️ Excel", "2️⃣ Preencher Marcas", "3️⃣ Preencher Valores", "4️⃣ Gerar Atas por Empresa"])

# Etapa 1 - PDF ➡️ Excel
if menu == "1️⃣ PDF ➡️ Excel":
    st.header("1️⃣ Converter PDFs em Excel (.xlsx)")
    uploaded_pdfs = st.file_uploader("📥 Envie um ou mais arquivos PDF", type=["pdf"], accept_multiple_files=True)

    if uploaded_pdfs:
        with st.spinner("Convertendo..."):
            output_paths = []
            for uploaded_file in uploaded_pdfs:
                pdf_path = f"/tmp/{uploaded_file.name}"
                with open(pdf_path, "wb") as f:
                    f.write(uploaded_file.read())
                reader = PdfReader(pdf_path)
                linhas = []
                for page in reader.pages:
                    texto = page.extract_text()
                    if texto:
                        linhas += texto.splitlines()
                df = pd.DataFrame(linhas, columns=["Linha"])
                xlsx_path = pdf_path.replace(".pdf", "_convertido.xlsx")
                df.to_excel(xlsx_path, index=False)
                output_paths.append(xlsx_path)
        st.success("✅ Conversão concluída.")
        with zipfile.ZipFile("/tmp/convertidos.zip", "w") as zipf:
            for p in output_paths:
                zipf.write(p, arcname=os.path.basename(p))
        with open("/tmp/convertidos.zip", "rb") as f:
            st.download_button("⬇️ Baixar ZIP com arquivos convertidos", f, file_name="pdfs_convertidos.zip")

# Etapa 2 - Preencher Marcas
elif menu == "2️⃣ Preencher Marcas":
    st.header("2️⃣ Preencher marcas dos fornecedores habilitados")
    planilhas_marca = st.file_uploader("📥 Envie os arquivos .xlsx convertidos do passo 1", type=["xlsx"], accept_multiple_files=True)
    planilha_base = st.file_uploader("📥 Envie a planilha modelo preenchida com os fornecedores", type="xlsx")

    if planilhas_marca and planilha_base:
        marcas_dict = {}

        for file in planilhas_marca:
            df = pd.read_excel(file)
            linhas = df["Linha"].astype(str).tolist()
            nome_empresa = ""
            item_num = None
            for i, linha in enumerate(linhas):
                if "Item" in linha and "-" in linha:
                    partes = linha.split("-")
                    if "Item" in partes[0]:
                        try:
                            item_num = int(partes[0].split("Item")[1].strip())
                        except:
                            item_num = None
                if "Fornecedor" in linha and i+1 < len(linhas) and "habilitado" in linhas[i+1].lower():
                    for j in range(i, -1, -1):
                        if "CNPJ" in linhas[j] and "-" in linhas[j]:
                            nome_empresa = linhas[j].split("-")[-1].strip()
                            break
                    for k in range(i, i+5):
                        if "Marca/Fabricante:" in linhas[k]:
                            marca = linhas[k].split(":", 1)[-1].strip()
                            if item_num and nome_empresa:
                                marcas_dict[(item_num, nome_empresa)] = marca
                            break

        df_base = pd.read_excel(planilha_base, skiprows=2)
        for idx, row in df_base.iterrows():
            chave = (int(row["ITEM"]), str(row["FORNECEDOR"]).strip())
            if chave in marcas_dict:
                df_base.at[idx, "MARCA"] = marcas_dict[chave]

        saida_marca = "/tmp/planilha_com_marcas.xlsx"
        df_base.to_excel(saida_marca, index=False)
        with open(saida_marca, "rb") as f:
            st.download_button("⬇️ Baixar planilha com MARCAS preenchidas", f, file_name="planilha_com_marcas.xlsx")

# Etapa 3 - Preencher Valores
elif menu == "3️⃣ Preencher Valores":
    st.header("3️⃣ Preencher VALORES e FORNECEDORES a partir do PDF de julgamento")
    pdf_julgamento = st.file_uploader("📥 Envie o PDF do julgamento (consolidado)", type="pdf")
    planilha_com_marcas = st.file_uploader("📥 Envie a planilha preenchida com marcas", type="xlsx")

    if pdf_julgamento and planilha_com_marcas:
        import re
        pdf_path = f"/tmp/{pdf_julgamento.name}"
        with open(pdf_path, "wb") as f:
            f.write(pdf_julgamento.read())

        padrao = re.compile(
            r"Item (\d+)[^\n]*?\n.*?"
            r"Aceito e Habilitado.*?para\s+(.*?),\s+CNPJ\s+(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}),.*?"
            r"melhor lance:\s*R\\$ ([\d\\.,]+).*?/ R\\$ ([\d\\.,]+)",
            re.DOTALL
        )

        dados = {}
        with fitz.open(pdf_path) as doc:
            for page in doc:
                texto = page.get_text()
                for match in padrao.findall(texto):
                    item = int(match[0])
                    empresa = match[1].replace("\n", " ").strip()
                    valor_unit = float(match[3].replace(".", "").replace(",", "."))
                    valor_total = float(match[4].replace(".", "").replace(",", "."))
                    dados[item] = {
                        "fornecedor": empresa,
                        "valor_unitario": valor_unit,
                        "valor_total": valor_total
                    }

        wb = load_workbook(planilha_com_marcas)
        ws = wb.active
        for row in ws.iter_rows(min_row=4, max_col=9):
            try:
                item_val = row[0].value
                item_num = int(str(item_val).replace(".", "").strip())
                if item_num in dados:
                    row[6].value = dados[item_num]["valor_unitario"]
                    row[7].value = dados[item_num]["valor_total"]
                    row[8].value = dados[item_num]["fornecedor"]
                else:
                    row[8].value = "Fracassado e/ou Deserto"
            except:
                row[8].value = "Fracassado e/ou Deserto"

        saida_valores = "/tmp/planilha_completa.xlsx"
        wb.save(saida_valores)
        with open(saida_valores, "rb") as f:
            st.download_button("⬇️ Baixar planilha com VALORES e FORNECEDORES", f, file_name="planilha_completa.xlsx")

# Etapa 4 - Gerar Atas
elif menu == "4️⃣ Gerar Atas por Empresa":
    st.header("4️⃣ Gerar atas de registro de preços por empresa")
    planilha_final = st.file_uploader("📥 Envie a planilha final com todas as colunas preenchidas", type="xlsx")

    if planilha_final:
        df = pd.read_excel(planilha_final, skiprows=2)
        df.columns = [col.strip() for col in df.columns]

        output_dir = "/tmp/atas"
        os.makedirs(output_dir, exist_ok=True)

        for fornecedor, grupo in df.groupby("FORNECEDOR"):
            nome_limpo = fornecedor[:40].replace(" ", "_").replace("/", "-")
            grupo["ITEM"] = grupo["ITEM"].astype("Int64")
            grupo_formatado = grupo[[
                "ITEM", "DESCRIÇÃO DO MATERIAL", "MARCA", "UNIDADE",
                "QUANTIDADE", "VALOR UNITÁRIO", "VALOR TOTAL"
            ]]

            wb = Workbook()
            ws = wb.active
            ws.title = "Itens Vencedores"
            ws.append(grupo_formatado.columns.tolist())
            for _, row in grupo_formatado.iterrows():
                ws.append(row.tolist())

            for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
                for cell in row:
                    if cell.row == 1:
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    elif cell.column_letter == "B":
                        cell.alignment = Alignment(horizontal="justify", vertical="center", wrap_text=True)
                    else:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                         top=Side(style='thin'), bottom=Side(style='thin'))

            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                row[5].number_format = 'R$ #,##0.00'
                row[6].number_format = 'R$ #,##0.00'

            ws.column_dimensions["A"].width = 6
            ws.column_dimensions["B"].width = 65
            ws.column_dimensions["C"].width = 20
            ws.column_dimensions["D"].width = 20
            ws.column_dimensions["E"].width = 12
            ws.column_dimensions["F"].width = 16
            ws.column_dimensions["G"].width = 16

            excel_path = f"{output_dir}/{nome_limpo}.xlsx"
            wb.save(excel_path)

            doc = Document()
            doc.add_heading(f'Fornecedor: {fornecedor}', 0)
            table = doc.add_table(rows=1, cols=len(grupo_formatado.columns))
            table.style = 'Table Grid'
            table.alignment = WD_TABLE_ALIGNMENT.CENTER

            hdr_cells = table.rows[0].cells
            for i, col in enumerate(grupo_formatado.columns):
                hdr_cells[i].text = str(col)
                hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

            for _, row in grupo_formatado.iterrows():
                row_cells = table.add_row().cells
                for i, value in enumerate(row):
                    texto = str(value).replace(" ", "\\n") if grupo_formatado.columns[i] == "UNIDADE" else str(value)
                    row_cells[i].text = texto
                    row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

            docx_path = f"{output_dir}/{nome_limpo}.docx"
            doc.save(docx_path)

        zip_path = "/tmp/atas_por_empresa.zip"
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for f in os.listdir(output_dir):
                zipf.write(os.path.join(output_dir, f), arcname=f)

        with open(zip_path, "rb") as f:
            st.download_button("⬇️ Baixar ZIP com Atas", f, file_name="atas_por_empresa.zip")
"""

# Salvar no caminho correto para download
path_final_script = "/mnt/data/app_unificado_streamlit.py"
with open(path_final_script, "w", encoding="utf-8") as f:
    f.write(conteudo_completo)

path_final_script
