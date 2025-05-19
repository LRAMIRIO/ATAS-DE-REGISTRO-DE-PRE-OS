# Criar o conteúdo completo do app_unificado_streamlit.py com as 4 etapas

conteudo_app = '''
import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO
import tempfile
from zipfile import ZipFile
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Side
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PyPDF2 import PdfReader
import fitz  # PyMuPDF

st.set_page_config(page_title="Super App de Licitação", layout="wide")
st.title("📑 Super App para Processamento de Relatórios de Licitação")

menu = st.sidebar.radio("Etapas do processo", [
    "1️⃣ Converter PDF para XLSX",
    "2️⃣ Preencher Marcas",
    "3️⃣ Preencher Fornecedores e Valores",
    "4️⃣ Gerar Atas por Fornecedor"
])

# Funções auxiliares
def converter_pdf_para_excel(pdf_file):
    reader = PdfReader(pdf_file)
    linhas_total = []
    for page in reader.pages:
        texto = page.extract_text()
        if texto:
            linhas_total += texto.splitlines()
    df = pd.DataFrame(linhas_total, columns=["Linha"])
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    return output

def extrair_marcas(files, planilha_modelo):
    marcas_por_item = {}
    for nome_arquivo, file in files.items():
        if "convertido" in nome_arquivo:
            df = pd.read_excel(file)
            linhas = df["Linha"].astype(str).tolist()
            item_match = re.search(r"item-(\\d+)", nome_arquivo)
            item = int(item_match.group(1)) if item_match else None
            for i in range(len(linhas) - 2):
                if "(total)" in linhas[i].lower() and "fornecedor" in linhas[i].lower() and "habilitado" in linhas[i+1].lower():
                    if "marca/fabricante:" in linhas[i+2].lower():
                        marca = linhas[i+2].split(":", 1)[-1].strip()
                        if item:
                            marcas_por_item[item] = marca
                        break
    wb = load_workbook(planilha_modelo)
    ws = wb.active
    for row in ws.iter_rows(min_row=4):
        item_val = row[0].value
        if item_val:
            item = int(str(item_val).replace(".", "").strip())
            if item in marcas_por_item:
                row[3].value = marcas_por_item[item]
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def preencher_relatorio_pdf(lista_pdfs, planilha_base):
    dados = {}
    padrao = re.compile(
        r"Item (\\d+)[^\\n]*?\\n.*?"
        r"Aceito e Habilitado.*?para\\s+(.*?),\\s+CNPJ\\s+(\\d{2}\\.\\d{3}\\.\\d{3}/\\d{4}-\\d{2}),.*?"
        r"melhor lance:\\s*R\\$ ([\\d\\.,]+).*?/ R\\$ ([\\d\\.,]+)",
        re.DOTALL
    )
    for pdf in lista_pdfs:
        with fitz.open(stream=pdf.read(), filetype="pdf") as doc:
            for page in doc:
                texto = page.get_text()
                for match in padrao.findall(texto):
                    item = int(match[0])
                    empresa = match[1].replace("\\n", " ").strip()
                    valor_unit = float(match[3].replace(".", "").replace(",", "."))
                    valor_total = float(match[4].replace(".", "").replace(",", "."))
                    dados[item] = {
                        "fornecedor": empresa,
                        "valor_unitario": valor_unit,
                        "valor_total": valor_total
                    }
    wb = load_workbook(planilha_base)
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
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def gerar_atas(planilha_final):
    df = pd.read_excel(planilha_final, skiprows=2)
    df.columns = [col.strip() for col in df.columns]
    zip_buffer = BytesIO()
    borda_fina = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    with tempfile.TemporaryDirectory() as temp_dir:
        for fornecedor, grupo in df.groupby("FORNECEDOR"):
            nome_limpo = fornecedor[:40].replace(" ", "_").replace("/", "-")
            grupo["ITEM"] = grupo["ITEM"].astype("Int64")
            grupo_fmt = grupo[["ITEM", "DESCRIÇÃO DO MATERIAL", "MARCA", "UNIDADE", "QUANTIDADE", "VALOR UNITÁRIO", "VALOR TOTAL"]]
            # Excel
            wb = Workbook()
            ws = wb.active
            ws.title = "Itens Vencedores"
            ws.append(grupo_fmt.columns.tolist())
            for _, row in grupo_fmt.iterrows():
                ws.append(row.tolist())
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
                for cell in row:
                    if cell.row == 1:
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    elif cell.column_letter == "B":
                        cell.alignment = Alignment(horizontal="justify", vertical="center", wrap_text=True)
                    else:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = borda_fina
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
            path_excel = os.path.join(temp_dir, f"{nome_limpo}.xlsx")
            wb.save(path_excel)

            # Word
            doc = Document()
            doc.add_heading(f'Fornecedor: {fornecedor}', 0)
            table = doc.add_table(rows=1, cols=len(grupo_fmt.columns))
            table.style = 'Table Grid'
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            hdr_cells = table.rows[0].cells
            for i, col in enumerate(grupo_fmt.columns):
                hdr_cells[i].text = str(col)
                hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            for _, row in grupo_fmt.iterrows():
                row_cells = table.add_row().cells
                for i, value in enumerate(row):
                    texto = str(value).replace(" ", "\\n") if grupo_fmt.columns[i] == "UNIDADE" else str(value)
                    row_cells[i].text = texto
                    row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            path_doc = os.path.join(temp_dir, f"{nome_limpo}.docx")
            doc.save(path_doc)

        with ZipFile(zip_buffer, "w") as zipf:
            for f in os.listdir(temp_dir):
                zipf.write(os.path.join(temp_dir, f), arcname=f)
    return zip_buffer.getvalue()

# Interface para cada etapa
if menu == "1️⃣ Converter PDF para XLSX":
    uploaded_pdfs = st.file_uploader("📂 Envie arquivos PDF:", type="pdf", accept_multiple_files=True)
    if uploaded_pdfs:
        for pdf in uploaded_pdfs:
            output = converter_pdf_para_excel(pdf)
            nome_saida = pdf.name.replace(".pdf", "_convertido.xlsx")
            st.download_button(f"📥 Baixar {nome_saida}", output, nome_saida, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif menu == "2️⃣ Preencher Marcas":
    st.markdown("Envie a planilha modelo e os arquivos `.xlsx` convertidos dos PDFs.")
    modelo = st.file_uploader("🧾 Planilha modelo (.xlsx)", type="xlsx")
    convertidos = st.file_uploader("📂 Arquivos .xlsx convertidos:", type="xlsx", accept_multiple_files=True)
    if modelo and convertidos:
        arquivos_dict = {file.name: file for file in convertidos}
        output = extrair_marcas(arquivos_dict, modelo)
        st.download_button("📥 Baixar planilha com marcas", output, file_name="Planilha_com_marcas.xlsx")

elif menu == "3️⃣ Preencher Fornecedores e Valores":
    planilha_base = st.file_uploader("🧾 Planilha base com marcas (.xlsx)", type="xlsx")
    relatorios = st.file_uploader("📂 Relatórios PDF de julgamento:", type="pdf", accept_multiple_files=True)
    if planilha_base and relatorios:
        saida = preencher_relatorio_pdf(relatorios, planilha_base)
        st.download_button("📥 Baixar planilha completa", saida, file_name="Planilha_completa.xlsx")

elif menu == "4️⃣ Gerar Atas por Fornecedor":
    planilha_final = st.file_uploader("🧾 Envie a planilha completa (.xlsx):", type="xlsx")
    if planilha_final:
        zip_bytes = gerar_atas(planilha_final)
        st.download_button("📦 Baixar .zip com atas", zip_bytes, file_name="atas_por_fornecedor.zip", mime="application/zip")
'''

# Escrever o app
with open("/mnt/data/app_unificado_streamlit.py", "w", encoding="utf-8") as f:
    f.write(conteudo_app)

# Atualizar o .zip
zip_path = "/mnt/data/super_app_streamlit.zip"
with zipfile.ZipFile(zip_path, "w") as zipf:
    zipf.write("/mnt/data/app_unificado_streamlit.py", arcname="app_unificado_streamlit.py")
    zipf.write("/mnt/data/requirements.txt", arcname="requirements.txt")
    zipf.write("/mnt/data/README.txt", arcname="README.txt")

zip_path
