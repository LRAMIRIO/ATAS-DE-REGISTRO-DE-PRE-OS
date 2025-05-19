# Recriar .zip após reset
from zipfile import ZipFile
import os

# Recriar conteúdo dos arquivos
os.makedirs("super_app/scripts", exist_ok=True)

# Conteúdo dos arquivos de script
scripts = {
    "transformar_pdf_para_xlsx.py": '''
import streamlit as st
import os
import pandas as pd
from PyPDF2 import PdfReader

def converter_pdf_para_excel(pdf_path, nome_saida=None):
    reader = PdfReader(pdf_path)
    linhas_total = []
    for page in reader.pages:
        texto = page.extract_text()
        if texto:
            linhas_total += texto.splitlines()
    df = pd.DataFrame(linhas_total, columns=["Linha"])
    if not nome_saida:
        base = os.path.basename(pdf_path).replace(".pdf", "")
        nome_saida = f"{base}_convertido.xlsx"
    df.to_excel(nome_saida, index=False)
    return nome_saida

def main():
    st.header("📄 Etapa 1: Converter PDFs para Planilhas XLSX")
    uploaded_files = st.file_uploader("📤 Envie um ou mais arquivos PDF", type="pdf", accept_multiple_files=True)
    if st.button("🔁 Converter"):
        for uploaded_file in uploaded_files:
            pdf_path = f"{uploaded_file.name}"
            with open(pdf_path, "wb") as f:
                f.write(uploaded_file.read())
            xlsx = converter_pdf_para_excel(pdf_path)
            st.success(f"✅ Convertido: {xlsx}")
            with open(xlsx, "rb") as f:
                st.download_button(f"📥 Baixar {xlsx}", f, file_name=xlsx)
''',

    "preencher_marcas.py": '''
import streamlit as st
import pandas as pd
import os

def main():
    st.header("🏷️ Etapa 2: Preencher Marcas por Item e Fornecedor")
    modelo = st.file_uploader("📤 Envie a planilha com fornecedores e itens preenchidos", type="xlsx")
    arquivos = st.file_uploader("📤 Envie os arquivos convertidos do PDF (.xlsx)", type="xlsx", accept_multiple_files=True)

    if st.button("🧠 Preencher Marcas"):
        if not modelo or not arquivos:
            st.warning("Envie os arquivos necessários.")
            return

        df_modelo = pd.read_excel(modelo, skiprows=2)
        df_modelo.columns = [col.strip() for col in df_modelo.columns]
        df_modelo["MARCA"] = ""

        for arq in arquivos:
            arq_path = arq.name
            with open(arq_path, "wb") as f:
                f.write(arq.read())
            df_arq = pd.read_excel(arq_path, header=None).astype(str)

            for i in range(len(df_arq) - 2):
                linha_atual = df_arq.iloc[i, 0].strip()
                proxima = df_arq.iloc[i+1, 0].strip()
                marca_linha = df_arq.iloc[i+2, 0].strip()

                if "Fornecedor" in linha_atual and "habilitado" in proxima and "Marca/Fabricante:" in marca_linha:
                    nome_fornecedor = linha_atual.split("-", 1)[-1].strip()
                    marca = marca_linha.split(":", 1)[-1].strip()
                    for j, row in df_modelo.iterrows():
                        if row["FORNECEDOR"].strip() == nome_fornecedor:
                            df_modelo.at[j, "MARCA"] = marca

        nome_saida = "planilha_com_marcas.xlsx"
        df_modelo.to_excel(nome_saida, index=False)
        st.success("✅ Marcas preenchidas!")
        with open(nome_saida, "rb") as f:
            st.download_button("📥 Baixar Planilha com Marcas", f, file_name=nome_saida)
''',

    "preencher_relatorio.py": '''
import streamlit as st
import fitz
import re
import pandas as pd
from openpyxl import load_workbook

def extrair_dados_pdf(caminho_pdf):
    padrao_geral = re.compile(
        r"Item (\\d+)[^\\n]*?\\n.*?"
        r"Aceito e Habilitado.*?para\\s+(.*?),\\s+CNPJ\\s+(\\d{2}\\.\\d{3}\\.\\d{3}/\\d{4}-\\d{2}),.*?"
        r"melhor lance:\\s*R\\$ ([\\d\\.,]+).*?/ R\\$ ([\\d\\.,]+)",
        re.DOTALL
    )
    dados = {}
    with fitz.open(caminho_pdf) as doc:
        for page in doc:
            texto = page.get_text()
            for match in padrao_geral.findall(texto):
                item = int(match[0])
                empresa = match[1].replace("\\n", " ").strip()
                valor_unit = float(match[3].replace(".", "").replace(",", "."))
                valor_total = float(match[4].replace(".", "").replace(",", "."))
                dados[item] = {
                    "fornecedor": empresa,
                    "valor_unitario": valor_unit,
                    "valor_total": valor_total
                }
    return dados

def main():
    st.header("📑 Etapa 3: Preencher Relatório de Julgamento")
    pdf = st.file_uploader("📤 Envie o arquivo PDF de julgamento", type="pdf")
    modelo = st.file_uploader("📤 Envie a planilha com marcas preenchidas", type="xlsx")

    if st.button("✅ Preencher Dados do Relatório"):
        if not pdf or not modelo:
            st.warning("Envie ambos os arquivos.")
            return

        with open(pdf.name, "wb") as f:
            f.write(pdf.read())

        dados = extrair_dados_pdf(pdf.name)
        wb = load_workbook(modelo)
        ws = wb.active

        for row in ws.iter_rows(min_row=4, max_col=9):
            try:
                item = int(str(row[0].value).replace(".", ""))
                if item in dados:
                    row[6].value = dados[item]["valor_unitario"]
                    row[7].value = dados[item]["valor_total"]
                    row[8].value = dados[item]["fornecedor"]
                else:
                    row[8].value = "Fracassado e/ou Deserto"
            except:
                row[8].value = "Fracassado e/ou Deserto"

        nome_saida = "planilha_com_relatorio.xlsx"
        wb.save(nome_saida)
        st.success("✅ Dados preenchidos com sucesso!")
        with open(nome_saida, "rb") as f:
            st.download_button("📥 Baixar Planilha Final", f, file_name=nome_saida)
''',

    "gerar_atas.py": '''
import streamlit as st
import pandas as pd
import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from zipfile import ZipFile

def main():
    st.header("📄 Etapa 4: Gerar Atas por Empresa (Excel e Word)")
    xlsx_file = st.file_uploader("📤 Envie a planilha final preenchida (.xlsx)", type="xlsx")

    if st.button("📁 Gerar Arquivos"):
        if not xlsx_file:
            st.warning("Envie a planilha preenchida.")
            return

        df = pd.read_excel(xlsx_file, skiprows=2)
        df.columns = [col.strip() for col in df.columns]
        output_dir = "outputs"
        os.makedirs(output_dir, exist_ok=True)

        borda_fina = Border(left=Side(style='thin'), right=Side(style='thin'),
                            top=Side(style='thin'), bottom=Side(style='thin'))

        for fornecedor, grupo in df.groupby("FORNECEDOR"):
            nome_limpo = fornecedor[:40].replace(" ", "_").replace("/", "-")
            grupo["ITEM"] = grupo["ITEM"].astype("Int64")
            grupo_formatado = grupo[["ITEM", "DESCRIÇÃO DO MATERIAL", "MARCA", "UNIDADE",
                                     "QUANTIDADE", "VALOR UNITÁRIO", "VALOR TOTAL"]]

            # Excel
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

            excel_path = f"{output_dir}/{nome_limpo}.xlsx"
            wb.save(excel_path)

            # Word
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

        # Compactar
        zip_path = "atas_geradas.zip"
        with ZipFile(zip_path, "w") as zipf:
            for f in os.listdir(output_dir):
                zipf.write(os.path.join(output_dir, f), arcname=f)

        with open(zip_path, "rb") as f:
            st.download_button("📥 Baixar Atas (.zip)", f, file_name=zip_path)
'''
}

# Criar scripts
for nome, conteudo in scripts.items():
    with open(f"super_app/scripts/{nome}", "w", encoding="utf-8") as f:
        f.write(conteudo.strip())

# Criar arquivos principais
with open("super_app/app_unificado_streamlit.py", "w", encoding="utf-8") as f:
    f.write("""
import streamlit as st
from scripts import transformar_pdf_para_xlsx, preencher_marcas, preencher_relatorio, gerar_atas

st.set_page_config(page_title="Super App Streamlit", layout="wide")
st.title("📄 Super App de Atas de Registro de Preços")

abas = st.tabs(["1️⃣ PDF ➜ XLSX", "2️⃣ Preencher Marcas", "3️⃣ Preencher Relatório", "4️⃣ Gerar Atas por Empresa"])

with abas[0]:
    transformar_pdf_para_xlsx.main()

with abas[1]:
    preencher_marcas.main()

with abas[2]:
    preencher_relatorio.main()

with abas[3]:
    gerar_atas.main()
""")

with open("super_app/requirements.txt", "w") as f:
    f.write("streamlit\npandas\nopenpyxl\nPyMuPDF\npython-docx")

with open("super_app/README.md", "w") as f:
    f.write("# Super App Streamlit para Atas de Registro de Preços\n\nEste app realiza o processo completo:\n- PDF ➜ XLSX\n- Preenche marcas\n- Preenche relatório\n- Gera atas por fornecedor\n\nExecute com:\n```bash\nstreamlit run app_unificado_streamlit.py\n```")

# Compactar
zip_path = "/mnt/data/super_app_streamlit_completo.zip"
with ZipFile(zip_path, "w") as zipf:
    for root, _, files in os.walk("super_app"):
        for file in files:
            full_path = os.path.join(root, file)
            arcname = os.path.relpath(full_path, "super_app")
            zipf.write(full_path, arcname=os.path.join("super_app", arcname))

zip_path
