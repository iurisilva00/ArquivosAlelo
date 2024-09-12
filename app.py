import fitz  # PyMuPDF
import pandas as pd
import shutil
import os
import streamlit as st
from zipfile import ZipFile

# Função para encontrar a pasta Downloads do Windows
def get_downloads_folder():
    return os.path.join(os.path.expanduser("~"), "Downloads")

# Função para processar os arquivos
def process_files(excel_path, pdf_path):
    download_folder = get_downloads_folder()
    processed_files = []  # Lista de arquivos processados

    try:
        # Carregar os dados do Excel
        df = pd.read_excel(excel_path, sheet_name='COMPROVANTES')
        nomes = df['NOME'].tolist()
        pdf_document = pdf_path

        working_pdf = os.path.join(download_folder, 'working_arquivo.pdf')
        shutil.copy(pdf_document, working_pdf)

        def marcar_e_salvar_pagina(nome):
            pdf = fitz.open(working_pdf)
            paginas_marcadas = []

            for page_num in range(len(pdf)):
                page = pdf[page_num]
                text = page.get_text("text")

                # Verificar se o texto da página contém o nome específico
                if nome in text:
                    areas = page.search_for(nome)
                    highlight_rects = []

                    for area in areas:
                        highlight_rects.append(area)

                    # Adicionar destaque amarelo nas áreas identificadas
                    for rect in highlight_rects:
                        highlight = page.add_highlight_annot(rect)
                        highlight.set_colors(stroke=None, fill=(1, 1, 0))  # Preenchimento amarelo
                        highlight.update()

                    paginas_marcadas.append(page_num)

            if paginas_marcadas:
                output_pdf = os.path.join(download_folder, f'{nome}_processed.pdf')

                novo_pdf = fitz.open()

                for page_num in paginas_marcadas:
                    novo_pdf.insert_pdf(pdf, from_page=page_num, to_page=page_num)

                novo_pdf.save(output_pdf)
                novo_pdf.close()
                processed_files.append(output_pdf)
            pdf.close()

        # Processar os nomes no Excel
        for nome in nomes:
            marcar_e_salvar_pagina(nome)
            shutil.copy(pdf_document, working_pdf)
        
    finally:
        if working_pdf and os.path.exists(working_pdf):
            os.remove(working_pdf)

    return processed_files

# Função para criar um arquivo zip contendo todos os arquivos processados
def zip_files(files, zip_name):
    zip_path = os.path.join(get_downloads_folder(), zip_name)
    with ZipFile(zip_path, 'w') as zipf:
        for file in files:
            zipf.write(file, os.path.basename(file))
    return zip_path

# Interface Streamlit
st.title("Processamento de PDF e Excel")
st.write("Faça upload de um arquivo Excel e um arquivo PDF.")

excel_file = st.file_uploader("Escolha o arquivo Excel (.xlsx)", type=['xlsx'])
pdf_file = st.file_uploader("Escolha o arquivo PDF (.pdf)", type=['pdf'])

if st.button("Processar Arquivos"):
    if excel_file and pdf_file:
        # Salvar os arquivos carregados na pasta Downloads
        download_folder = get_downloads_folder()
        excel_path = os.path.join(download_folder, excel_file.name)
        pdf_path = os.path.join(download_folder, pdf_file.name)

        with open(excel_path, "wb") as f:
            f.write(excel_file.getbuffer())
        with open(pdf_path, "wb") as f:
            f.write(pdf_file.getbuffer())

        # Processar os arquivos
        try:
            processed_files = process_files(excel_path, pdf_path)
            zip_name = "arquivos_processados.zip"
            zip_path = zip_files(processed_files, zip_name)

            st.success("Arquivos processados com sucesso!")
            st.write("Clique no link abaixo para baixar todos os arquivos processados como um arquivo ZIP.")
            st.download_button(label="Baixar arquivos ZIP", data=open(zip_path, "rb"), file_name=zip_name, mime="application/zip")
        
        except Exception as e:
            st.error(f"Erro ao processar os arquivos: {str(e)}")
    else:
        st.error("Por favor, faça o upload de ambos os arquivos Excel e PDF.")
