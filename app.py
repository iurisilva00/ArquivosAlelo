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

    working_pdf = None  # Inicializar working_pdf aqui para garantir que ele esteja no escopo de finally

    try:
        # Carregar os dados do Excel
        df = pd.read_excel(excel_path, sheet_name='COMPROVANTES')
        nomes = df['NOME'].tolist()
        fixed_info = ['PROGEN S.A.', 'PRODUTO', 'DATA DE ENVIO:', 'RELATÓRIO ANALÍTICO', 'NOME', 'LOCAL DE ENTREGA:']
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

                    # Adicionar retângulos pretos para ocultar outros blocos de texto
                    text_blocks = page.get_text("blocks")
                    for block in text_blocks:
                        block_rect = fitz.Rect(block[:4])
                        if not any(rect.intersects(block_rect) for rect in highlight_rects) and not any(info in block[4] for info in fixed_info):
                            black_rect = page.add_rect_annot(block_rect)
                            black_rect.set_colors(stroke=(0, 0, 0), fill=(0, 0, 0))  # Preenchimento preto
                            black_rect.update()

                    paginas_marcadas.append(page_num)

            if paginas_marcadas:
                # Verificar se o nome do PDF contém "HOME"
                if "HOME" in pdf_document.upper():
                    output_pdf = os.path.join(download_folder, f'{nome}_VRHO.pdf')
                else:
                    output_pdf = os.path.join(download_folder, f'{nome}_AL.pdf')

                novo_pdf = fitz.open()

                for page_num in paginas_marcadas:
                    novo_pdf.insert_pdf(pdf, from_page=page_num, to_page=page_num)

                novo_pdf.save(output_pdf)
                novo_pdf.close()
                processed_files.append(output_pdf)
            pdf.close()

        # Processar para cada nome
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
st.title("Processamento de PDF e Excel para Arquivos Alelo")
st.write("""
    Faça upload de um arquivo Excel e um arquivo PDF. 
    Esse produto destina-se à medição. Com essa alternativa é possível gerar 
    todos os arquivos referentes às evidências do ticket **ALELO**. Seja o **VR** ou o **VR HOME OFFICE**.
""")


excel_file = st.file_uploader("Escolha o arquivo Excel (.xlsx)", type=['xlsx'])
st.write("O arquivo excel, precisa que a Planilha chame-se **COMPROVANTE**,  e que tenha colunas de **NOME** e **MATRICULA**")
pdf_file = st.file_uploader("Escolha o arquivo PDF (.pdf)", type=['pdf'])
st.write("Esse é o arquivo ALELO.")
if st.button("Executar"):
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
