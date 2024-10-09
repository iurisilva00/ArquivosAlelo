import fitz  # PyMuPDF
import pandas as pd
import shutil
import os
import streamlit as st
from zipfile import ZipFile
from io import BytesIO
import re

# Função para encontrar a pasta Downloads do Windows
def get_downloads_folder():
    return os.path.join(os.path.expanduser("~"), "Downloads")

# Função para processar os arquivos e salvá-los diretamente no ZIP
def process_files_and_zip(excel_file, pdf_file):
    download_folder = get_downloads_folder()
    processed_files = []  # Lista de arquivos processados
    zip_buffer = BytesIO()  # Usar BytesIO para criar o zip em memória

    if excel_file is None or pdf_file is None:
        raise ValueError("Os arquivos Excel e PDF devem ser fornecidos.")

    working_pdf = None  # Inicializar working_pdf aqui para garantir que ele esteja no escopo de finally

    try:
        # Carregar os dados do Excel
        df = pd.read_excel(excel_file, sheet_name='COMPROVANTES')
        nomes = df['NOME'].tolist()
        matriculas = df['MATRICULA'].astype(str).tolist()  # Converter as matrículas para string
        fixed_info = ['PROGEN S.A.', 'PRODUTO', 'DATA DE ENVIO:', 'RELATÓRIO ANALÍTICO', 'NOME', 'LOCAL DE ENTREGA:']
        
        # Ler o conteúdo do arquivo PDF em BytesIO
        working_pdf = BytesIO(pdf_file.read())
        if not working_pdf.getvalue():
            raise ValueError("O arquivo PDF está vazio.")

        # Verificar se o nome do PDF contém a palavra "HOME" usando regex (case-insensitive)
        pdf_name = pdf_file.name
        if re.search(r'home', pdf_name, re.IGNORECASE):
            sufixo = '_VRHO'
        else:
            sufixo = '_AL'

        selected_data = []  # Lista para armazenar as matrículas e nomes selecionados

        with ZipFile(zip_buffer, 'w') as zipf:
            def marcar_e_salvar_pagina(nome, matricula):
                pdf = fitz.open(stream=working_pdf.getvalue())
                paginas_marcadas = []

                for page_num in range(len(pdf)):
                    page = pdf[page_num]
                    text = page.get_text("text")

                    # Verificar se o texto da página contém a matrícula específica usando igualdade
                    pattern = rf'\b{matricula}\b'
                    if re.search(pattern, text):
                        areas = page.search_for(matricula)
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
                            if not any(rect.intersects(block_rect) for rect in highlight_rects) and not any(info in block for info in fixed_info):
                                black_rect = page.add_rect_annot(block_rect)
                                black_rect.set_colors(stroke=(0, 0, 0), fill=(0, 0, 0))  # Preenchimento preto
                                black_rect.update()

                        paginas_marcadas.append(page_num)

                if paginas_marcadas:
                    output_pdf_name = f'{nome}{sufixo}.pdf'

                    novo_pdf = fitz.open()

                    for page_num in paginas_marcadas:
                        novo_pdf.insert_pdf(pdf, from_page=page_num, to_page=page_num)

                    # Salvar diretamente no ZIP
                    output_pdf_bytes = BytesIO()
                    novo_pdf.save(output_pdf_bytes)
                    novo_pdf.close()

                    # Escrever no ZIP sem salvar no sistema de arquivos
                    zipf.writestr(output_pdf_name, output_pdf_bytes.getvalue())

                    # Adicionar a matrícula e o nome à lista de dados selecionados
                    selected_data.append({'MATRICULA': matricula, 'NOME': nome})

                pdf.close()

            # Processar para cada nome e matrícula
            for nome, matricula in zip(nomes, matriculas):
                marcar_e_salvar_pagina(nome, matricula)

            # Criar um DataFrame com os dados selecionados e salvar como Excel no ZIP
            selected_df = pd.DataFrame(selected_data)
            excel_buffer = BytesIO()
            selected_df.to_excel(excel_buffer, index=False)
            excel_buffer.seek(0)
            zipf.writestr('dados_selecionados.xlsx', excel_buffer.getvalue())
        
    finally:
        if working_pdf:
            working_pdf.close()

    return zip_buffer

# Função para resetar a interface e variáveis após o download ou ao clicar em limpar campos
def reset_state():
    st.session_state.clear()

# Interface Streamlit
st.title("Processamento de PDF e Excel para Arquivos Alelo")
st.write("""
    Faça upload de um arquivo Excel e um arquivo PDF. 
    Esse produto destina-se à medição. Com essa alternativa é possível gerar 
    todos os arquivos referentes às evidências do ticket **ALELO**. Seja o **VR** ou o **VR HOME OFFICE**.
""")

# Campos de upload de arquivos
st.write("""
O arquivo execel deve chamar-se VR, o nome a primeira Planilha dentro do arquivo deve ser COMPROVANTES. Em COMPROVANTES preciso adicionar duas colunas NOME e MATRICULA
""")
excel_file = st.file_uploader("Escolha o arquivo Excel (.xlsx)", type=['xlsx'])
pdf_file = st.file_uploader("Escolha o arquivo PDF (.pdf)", type=['pdf'])

if st.button("Executar"):
    if excel_file and pdf_file:
        try:
            # Processar os arquivos e gerar o ZIP
            zip_buffer = process_files_and_zip(excel_file, pdf_file)
            st.success("Arquivos processados com sucesso!")
            st.write("Clique no link abaixo para baixar todos os arquivos processados como um arquivo ZIP.")
            st.download_button(
                label="Baixar arquivos ZIP", 
                data=zip_buffer.getvalue(), 
                file_name="arquivos_processados.zip", 
                mime="application/zip", 
                on_click=reset_state  # Limpa as variáveis após o download
            )

        except Exception as e:
            st.error(f"Erro ao processar os arquivos: {str(e)}")
    else:
        st.error("Por favor, faça o upload de ambos os arquivos Excel e PDF.")

# Adicionar botão para limpar campos e resetar estado
if st.button("Limpar Campos"):
    reset_state()
