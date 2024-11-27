import fitz  # PyMuPDF
import pandas as pd
import shutil
import os
import streamlit as st
from zipfile import ZipFile
from io import BytesIO
import re

def get_downloads_folder():
    return os.path.join(os.path.expanduser("~"), "Downloads")

def process_files_and_zip(excel_file, pdf_file):
    download_folder = get_downloads_folder()
    processed_files = []  
    zip_buffer = BytesIO()  

    if excel_file is None or pdf_file is None:
        raise ValueError("Os arquivos Excel e PDF devem ser fornecidos.")

    working_pdf = None  

    try:
        # Ler o Excel e preparar dados
        df = pd.read_excel(excel_file, sheet_name='COMPROVANTES')
        nomes = df['NOME'].tolist()
        matriculas = df['MATRICULA'].astype(str).tolist()  
        fixed_info = ['PROGEN S.A.', 'PRODUTO', 'DATA DE ENVIO:', 'RELATÓRIO ANALÍTICO', 
                      'NOME', 'LOCAL DE ENTREGA:', "CPF", "NASCIMENTO", "MATRICULA", "VL BENEFICIO"]
        
        # Ler o arquivo PDF
        working_pdf = BytesIO(pdf_file.read())
        if not working_pdf.getvalue():
            raise ValueError("O arquivo PDF está vazio.")

        pdf_name = pdf_file.name
        sufixo = '_VRHO' if re.search(r'home', pdf_name, re.IGNORECASE) else '_AL'

        selected_data = []  

        with ZipFile(zip_buffer, 'w') as zipf:
            # Função interna para processar e marcar PDFs
            def marcar_e_salvar_pagina(nome, matricula):
                pdf = fitz.open(stream=working_pdf.getvalue())
                paginas_marcadas = []

                for page_num in range(len(pdf)):
                    page = pdf[page_num]
                    text = page.get_text("text")

                    # Procurar pela matrícula na página
                    pattern = rf'\b{matricula}\b'
                    if re.search(pattern, text):
                        areas = page.search_for(matricula)
                        highlight_rects = []

                        for area in areas:
                            highlight_rects.append(area)

                        # Adicionar destaques amarelos nas áreas encontradas
                        for rect in highlight_rects:
                            highlight = page.add_highlight_annot(rect)
                            highlight.set_colors(stroke=None, fill=(1, 1, 0))  
                            highlight.update()

                        # Adicionar retângulos pretos para ocultar informações
                        text_blocks = page.get_text("blocks")
                        for block in text_blocks:
                            block_rect = fitz.Rect(block[:4])
                            if not any(rect.intersects(block_rect) for rect in highlight_rects) and \
                               not any(info in block[4] for info in fixed_info):  # Comparação com conteúdo do bloco
                                black_rect = page.add_rect_annot(block_rect)
                                black_rect.set_colors(stroke=(0, 0, 0), fill=(0, 0, 0))  
                                black_rect.update()

                        paginas_marcadas.append(page_num)

                if paginas_marcadas:
                    output_pdf_name = f'{nome}{sufixo}.pdf'

                    novo_pdf = fitz.open()

                    for page_num in paginas_marcadas:
                        novo_pdf.insert_pdf(pdf, from_page=page_num, to_page=page_num)

                    # Salvar PDF com proteção
                    output_pdf_bytes = BytesIO()
                    novo_pdf.save(
                        output_pdf_bytes,
                        encryption=fitz.PDF_ENCRYPT_AES_256,  # Criptografia AES-256
                        permissions=fitz.PDF_PERM_PRINT,     # Permitir somente impressão
                        owner_pw="senha_segura"             # Senha de proprietário
                    )
                    novo_pdf.close()

                    # Adicionar ao ZIP
                    zipf.writestr(output_pdf_name, output_pdf_bytes.getvalue())
                    selected_data.append({'MATRICULA': matricula, 'NOME': nome})

                pdf.close()

            # Processar cada nome e matrícula
            for nome, matricula in zip(nomes, matriculas):
                marcar_e_salvar_pagina(nome, matricula)

            # Criar arquivo Excel com dados selecionados
            selected_df = pd.DataFrame(selected_data)
            excel_buffer = BytesIO()
            selected_df.to_excel(excel_buffer, index=False)
            excel_buffer.seek(0)
            zipf.writestr('dados_selecionados.xlsx', excel_buffer.getvalue())
        
    finally:
        if working_pdf:
            working_pdf.close()

    return zip_buffer

def reset_state():
    st.session_state.clear()

# Interface Streamlit
st.title("Processamento de PDF e Excel para Arquivos Alelo")
st.write("""
    Faça upload de um arquivo Excel e um arquivo PDF. 
    Esse produto destina-se à medição. Com essa alternativa é possível gerar 
    todos os arquivos referentes às evidências do ticket **ALELO**. Seja o **VR** ou o **VR HOME OFFICE**.
""")

st.write("""
O arquivo Excel deve chamar-se VR, e a primeira planilha dentro do arquivo deve ser 'COMPROVANTES'. 
Na planilha 'COMPROVANTES', adicione duas colunas: NOME e MATRICULA.
""")
excel_file = st.file_uploader("Escolha o arquivo Excel (.xlsx)", type=['xlsx'])
pdf_file = st.file_uploader("Escolha o arquivo PDF (.pdf)", type=['pdf'])

if st.button("Executar"):
    if excel_file and pdf_file:
        try:
            zip_buffer = process_files_and_zip(excel_file, pdf_file)
            st.success("Arquivos processados com sucesso!")
            st.write("Clique no link abaixo para baixar todos os arquivos processados como um arquivo ZIP.")
            st.download_button(
                label="Baixar arquivos ZIP", 
                data=zip_buffer.getvalue(), 
                file_name="arquivos_processados.zip", 
                mime="application/zip", 
                on_click=reset_state  
            )

        except Exception as e:
            st.error(f"Erro ao processar os arquivos: {str(e)}")
    else:
        st.error("Por favor, faça o upload de ambos os arquivos Excel e PDF.")

if st.button("Limpar Campos"):
    reset_state()