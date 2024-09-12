import fitz  # PyMuPDF
import pandas as pd
import shutil
import os
import streamlit as st

# Função para obter o caminho da pasta Downloads do Windows
def get_downloads_folder():
    return os.path.join(os.path.expanduser("~"), "Downloads")

UPLOAD_FOLDER = get_downloads_folder()
processed_files = []  # Lista para armazenar caminhos dos arquivos processados

# Cria o diretório 'uploads' na pasta Downloads se não existir
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def process_files(excel_path, pdf_path):
    global processed_files
    processed_files = []  # Resetar lista de arquivos processados

    working_pdf = None  # Inicializar working_pdf aqui para garantir que ele esteja no escopo de finally

    try:
        # Carregar os dados do Excel
        df = pd.read_excel(excel_path, sheet_name='COMPROVANTES')
        nomes = df['NOME'].tolist()
        matricula = df['Matricula'].tolist()
        fixed_info = ['PROGEN S.A.', 'PRODUTO', 'DATA DE ENVIO:', 'RELATÓRIO ANALÍTICO', 'NOME', 'LOCAL DE ENTREGA:']
        pdf_document = pdf_path
        working_pdf = os.path.join(UPLOAD_FOLDER, 'working_arquivo.pdf')
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
                    output_pdf = os.path.join(UPLOAD_FOLDER, f'{nome}_VRHO.pdf')
                else:
                    output_pdf = os.path.join(UPLOAD_FOLDER, f'{nome}_AL.pdf')

                novo_pdf = fitz.open()

                for page_num in paginas_marcadas:
                    novo_pdf.insert_pdf(pdf, from_page=page_num, to_page=page_num)

                novo_pdf.save(output_pdf)
                novo_pdf.close()
                processed_files.append(output_pdf)  # Adicionar o arquivo processado à lista
            pdf.close()

        # Processar para cada nome
        for nome in nomes:
            marcar_e_salvar_pagina(nome)
            shutil.copy(pdf_document, working_pdf)
        
    finally:
        if working_pdf and os.path.exists(working_pdf):
            os.remove(working_pdf)

    return nomes

# Streamlit UI
st.title("Processamento de Arquivos")
st.write("Envie um arquivo Excel e um PDF para processamento:")

uploaded_excel = st.file_uploader("Escolha o arquivo Excel", type="xlsx")
uploaded_pdf = st.file_uploader("Escolha o arquivo PDF", type="pdf")

if st.button("Processar Arquivos"):
    if uploaded_excel and uploaded_pdf:
        excel_path = os.path.join(UPLOAD_FOLDER, uploaded_excel.name)
        pdf_path = os.path.join(UPLOAD_FOLDER, uploaded_pdf.name)
        
        # Salvar arquivos no diretório de Downloads
        with open(excel_path, "wb") as f:
            f.write(uploaded_excel.getbuffer())
        with open(pdf_path, "wb") as f:
            f.write(uploaded_pdf.getbuffer())
        
        try:
            nomes = process_files(excel_path, pdf_path)

            # Exibir links para os arquivos processados
            st.write("Arquivos processados:")
            for file in processed_files:
                with open(file, 'rb') as f:
                    st.download_button(label=os.path.basename(file), data=f.read(), file_name=os.path.basename(file))
        
        except Exception as e:
            st.error(f'Ocorreu um erro ao processar os arquivos: {str(e)}')
    else:
        st.warning("Por favor, faça o upload dos dois arquivos antes de processar.")

# Limpar arquivos processados
if st.button("Limpar Arquivos Processados"):
    for file_path in processed_files:
        if os.path.isfile(file_path):
            os.remove(file_path)
    processed_files.clear()
    st.write("Arquivos processados limpos.")
