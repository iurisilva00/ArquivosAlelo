
import fitz  # PyMuPDF
import pandas as pd
import os
import streamlit as st
from zipfile import ZipFile
from io import BytesIO
import re
import streamlit.components.v1 as components

# 🔄 Sempre reseta o estado ao entrar na página
@st.cache_resource
def reset_on_load():
    st.session_state.clear()
    return True

reset_on_load()  # 🔄 Garante que a sessão sempre começa limpa

def get_downloads_folder():
    return os.path.join(os.path.expanduser("~"), "Downloads")

def process_files_and_zip(excel_file, pdf_file):
    download_folder = get_downloads_folder()
    zip_buffer = BytesIO()

    if excel_file is None or pdf_file is None:
        raise ValueError("Os arquivos Excel e PDF devem ser fornecidos.")

    working_pdf = None  

    try:
        # Ler o Excel e preparar dados
        df = pd.read_excel(excel_file, sheet_name='COMPROVANTES')
        df.columns = df.columns.str.strip()
        nomes = df['NOME'].tolist()
        matriculas = df['MATRICULA'].astype(str).tolist()  

        # 🔹 Se a linha contém essas palavras, não será pintada
        fixed_info = ['PROGEN S.A.', 'PRODUTO', 'DATA DE ENVIO:', 'RELATÓRIO ANALÍTICO', 
                      'NOME', 'LOCAL DE ENTREGA:', "CPF", "MATRICULA", "VL BENEFICIO"]
        # Ler o arquivo PDF
        working_pdf = BytesIO(pdf_file.read())
        if not working_pdf.getvalue():
            raise ValueError("O arquivo PDF está vazio.")

        pdf_name = pdf_file.name
        pdf_document = fitz.open(stream=working_pdf.getvalue())

        # 🔹 Verifica se o PDF contém "HOME" → Define o sufixo do nome do arquivo
        pdf_text = "\n".join([pdf_document[page].get_text("text") for page in range(len(pdf_document))])
        sufixo = '_VRHO' if re.search(r'home', pdf_text, re.IGNORECASE) else '_AL'

        selected_data = []

        with ZipFile(zip_buffer, 'w') as zipf:
            def marcar_e_salvar_pagina(nome, matricula):
                pdf = fitz.open(stream=working_pdf.getvalue())
                paginas_marcadas = []

                for page_num in range(len(pdf)):
                    page = pdf[page_num]
                    text = page.get_text("text")

                    # 🔹 Marcar apenas a matrícula pesquisada
                    pattern = rf'\b{matricula}\b'
                    if re.search(pattern, text):
                        areas = page.search_for(matricula)
                        highlight_rects = []

                        for area in areas:
                            highlight_rects.append(area)

                        # 🔹 Marcar toda a linha onde a matrícula aparece
                        text_blocks = page.get_text("blocks")
                        for block in text_blocks:
                            block_rect = fitz.Rect(block[:4])
                            block_text = block[4]

                            if re.search(pattern, block_text):
                                highlight_rects.append(block_rect)

                        # 🔹 Aplicar destaques
                        for rect in highlight_rects:
                            highlight = page.add_highlight_annot(rect)
                            highlight.set_colors(stroke=None, fill=(1, 1, 0))  
                            highlight.update()

                        # 🔹 Ocultar informações sensíveis, **exceto** quando a linha tem "NOME", "CPF" ou "MATRICULA"
                        for block in text_blocks:
                            block_rect = fitz.Rect(block[:4])
                            block_text = block[4]

                            # 🔹 Se a linha contém "NOME", "CPF" ou "MATRICULA", **não será pintada**
                            if any(info in block_text for info in fixed_info):
                                continue

                            # 🔹 Se não for a linha da matrícula, oculta normalmente
                            if not any(rect.intersects(block_rect) for rect in highlight_rects):
                                black_rect = page.add_rect_annot(block_rect)
                                black_rect.set_colors(stroke=(0, 0, 0), fill=(0, 0, 0))  
                                black_rect.update()

                        paginas_marcadas.append(page_num)

                if paginas_marcadas:
                    output_pdf_name = f'{nome}{sufixo}.pdf'
                    novo_pdf = fitz.open()

                    for page_num in paginas_marcadas:
                        novo_pdf.insert_pdf(pdf, from_page=page_num, to_page=page_num)

                    output_pdf_bytes = BytesIO()
                    novo_pdf.save(
                        output_pdf_bytes,
                        encryption=fitz.PDF_ENCRYPT_AES_256,
                        permissions=fitz.PDF_PERM_PRINT,
                        owner_pw="senha_segura"
                    )
                    novo_pdf.close()

                    zipf.writestr(output_pdf_name, output_pdf_bytes.getvalue())
                    selected_data.append({'MATRICULA': matricula, 'NOME': nome})

                pdf.close()

            for nome, matricula in zip(nomes, matriculas):
                marcar_e_salvar_pagina(nome, matricula)

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
    components.html("<script>window.location.reload();</script>", height=0)  # 🔄 Recarrega a página

# 🔹 Adiciona um script para limpar o estado ao sair da página
st.markdown(
    """<script>
        window.addEventListener('beforeunload', function () {
            fetch('/_stcore/clear_session_cache');
        });
    </script>""",
    unsafe_allow_html=True
)

# Interface Streamlit
st.title("Processamento de PDF e Excel para Arquivos Alelo")
st.write("""
    Faça upload de um arquivo Excel e um arquivo PDF. Atenção para esses detalhes!
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
