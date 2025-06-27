import streamlit as st
import tempfile
import os
import zipfile
from fpdf import FPDF
from resumo_util import (
    processar_arquivo,
    resumir_noticias,
    exportar_resumos_para_word,
    extrair_valor_economico,
    salvar_noticias_valor_pdf,
    compactar_em_zip
)

st.set_page_config(page_title="Resumos de Notícias", layout="centered")
st.title("📰 Resumidor de Notícias (.docx)")

# Upload do arquivo
arquivo_doc = st.file_uploader("📎 Envie um arquivo .docx com notícias", type=["docx"])

if arquivo_doc:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(arquivo_doc.read())
        caminho_temp = tmp.name

    with st.spinner("⏳ Processando documento..."):
        noticias = processar_arquivo(caminho_temp)
        resumos = resumir_noticias(noticias)
        exportar_resumos_para_word(noticias, resumos, "resumos_final.docx")

    st.success("✅ Resumos prontos!")

    # Download do DOCX
    st.download_button(
        label="📥 Baixar arquivo Word com resumos",
        data=open("resumos_final.docx", "rb").read(),
        file_name="resumos_noticias.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    # PDFs do Valor Econômico
    st.subheader("📄 PDFs das notícias do Valor Econômico")

    noticias_valor = extrair_valor_economico(noticias)

    if not noticias_valor:
        st.info("Nenhuma notícia do Valor Econômico encontrada.")
    else:
        pdfs = salvar_noticias_valor_pdf(noticias_valor)

        if len(pdfs) == 1:
            with open(pdfs[0], "rb") as f:
                st.download_button(
                    label="📄 Baixar PDF (Valor Econômico)",
                    data=f,
                    file_name=os.path.basename(pdfs[0]),
                    mime="application/pdf"
                )
        else:
            zip_path = "noticias_valor.zip"
            compactar_em_zip(pdfs, zip_path)
            with open(zip_path, "rb") as f:
                st.download_button(
                    label="📦 Baixar ZIP com PDFs (Valor Econômico)",
                    data=f,
                    file_name="noticias_valor.zip",
                    mime="application/zip"
                )
