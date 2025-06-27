import streamlit as st
import tempfile
from resumo_util import processar_arquivo, resumir_noticias, exportar_resumos_para_word

st.set_page_config(page_title="Resumos de Notícias", layout="centered")

st.title("📰 Resumidor de Notícias (.docx)")

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

    st.download_button(
        label="📥 Baixar arquivo Word com resumos",
        data=open("resumos_final.docx", "rb").read(),
        file_name="resumos_noticias.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )