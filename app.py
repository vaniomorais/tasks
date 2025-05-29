import streamlit as st
import os
from extracao import baixar_pdfs_outlook, extrair_informacoes_pdf, comparar_pdfs, gerar_oficio
from dotenv import load_dotenv

load_dotenv()  # Carrega variáveis do .env

st.title("📜 Verificação de Documentos Jurídicos")

remetentes = os.environ.get("REMETENTES", "").split(",")
remetente = st.selectbox("Selecione o remetente do e-mail:", remetentes)
conta_email = os.environ.get("CONTA_EMAIL", "")
pasta_pdfs = os.path.join(os.path.dirname(__file__), "pdfs")

def exibir_resultados(resultados):
    st.subheader("🔎 Resultados da comparação")
    for pdf_email, pdf_pasta, autenticidade in resultados:
        st.write(f"**{pdf_email}** vs **{pdf_pasta}** → **{autenticidade}**")

if st.button("Baixar e verificar PDFs"):
    st.write("📩 Buscando e-mails...")
    pdfs_recebidos = baixar_pdfs_outlook(remetente, conta_email)

    if not pdfs_recebidos:
        st.error("Nenhum PDF encontrado!")
    else:
        st.success(f"{len(pdfs_recebidos)} PDFs baixados com sucesso!")
        juiz, processo = extrair_informacoes_pdf(pdfs_recebidos[0])
        st.write(f"**Juiz:** {juiz}")
        st.write(f"**Processo:** {processo}")

        resultados = comparar_pdfs(pdfs_recebidos, pasta_pdfs)
        exibir_resultados(resultados)

        if resultados:
            caminho_oficio = gerar_oficio(resultados[0][2], pdfs_recebidos[0], juiz, processo)
            st.success("📄 Ofício gerado com sucesso!")
            with open(caminho_oficio, "rb") as f:
                st.download_button("📥 Baixar Ofício", f, "Oficio_Resposta.docx")