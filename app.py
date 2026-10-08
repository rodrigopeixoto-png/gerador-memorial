import streamlit as st
import tempfile
import os
from processor import gerar_memorial_docx

# Configuração da página
st.set_page_config(page_title="Gerador de Memorial BIM", page_icon="🏗️", layout="centered")

st.title("Gerador Automático de Memorial 🏗️")
st.markdown("Faça o upload do seu arquivo **IFC estrutural** para gerar o documento Word.")

# Componente de Upload
arquivo_ifc = st.file_uploader("Arraste ou selecione o arquivo .ifc", type=['ifc'])

if arquivo_ifc is not None:
    # Botão de gatilho
    if st.button("Gerar Memorial Descritivo", type="primary"):
        with st.spinner("Lendo metadados do IFC e montando o documento..."):
            try:
                # 1. Salvar o arquivo upado temporariamente no servidor
                with tempfile.NamedTemporaryFile(delete=False, suffix=".ifc") as tmp_file:
                    tmp_file.write(arquivo_ifc.getvalue())
                    caminho_temporario = tmp_file.name
                
                # 2. Chamar o nosso motor de processamento
                docx_buffer = gerar_memorial_docx(caminho_temporario)
                
                st.success("Memorial gerado com sucesso!")
                
                # 3. Disponibilizar o arquivo para download
                st.download_button(
                    label="📥 Baixar Documento (.docx)",
                    data=docx_buffer,
                    file_name="Memorial_Descritivo.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            except Exception as e:
                st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
                
            finally:
                # 4. Limpeza: apagar o arquivo temporário do servidor
                if os.path.exists(caminho_temporario):
                    os.remove(caminho_temporario)
