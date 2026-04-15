import streamlit as st
import pandas as pd
import base64
from datetime import datetime
import os
from typing import List
import tempfile

from extractor import NFeExtractor
from excel_generator import NFeExcelGenerator

# Configuração da página
st.set_page_config(
    page_title="Extrator de NF-e",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS customizado
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1F4E78;
        text-align: center;
        margin-bottom: 2rem;
    }
    .subheader {
        font-size: 1.2rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        background-color: #d4edda;
        color: #155724;
        padding: 1rem;
        border-radius: 0.5rem;
        border: 1px solid #c3e6cb;
    }
    .info-box {
        background-color: #d1ecf1;
        color: #0c5460;
        padding: 1rem;
        border-radius: 0.5rem;
        border: 1px solid #bee5eb;
        margin-bottom: 1rem;
    }
    .stProgress > div > div > div > div {
        background-color: #1F4E78;
    }
</style>
""", unsafe_allow_html=True)

def get_download_link(file_bytes: bytes, filename: str, text: str):
    """Gera link de download"""
    b64 = base64.b64encode(file_bytes).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}" style="text-decoration:none;"><button style="background-color:#1F4E78;color:white;padding:10px 24px;border:none;border-radius:5px;cursor:pointer;font-size:16px;">{text}</button></a>'
    return href

def processar_nfe(uploaded_file) -> dict:
    """Processa um arquivo PDF de NF-e"""
    try:
        # Salva temporariamente
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = tmp_file.name
        
        # Extrai dados
        extractor = NFeExtractor(tmp_path)
        dados = extractor.processar()
        
        # Remove arquivo temporário
        os.unlink(tmp_path)
        
        return dados
    except Exception as e:
        st.error(f"Erro ao processar {uploaded_file.name}: {str(e)}")
        return None

def main():
    st.markdown('<div class="main-header">📄 Sistema de Extração de NF-e</div>', unsafe_allow_html=True)
    st.markdown('<div class="subheader">Extraia automaticamente todos os dados de Notas Fiscais Eletrônicas para Excel</div>', unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.image("https://via.placeholder.com/150x150.png?text=NF-e+Extractor", use_column_width=True)
        st.markdown("---")
        st.markdown("### 📋 Instruções")
        st.markdown("""
        1. Faça upload dos arquivos PDF das NF-es (DANFE)
        2. O sistema extrairá automaticamente:
           - Dados do emitente e destinatário
           - Todos os produtos e serviços
           - Valores e impostos
           - Informações complementares
        3. Download do Excel formatado
        """)
        st.markdown("---")
        st.markdown("### ⚙️ Configurações")
        modo_extracao = st.selectbox(
            "Modo de extração:",
            ["Padrão", "OCR (para PDFs escaneados)"]
        )
        incluir_aba_resumo = st.checkbox("Incluir aba de resumo", value=True)
        incluir_detalhamento = st.checkbox("Incluir abas detalhadas", value=True)
    
    # Área principal
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown('<div class="info-box">📎 <strong>Selecione os arquivos PDF das NF-es</strong> (pode selecionar múltiplos arquivos)</div>', unsafe_allow_html=True)
        
        uploaded_files = st.file_uploader(
            "Arraste os arquivos ou clique para selecionar",
            type=['pdf'],
            accept_multiple_files=True,
            help="Selecione um ou mais arquivos PDF de Notas Fiscais Eletrônicas (DANFE)"
        )
    
    with col2:
        st.metric("Arquivos Selecionados", len(uploaded_files) if uploaded_files else 0)
        if uploaded_files:
            st.info(f"Tamanho total: {sum(f.size for f in uploaded_files)/1024/1024:.2f} MB")
    
    if uploaded_files:
        st.markdown("---")
        
        if st.button("🚀 PROCESSAR NOTAS FISCAIS", type="primary", use_container_width=True):
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            dados_processados = []
            erros = []
            
            for idx, uploaded_file in enumerate(uploaded_files):
                progress = (idx + 1) / len(uploaded_files)
                progress_bar.progress(progress)
                status_text.text(f"Processando {idx + 1} de {len(uploaded_files)}: {uploaded_file.name}...")
                
                dados = processar_nfe(uploaded_file)
                if dados:
                    dados_processados.append(dados)
                else:
                    erros.append(uploaded_file.name)
            
            progress_bar.empty()
            status_text.empty()
            
            if dados_processados:
                st.success(f"✅ {len(dados_processados)} nota(s) processada(s) com sucesso!")
                
                if erros:
                    st.warning(f"⚠️ {len(erros)} arquivo(s) não puderam ser processados: {', '.join(erros)}")
                
                # Preview dos dados
                with st.expander("👁️ Visualizar Dados Extraídos", expanded=True):
                    tabs = st.tabs(["Resumo", "Detalhes"])
                    
                    with tabs[0]:
                        # DataFrame de resumo
                        df_resumo = pd.DataFrame([
                            {
                                'Arquivo': d['nome_arquivo'],
                                'NF Nº': d['cabecalho']['numero'],
                                'Série': d['cabecalho']['serie'],
                                'Data': d['cabecalho']['data_emissao'],
                                'Emitente': d['emitente']['nome'],
                                'Destinatário': d['destinatario']['nome'],
                                'Valor Total': d['totais']['valor_total'],
                                'Itens': d['quantidade_produtos']
                            }
                            for d in dados_processados
                        ])
                        st.dataframe(df_resumo, use_container_width=True)
                    
                    with tabs[1]:
                        # Detalhes do primeiro arquivo como exemplo
                        if dados_processados:
                            d = dados_processados[0]
                            col_a, col_b = st.columns(2)
                            with col_a:
                                st.markdown("**Emitente:**")
                                st.json(d['emitente'])
                            with col_b:
                                st.markdown("**Destinatário:**")
                                st.json(d['destinatario'])
                            
                            st.markdown("**Produtos:**")
                            df_prod = pd.DataFrame(d['produtos'])
                            st.dataframe(df_prod, use_container_width=True)
                
                # Geração do Excel
                st.markdown("---")
                st.markdown("### 📥 Download do Excel")
                
                try:
                    generator = NFeExcelGenerator()
                    excel_bytes = generator.gerar_excel(dados_processados)
                    
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"NFes_Extraidas_{timestamp}.xlsx"
                    
                    st.markdown(
                        get_download_link(excel_bytes, filename, "📥 BAIXAR EXCEL COMPLETO"),
                        unsafe_allow_html=True
                    )
                    
                    st.markdown("""
                    <div style="margin-top:1rem; padding:1rem; background-color:#f8f9fa; border-radius:0.5rem; border-left:4px solid #1F4E78;">
                        <strong>O arquivo Excel contém:</strong><br>
                        • Aba "Resumo NF-es" com todas as notas em uma visão consolidada<br>
                        • Abas individuais para cada NF-e com detalhamento completo<br>
                        • Formatação profissional, filtros e colunas ajustadas<br>
                        • Fórmulas de totais preservadas
                    </div>
                    """, unsafe_allow_html=True)
                    
                except Exception as e:
                    st.error(f"Erro ao gerar Excel: {str(e)}")
            else:
                st.error("❌ Nenhuma nota fiscal pôde ser processada. Verifique se os PDFs são válidos (DANFE).")

if __name__ == "__main__":
    main()
