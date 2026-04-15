import streamlit as st
import pdfplumber
import pandas as pd
import re
from datetime import datetime
from io import BytesIO

# ====================== FUNÇÃO DE EXTRAÇÃO MELHORADA ======================
def limpar_valor(valor):
    if not valor:
        return ""
    valor = re.sub(r'[^\d.,-]', '', str(valor))
    valor = valor.replace('.', '').replace(',', '.')
    try:
        return float(valor)
    except:
        return valor

def extrair_dados_pdf(file_obj):
    with pdfplumber.open(file_obj) as pdf:
        texto_completo = ""
        tabelas = []
        for pagina in pdf.pages:
            texto_completo += (pagina.extract_text() or "") + "\n"
            tabelas.extend(pagina.extract_tables() or [])
    
    texto = texto_completo.upper()

    dados = {
        "CPF/CNPJ": "", "RAZÃO SOCIAL": "", "UF": "", "MUNICÍPIO E ENDEREÇO": "",
        "NÚMERO DE DOCUMENTO": "", "DATA DE EMISSÃO": "", "DATA DE ENTRADA": "",
        "SITUAÇÃO": "", "ACUMULADOR": "", "CFOP": "", "VALOR DE SERVIÇOS": "",
        "VALOR DESCONTO": "", "VALOR CONTÁBIL": "", "BASE DE CÁLCULO": "",
        "ALÍQUOTA ISS": "", "VALOR ISS NORMAL": "", "VALOR ISS RETIDO": "",
        "VALOR IRRF": "", "VALOR PIS": "", "VALOR COFINS": "", "VALOR CSLL": "",
        "VALOR CRF": "", "VALOR INSS": ""
    }

    # ================== PADRÕES MELHORADOS (baseados nos seus PDFs reais) ==================
    padroes = {
        "CPF/CNPJ": [
            r'60\.219\.250/0002-20',                    # CNPJ da Promaflex (Tomador)
            r'CNPJ[:\s]*([\d./-]+)',
            r'CPF/CNPJ[:\s]*([\d./-]+)'
        ],
        "RAZÃO SOCIAL": [
            r'PROMAFLEX INDUSTRIAL LTDA',
            r'Nome/Razão Social[:\s]*(.+?)(?=\s*(CNPJ|Endereço|Município|$))',
            r'RAZÃO SOCIAL[:\s]*(.+?)(?=\s*(CNPJ|UF|MUNICÍPIO|$))'
        ],
        "UF": [
            r'UF[:\s]*([A-Z]{2})',
            r'(SP|RS|SC|RJ)'
        ],
        "MUNICÍPIO E ENDEREÇO": [
            r'(TABOÃO DA SERRA|SÃO PAULO|OSASCO|JOINVILLE|PORTO ALEGRE)[\s\S]*?CEP[:\s]*([\d-]+)',
            r'Endereço[:\s]*(.+?)(?=\s*(Município|CEP|$))'
        ],
        "NÚMERO DE DOCUMENTO": [
            r'Número da Nota[:\s]*(\d+)',
            r'NFS-e[:\s]*(\d+)',
            r'Nota[:\s]*(\d+)',
            r'Número NFS-e[:\s]*(\d+)',
            r'Número / Série[:\s]*(\d+)'
        ],
        "DATA DE EMISSÃO": [
            r'Data e Hora de Emissão[:\s]*(\d{2}/\d{2}/\d{4})',
            r'Data de Emissão[:\s]*(\d{2}/\d{2}/\d{4})',
            r'Emissão[:\s]*(\d{2}/\d{2}/\d{4})'
        ],
        "VALOR DE SERVIÇOS": [
            r'VALOR TOTAL DA NOTA = R?\$?\s*([\d.,]+)',
            r'VALOR DO SERVIÇO = R?\$?\s*([\d.,]+)',
            r'VALOR TOTAL DA NFS-e = R?\$?\s*([\d.,]+)',
            r'VALOR TOTAL DO SERVIÇO[:\s]*R?\$?\s*([\d.,]+)',
            r'VALOR TOTAL[:\s]*R?\$?\s*([\d.,]+)'
        ],
        "ALÍQUOTA ISS": [
            r'Alíquota ISSQN \(%\)[:\s]*([\d.,]+)',
            r'Alíquota ISS[:\s]*([\d.,]+)',
            r'ALÍQUOTA \(\%\):[:\s]*([\d.,]+)'
        ],
        "VALOR ISS NORMAL": [
            r'Valor do ISSQN \(R\$\)[:\s]*([\d.,]+)',
            r'Valor do ISS \(R\$\)[:\s]*([\d.,]+)'
        ],
        "VALOR ISS RETIDO": [
            r'ISS Retido[:\s]*R?\$?\s*([\d.,]+)',
            r'Valor ISS Retido[:\s]*R?\$?\s*([\d.,]+)'
        ]
    }

    for campo, lista_regex in padroes.items():
        for regex in lista_regex:
            match = re.search(regex, texto_completo, re.IGNORECASE | re.DOTALL)
            if match:
                valor = match.group(1).strip() if len(match.groups()) > 0 else match.group(0).strip()
                dados[campo] = valor
                break

    # ================== EXTRAÇÃO DE ITENS (melhorada) ==================
    itens = []
    for tabela in tabelas:
        for linha in tabela:
            if linha and len(linha) >= 2:
                cod = str(linha[0]).strip() if linha[0] else ""
                if cod and re.search(r'\d', cod):
                    itens.append({
                        "CÓDIGO DO ITEM": cod,
                        "QUANTIDADE": limpar_valor(linha[1] if len(linha) > 1 else ""),
                        "VALOR UNITÁRIO": limpar_valor(linha[2] if len(linha) > 2 else "")
                    })
    if not itens:
        # Fallback por texto
        linhas_item = re.findall(r'(\d{1,10})\s+(\d+[,.]?\d*)\s+R?\$?\s*([\d.,]+)', texto_completo)
        for cod, qtd, vunit in linhas_item:
            itens.append({
                "CÓDIGO DO ITEM": cod,
                "QUANTIDADE": limpar_valor(qtd),
                "VALOR UNITÁRIO": limpar_valor(vunit)
            })

    return dados, itens

# ====================== INTERFACE STREAMLIT ======================
st.set_page_config(page_title="NF-PDF → Excel Automático", layout="wide", page_icon="🚀")

st.title("🚀 NF-PDF → Excel Automático")
st.markdown("**Extração automática de NFS-e com suporte a múltiplos layouts (Taboão da Serra, SP, Joinville, etc.)**")

uploaded_files = st.file_uploader(
    "📤 Selecione um ou vários arquivos PDF de Notas Fiscais",
    type="pdf",
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"✅ {len(uploaded_files)} arquivo(s) PDF carregado(s)")

    if st.button("🔄 Processar Notas Fiscais", type="primary", use_container_width=True):
        with st.spinner("Processando PDFs..."):
            progress_bar = st.progress(0)
            todas_linhas = []
            
            for idx, uploaded_file in enumerate(uploaded_files):
                try:
                    file_bytes = BytesIO(uploaded_file.getvalue())
                    cabecalho, lista_itens = extrair_dados_pdf(file_bytes)
                    
                    if not lista_itens:
                        lista_itens = [{}]
                    
                    for item in lista_itens:
                        linha = {**cabecalho, **item}
                        linha["Nome do arquivo PDF"] = uploaded_file.name
                        linha["Data/Hora do processamento"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                        todas_linhas.append(linha)
                
                except Exception as e:
                    st.error(f"❌ Erro ao processar {uploaded_file.name}: {str(e)}")
                    continue
                
                progress_bar.progress((idx + 1) / len(uploaded_files))
            
            df = pd.DataFrame(todas_linhas)
            
            ordem_colunas = [
                "CPF/CNPJ", "RAZÃO SOCIAL", "UF", "MUNICÍPIO E ENDEREÇO",
                "NÚMERO DE DOCUMENTO", "DATA DE EMISSÃO", "DATA DE ENTRADA", "SITUAÇÃO",
                "ACUMULADOR", "CFOP", "VALOR DE SERVIÇOS", "VALOR DESCONTO",
                "VALOR CONTÁBIL", "BASE DE CÁLCULO", "ALÍQUOTA ISS",
                "VALOR ISS NORMAL", "VALOR ISS RETIDO", "VALOR IRRF",
                "VALOR PIS", "VALOR COFINS", "VALOR CSLL", "VALOR CRF", "VALOR INSS",
                "CÓDIGO DO ITEM", "QUANTIDADE", "VALOR UNITÁRIO",
                "Nome do arquivo PDF", "Data/Hora do processamento"
            ]
            
            for col in ordem_colunas:
                if col not in df.columns:
                    df[col] = ""
            df = df[ordem_colunas]
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name="Notas_Fiscais")
                # Auto ajuste de colunas
                workbook = writer.book
                worksheet = writer.sheets["Notas_Fiscais"]
                for col in worksheet.columns:
                    max_length = 0
                    column_letter = col[0].column_letter
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    worksheet.column_dimensions[column_letter].width = min(max_length + 2, 50)
            
            output.seek(0)
            
            st.success(f"✅ **Processamento concluído!**  \n📄 {len(uploaded_files)} PDFs  \n📊 {len(df)} linhas geradas")
            st.subheader("🔍 Prévia dos dados")
            st.dataframe(df.head(10), use_container_width=True)
            
            st.download_button(
                label="⬇️ Baixar Excel Completo",
                data=output,
                file_name=f"Notas_Fiscais_Extraidas_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

st.markdown("---")
st.caption("Versão corrigida • Suporte a múltiplos layouts de NFS-e • 2026")
