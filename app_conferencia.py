import streamlit as st
import pandas as pd
import io

# --- Configuração da Página ---
st.set_page_config(page_title="Conferência Fiscal", layout="wide")
st.title("📊 Conferência Fiscal: Estoque (Nascel) vs. DRT (Dominio)")

# --- Funções de Limpeza e Conversão ---
def limpar_moeda_brasil(x):
    """Converte 1.000,00 (BR) para float"""
    if pd.isna(x) or str(x).strip() == '': return 0.0
    x_str = str(x).replace('.', '').replace(',', '.')
    try:
        return float(x_str)
    except:
        return 0.0

def limpar_moeda_padrao(x):
    """Converte 1000.00 (US) para float"""
    if pd.isna(x) or str(x).strip() == '': return 0.0
    try:
        return float(x)
    except:
        return 0.0

def normalizar_doc(x):
    """Remove zeros à esquerda e espaços"""
    if pd.isna(x): return ''
    # Pega apenas números, remove .0 se houver
    s = str(x).split('.')[0].strip()
    return s.lstrip('0')

# --- 1. Carregar ESTOQUE (Nascel) ---
def carregar_estoque(arquivo_ent, arquivo_sai):
    df_total = pd.DataFrame()

    # Layout ENTRADAS
    # Col 0: Doc, Col 27: PIS, Col 30: COFINS
    if arquivo_ent:
        try:
            df_ent = pd.read_csv(arquivo_ent, sep=';', header=None, dtype=str,
                                 usecols=[0, 1, 2, 27, 30])
            df_ent.columns = ['Doc', 'Data', 'CNPJ', 'PIS_Interno', 'COFINS_Interno']
            df_total = pd.concat([df_total, df_ent])
        except Exception as e:
            st.error(f"Erro ao ler Entradas: {e}")

    # Layout SAIDAS
    # Col 0: Doc, Col 28: PIS, Col 31: COFINS
    if arquivo_sai:
        try:
            df_sai = pd.read_csv(arquivo_sai, sep=';', header=None, dtype=str,
                                 usecols=[0, 1, 2, 28, 31])
            df_sai.columns = ['Doc', 'Data', 'CNPJ', 'PIS_Interno', 'COFINS_Interno']
            df_total = pd.concat([df_total, df_sai])
        except Exception as e:
            st.error(f"Erro ao ler Saídas: {e}")

    if not df_total.empty:
        # Limpeza
        for col in ['PIS_Interno', 'COFINS_Interno']:
            df_total[col] = df_total[col].apply(limpar_moeda_brasil)
        
        df_total['Doc_Norm'] = df_total['Doc'].apply(normalizar_doc)
        
        # Agrupa por documento (soma itens da mesma nota)
        return df_total.groupby('Doc_Norm').agg({
            'PIS_Interno': 'sum', 'COFINS_Interno': 'sum',
            'Data': 'first', 'CNPJ': 'first'
        }).reset_index()
    
    return pd.DataFrame()

# --- 2. Carregar RELATÓRIO CLIENTE (DRT) ---
def carregar_cliente(arquivo, tipo_imposto):
    if not arquivo: return pd.DataFrame()

    try:
        # Lê o arquivo para achar onde começa o cabeçalho real
        string_io = io.StringIO(arquivo.getvalue().decode("utf-8", errors='ignore'))
        linhas = string_io.readlines()
        
        skip = 0
        header_found = False
        # Procura a linha que tem 'Documento' e 'Valor'
        for i, linha in enumerate(linhas[:60]):
            if "Documento" in linha and "Valor" in linha:
                skip = i
                header_found = True
                break
        
        if not header_found:
            st.warning(f"Cabeçalho não encontrado no arquivo {tipo_imposto}. Verifique se é o relatório detalhado correto.")
            return pd.DataFrame()

        # Lê o CSV pulando as linhas inúteis
        arquivo.seek(0)
        df = pd.read_csv(arquivo, skiprows=skip, dtype=str)
        
        # Limpa colunas vazias
        df = df.dropna(axis=1, how='all')

        # Verifica colunas necessárias
        if 'Documento' not in df.columns or 'Valor' not in df.columns:
            st.error(f"Colunas 'Documento' ou 'Valor' não encontradas em {tipo_imposto}.")
            return pd.DataFrame()

        # Filtra e Limpa
        df_clean = df[['Documento', 'Valor']].copy()
        df_clean = df_clean.dropna(subset=['Documento'])
        
        # Remove linhas de totalizadores (se houver texto na coluna valor)
        df_clean = df_clean[pd.to_numeric(df_clean['Valor'], errors='coerce').notnull()]

        nome_col = f"{tipo_imposto}_Cliente"
        df_clean[nome_col] = df_clean['Valor'].apply(limpar_moeda_padrao)
        df_clean['Doc_Norm'] = df_clean['Documento'].apply(normalizar_doc)

        return df_clean.groupby('Doc_Norm')[nome_col].sum().reset_index()

    except Exception as e:
        st.error(f"Erro ao processar {tipo_imposto}: {e}")
        return pd.DataFrame()

# --- INTERFACE ---
col1, col2 = st.columns(2)

with col1:
    st.info("📂 Arquivos NASCEL (Estoque)")
    f_ent = st.file_uploader("Upload ENTRADAS_ESTOQUE (.csv)", key="ent")
    f_sai = st.file_uploader("Upload SAIDAS_ESTOQUE (.csv)", key="sai")

with col2:
    st.warning("📂 Arquivos DRT (Relatórios)")
    f_pis = st.file_uploader("Upload Relatório PIS (.csv)", key="pis")
    f_cof = st.file_uploader("Upload Relatório COFINS (.csv)", key="cof")

st.markdown("---")

if st.button("🚀 PROCESSAR CONFERÊNCIA", type="primary"):
    if not (f_ent or f_sai):
        st.error("Por favor, anexe os arquivos de Estoque.")
    else:
        # 1. Processar Estoque
        df_est = carregar_estoque(f_ent, f_sai)
        
        # 2. Processar Cliente
        df_pis = carregar_cliente(f_pis, "PIS")
        df_cof = carregar_cliente(f_cof, "COFINS")

        # Garante DataFrames mínimos
        if df_pis.empty: df_pis = pd.DataFrame(columns=['Doc_Norm', 'PIS_Cliente'])
        if df_cof.empty: df_cof = pd.DataFrame(columns=['Doc_Norm', 'COFINS_Cliente'])

        # 3. Cruzamento
        # Junta PIS e COFINS do Cliente
        df_cli = pd.merge(df_pis, df_cof, on='Doc_Norm', how='outer').fillna(0)

        # Junta com Estoque
        if not df_est.empty:
            df_final = pd.merge(df_est, df_cli, on='Doc_Norm', how='outer').fillna(0)

            # Cálculos
            df_final['Diff_PIS'] = df_final['PIS_Interno'] - df_final['PIS_Cliente']
            df_final['Diff_COFINS'] = df_final['COFINS_Interno'] - df_final['COFINS_Cliente']
            
            # Ordenação por maior diferença
            df_final['Total_Diff'] = df_final['Diff_PIS'].abs() + df_final['Diff_COFINS'].abs()
            df_final = df_final.sort_values('Total_Diff', ascending=False).drop(columns=['Total_Diff'])

            # Seleção de Colunas
            cols = ['Doc_Norm', 'Data', 'CNPJ', 
                    'PIS_Interno', 'PIS_Cliente', 'Diff_PIS',
                    'COFINS_Interno', 'COFINS_Cliente', 'Diff_COFINS']
            
            # Formatação para Exibição
            st.success("Conferência realizada com sucesso!")
            st.write("Visualização das maiores diferenças:")
            st.dataframe(df_final[cols].head(50).style.format("{:.2f}"))

            # Download Excel
            try:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_final[cols].to_excel(writer, index=False)
                
                st.download_button("📥 Baixar Relatório Completo (Excel)", 
                                data=buffer.getvalue(), 
                                file_name="Conferencia_Nascel_DRT.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except:
                st.warning("Erro ao gerar Excel. Baixe em CSV.")
                csv = df_final[cols].to_csv(index=False, sep=';', decimal=',').encode('utf-8')
                st.download_button("📥 Baixar CSV", data=csv, file_name="Conferencia.csv")
        else:
            st.error("Não foi possível ler os dados de estoque. Verifique os arquivos CSV.")
