import streamlit as st
import pandas as pd
import io

# --- Configuração da Página ---
st.set_page_config(page_title="Conferência Fiscal PIS/COFINS", layout="wide")

st.title("📊 Conferência de Impostos: Estoque vs. Cliente")
st.markdown("""
Esta ferramenta cruza os dados dos seus relatórios de Estoque (Entrada/Saída) 
com os relatórios de PIS/COFINS fornecidos pelo cliente/contabilidade.
""")

# --- Funções de Limpeza ---
def limpar_moeda_brasil(x):
    """Converte 1.000,00 para float"""
    if pd.isna(x) or str(x).strip() == '':
        return 0.0
    x_str = str(x).replace('.', '').replace(',', '.')
    try:
        return float(x_str)
    except:
        return 0.0

def limpar_moeda_padrao(x):
    """Converte 1000.00 para float"""
    if pd.isna(x) or str(x).strip() == '':
        return 0.0
    try:
        return float(x)
    except:
        return 0.0

def normalizar_doc(x):
    """Remove zeros à esquerda e espaços"""
    if pd.isna(x):
        return ''
    return str(x).strip().lstrip('0')

# --- Função para processar Estoque ---
def carregar_estoque(arquivo_ent, arquivo_sai):
    df_total = pd.DataFrame()

    # Processar Entradas
    if arquivo_ent:
        try:
            # Layout Entradas: Doc(0), Data(1), CNPJ(2), PIS(27), COFINS(30)
            df_ent = pd.read_csv(arquivo_ent, sep=';', header=None, dtype=str,
                                 usecols=[0, 1, 2, 27, 30])
            df_ent.columns = ['Doc', 'Data', 'CNPJ', 'PIS_Interno', 'COFINS_Interno']
            df_total = pd.concat([df_total, df_ent])
        except Exception as e:
            st.error(f"Erro ao ler Entradas: {e}")

    # Processar Saídas
    if arquivo_sai:
        try:
            # Layout Saídas: Doc(0), Data(1), CNPJ(2), PIS(28), COFINS(31)
            df_sai = pd.read_csv(arquivo_sai, sep=';', header=None, dtype=str,
                                 usecols=[0, 1, 2, 28, 31])
            df_sai.columns = ['Doc', 'Data', 'CNPJ', 'PIS_Interno', 'COFINS_Interno']
            df_total = pd.concat([df_total, df_sai])
        except Exception as e:
            st.error(f"Erro ao ler Saídas: {e}")

    if not df_total.empty:
        df_total['PIS_Interno'] = df_total['PIS_Interno'].apply(limpar_moeda_brasil)
        df_total['COFINS_Interno'] = df_total['COFINS_Interno'].apply(limpar_moeda_brasil)
        df_total['Doc_Norm'] = df_total['Doc'].apply(normalizar_doc)
        
        # Agrupar
        return df_total.groupby('Doc_Norm').agg({
            'PIS_Interno': 'sum',
            'COFINS_Interno': 'sum',
            'Data': 'first',
            'CNPJ': 'first'
        }).reset_index()
    return pd.DataFrame()

# --- Função para processar Cliente ---
def carregar_cliente(arquivo, tipo_imposto):
    if not arquivo:
        return pd.DataFrame()

    try:
        # Lê o arquivo como string para achar o cabeçalho dinamicamente
        string_io = io.StringIO(arquivo.getvalue().decode("utf-8", errors='ignore'))
        linhas = string_io.readlines()
        
        skip = 0
        header_found = False
        for i, linha in enumerate(linhas[:60]): # Procura nas primeiras 60 linhas
            if "Documento" in linha and "Valor" in linha:
                skip = i
                header_found = True
                break
        
        if not header_found:
            st.warning(f"Cabeçalho não encontrado no arquivo {tipo_imposto}. Verifique o layout.")
            return pd.DataFrame()

        # Volta o ponteiro para ler com pandas
        arquivo.seek(0)
        df = pd.read_csv(arquivo, skiprows=skip, dtype=str)
        df = df.dropna(axis=1, how='all')

        # Identificar coluna de Valor
        col_valor = None
        if 'Valor' in df.columns:
            col_valor = 'Valor'
        else:
            # Tenta pegar pelo índice se o nome não for exato (segurança)
            if len(df.columns) > 26: 
                col_valor = df.columns[26]
        
        if 'Documento' not in df.columns or not col_valor:
            return pd.DataFrame()

        df_clean = df[['Documento', col_valor]].copy()
        df_clean = df_clean.dropna(subset=['Documento'])
        
        nome_col_final = f"{tipo_imposto}_Cliente"
        df_clean[nome_col_final] = df_clean[col_valor].apply(limpar_moeda_padrao)
        df_clean['Doc_Norm'] = df_clean['Documento'].apply(normalizar_doc)

        return df_clean.groupby('Doc_Norm')[nome_col_final].sum().reset_index()

    except Exception as e:
        st.error(f"Erro ao processar arquivo {tipo_imposto}: {e}")
        return pd.DataFrame()

# --- Interface Gráfica ---

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Seus Arquivos (Estoque)")
    file_ent = st.file_uploader("Upload Estoque ENTRADAS (.csv)", type=["csv"])
    file_sai = st.file_uploader("Upload Estoque SAÍDAS (.csv)", type=["csv"])

with col2:
    st.subheader("2. Arquivos do Cliente (PIS/COFINS)")
    file_pis = st.file_uploader("Upload Relatório PIS (.csv/txt)", type=["csv", "txt", "xlsx"])
    file_cof = st.file_uploader("Upload Relatório COFINS (.csv/txt)", type=["csv", "txt", "xlsx"])

st.markdown("---")

if st.button("🚀 Processar Conferência", type="primary"):
    if not (file_ent or file_sai):
        st.warning("Por favor, faça upload de pelo menos um arquivo de ESTOQUE.")
    else:
        with st.spinner('Processando dados...'):
            # 1. Carregar Estoque
            df_estoque = carregar_estoque(file_ent, file_sai)
            
            # 2. Carregar Cliente
            df_pis = carregar_cliente(file_pis, "PIS")
            df_cof = carregar_cliente(file_cof, "COFINS")

            if df_pis.empty:
                df_pis = pd.DataFrame(columns=['Doc_Norm', 'PIS_Cliente'])
            if df_cof.empty:
                df_cof = pd.DataFrame(columns=['Doc_Norm', 'COFINS_Cliente'])

            # 3. Cruzamento
            # Junta PIS e COFINS Cliente
            df_cli_total = pd.merge(df_pis, df_cof, on='Doc_Norm', how='outer').fillna(0)
            
            # Junta com Estoque
            if not df_estoque.empty:
                df_final = pd.merge(df_estoque, df_cli_total, on='Doc_Norm', how='outer').fillna(0)

                # Cálculos
                df_final['Diferenca_PIS'] = df_final['PIS_Interno'] - df_final['PIS_Cliente']
                df_final['Diferenca_COFINS'] = df_final['COFINS_Interno'] - df_final['COFINS_Cliente']

                # Ordenação e Layout
                df_final['Abs_Diff'] = df_final['Diferenca_PIS'].abs() + df_final['Diferenca_COFINS'].abs()
                df_final = df_final.sort_values('Abs_Diff', ascending=False).drop(columns=['Abs_Diff'])

                cols_order = ['Doc_Norm', 'Data', 'CNPJ', 
                              'PIS_Interno', 'PIS_Cliente', 'Diferenca_PIS',
                              'COFINS_Interno', 'COFINS_Cliente', 'Diferenca_COFINS']
                
                df_display = df_final[cols_order].copy()

                # Mostrar Resultados
                st.success("Conferência Finalizada!")
                
                st.subheader("Visualização das Maiores Divergências")
                st.dataframe(df_display.style.format("{:.2f}", subset=['PIS_Interno', 'PIS_Cliente', 'Diferenca_PIS', 
                                                                     'COFINS_Interno', 'COFINS_Cliente', 'Diferenca_COFINS']))

                # Botão de Download
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_display.to_excel(writer, index=False, sheet_name='Conferencia')
                
                st.download_button(
                    label="📥 Baixar Relatório em Excel",
                    data=buffer.getvalue(),
                    file_name="Resultado_Conferencia_PIS_COFINS.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            else:
                st.error("Não foi possível gerar dados de estoque. Verifique os arquivos de entrada.")