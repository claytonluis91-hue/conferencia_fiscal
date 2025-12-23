import pandas as pd
import glob
import os

def limpar_moeda_brasil(x):
    """Converte formato 1.000,00 para 1000.00"""
    if pd.isna(x) or str(x).strip() == '':
        return 0.0
    # Remove pontos de milhar e troca vírgula por ponto
    x_str = str(x).replace('.', '').replace(',', '.')
    try:
        return float(x_str)
    except:
        return 0.0

def limpar_moeda_padrao(x):
    """Converte formato padrão 1000.00"""
    if pd.isna(x) or str(x).strip() == '':
        return 0.0
    try:
        return float(x)
    except:
        return 0.0

def normalizar_doc(x):
    """Remove zeros à esquerda e espaços do número do documento"""
    if pd.isna(x):
        return ''
    return str(x).strip().lstrip('0')

def encontrar_arquivo(padrao):
    """Busca arquivos na pasta que comecem com o padrão informado"""
    arquivos = glob.glob(f"{padrao}*")
    # Filtra para ignorar o próprio arquivo de saída se ele começar com o padrão
    arquivos = [f for f in arquivos if "RELATORIO_FINAL" not in f]
    if arquivos:
        print(f"Arquivo encontrado para '{padrao}': {arquivos[0]}")
        return arquivos[0]
    else:
        print(f"ATENÇÃO: Nenhum arquivo encontrado começando com '{padrao}'")
        return None

def processar_conferencia():
    print("--- Iniciando Conferência PIS/COFINS ---\n")

    # --- 1. Carregar Arquivos de ESTOQUE ---
    colunas_finais = ['Doc', 'Data', 'CNPJ', 'PIS_Interno', 'COFINS_Interno']
    df_estoque_total = pd.DataFrame()

    # Ler ENTRADAS
    arq_ent = encontrar_arquivo("ESTOQUE_ENTRADAS")
    if arq_ent:
        # Colunas específicas do layout de Entrada
        # 0:Doc, 1:Data, 2:CNPJ, 27:Val_PIS, 30:Val_COFINS (ajustado conforme análise)
        try:
            df_ent = pd.read_csv(arq_ent, sep=';', header=None, dtype=str, 
                               usecols=[0, 1, 2, 27, 30])
            df_ent.columns = ['Doc', 'Data', 'CNPJ', 'PIS_Interno', 'COFINS_Interno']
            df_estoque_total = pd.concat([df_estoque_total, df_ent])
        except Exception as e:
            print(f"Erro ao ler entradas: {e}")

    # Ler SAÍDAS
    arq_sai = encontrar_arquivo("ESTOQUE_SAIDAS")
    if arq_sai:
        # Colunas específicas do layout de Saída (deslocamento de colunas observado)
        # 0:Doc, 1:Data, 2:CNPJ, 28:Val_PIS, 31:Val_COFINS
        try:
            df_sai = pd.read_csv(arq_sai, sep=';', header=None, dtype=str, 
                               usecols=[0, 1, 2, 28, 31])
            df_sai.columns = ['Doc', 'Data', 'CNPJ', 'PIS_Interno', 'COFINS_Interno']
            df_estoque_total = pd.concat([df_estoque_total, df_sai])
        except Exception as e:
            print(f"Erro ao ler saídas: {e}")

    if df_estoque_total.empty:
        print("Erro: Nenhum dado de estoque carregado. Verifique os arquivos.")
        return

    # Limpeza Estoque
    df_estoque_total['PIS_Interno'] = df_estoque_total['PIS_Interno'].apply(limpar_moeda_brasil)
    df_estoque_total['COFINS_Interno'] = df_estoque_total['COFINS_Interno'].apply(limpar_moeda_brasil)
    df_estoque_total['Doc_Norm'] = df_estoque_total['Doc'].apply(normalizar_doc)

    # Agrupar por documento (soma itens da mesma nota)
    df_estoque_agg = df_estoque_total.groupby('Doc_Norm').agg({
        'PIS_Interno': 'sum',
        'COFINS_Interno': 'sum',
        'Data': 'first',
        'CNPJ': 'first'
    }).reset_index()

    # --- 2. Carregar Arquivos do CLIENTE ---
    # Função auxiliar para ler relatório do cliente (procura cabeçalho)
    def ler_relatorio_cliente(arquivo, coluna_valor_nome):
        if not arquivo: return pd.DataFrame()
        
        # Lê primeiras linhas para achar onde começa o cabeçalho
        with open(arquivo, 'r', encoding='utf-8', errors='ignore') as f:
            linhas = [f.readline() for _ in range(50)]
        
        skip = 0
        for i, linha in enumerate(linhas):
            if "Documento" in linha and "Valor" in linha:
                skip = i
                break
        
        df = pd.read_csv(arquivo, skiprows=skip, dtype=str)
        # Limpa colunas vazias ou sem nome
        df = df.dropna(axis=1, how='all')
        
        # Garante que temos as colunas necessárias
        if 'Documento' not in df.columns:
            print(f"Aviso: Coluna 'Documento' não encontrada em {arquivo}")
            return pd.DataFrame()
            
        # Pega a coluna de Valor (geralmente a coluna 'Valor' ou similar)
        # No seu arquivo, a coluna se chamava 'Valor'
        col_valor = 'Valor' if 'Valor' in df.columns else df.columns[26] # Tentativa pelo indice se nome falhar
        
        df_clean = df[['Documento', col_valor]].copy()
        df_clean = df_clean.dropna(subset=['Documento'])
        df_clean['Valor_Formatado'] = df_clean[col_valor].apply(limpar_moeda_padrao)
        df_clean['Doc_Norm'] = df_clean['Documento'].apply(normalizar_doc)
        
        return df_clean.groupby('Doc_Norm')['Valor_Formatado'].sum().reset_index()

    # Processar PIS Cliente
    arq_pis = encontrar_arquivo("PIS")
    df_pis_cli = ler_relatorio_cliente(arq_pis, 'PIS')
    if not df_pis_cli.empty:
        df_pis_cli = df_pis_cli.rename(columns={'Valor_Formatado': 'PIS_Cliente'})
    else:
        df_pis_cli = pd.DataFrame(columns=['Doc_Norm', 'PIS_Cliente'])

    # Processar COFINS Cliente
    arq_cofins = encontrar_arquivo("COFINS")
    df_cofins_cli = ler_relatorio_cliente(arq_cofins, 'COFINS')
    if not df_cofins_cli.empty:
        df_cofins_cli = df_cofins_cli.rename(columns={'Valor_Formatado': 'COFINS_Cliente'})
    else:
        df_cofins_cli = pd.DataFrame(columns=['Doc_Norm', 'COFINS_Cliente'])

    # --- 3. Unificar Tudo ---
    print("\nCruzando informações...")
    
    # Junta PIS e COFINS do cliente
    df_cliente = pd.merge(df_pis_cli, df_cofins_cli, on='Doc_Norm', how='outer').fillna(0)
    
    # Junta com Estoque Interno
    df_final = pd.merge(df_estoque_agg, df_cliente, on='Doc_Norm', how='outer').fillna(0)

    # Calcular Diferenças
    df_final['Diferenca_PIS'] = df_final['PIS_Interno'] - df_final['PIS_Cliente']
    df_final['Diferenca_COFINS'] = df_final['COFINS_Interno'] - df_final['COFINS_Cliente']

    # Ordenar por maior diferença total (absoluta)
    df_final['Diff_Abs'] = df_final['Diferenca_PIS'].abs() + df_final['Diferenca_COFINS'].abs()
    df_final = df_final.sort_values('Diff_Abs', ascending=False).drop(columns=['Diff_Abs'])

    # Organizar colunas
    colunas_ordem = ['Doc_Norm', 'Data', 'CNPJ', 
                     'PIS_Interno', 'PIS_Cliente', 'Diferenca_PIS',
                     'COFINS_Interno', 'COFINS_Cliente', 'Diferenca_COFINS']
    
    # Preencher vazios em Data/CNPJ para notas que só existem no cliente
    df_final['Data'] = df_final['Data'].replace(0, '')
    df_final['CNPJ'] = df_final['CNPJ'].replace(0, '')

    # Salvar
    nome_saida = 'RELATORIO_FINAL_CONFERENCIA.xlsx'
    df_final[colunas_ordem].to_excel(nome_saida, index=False)
    
    print(f"\nSucesso! Arquivo gerado: {nome_saida}")
    print("Pressione Enter para sair.")
    input()

if __name__ == "__main__":
    processar_conferencia()