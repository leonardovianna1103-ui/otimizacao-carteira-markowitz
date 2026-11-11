import os
import glob
import re
import pandas as pd

print("=" * 70)
print("CONSOLIDAÇÃO AUTOMÁTICA DOS DADOS MSCI (EUROPA DESENVOLVIDA)")
print("=" * 70)

# 1. Definição dos diretórios (data/raw e data/processed)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
pasta_raw = os.path.join(BASE_DIR, "data", "raw")
pasta_processed = os.path.join(BASE_DIR, "data", "processed")

os.makedirs(pasta_raw, exist_ok=True)
os.makedirs(pasta_processed, exist_ok=True)

# Busca todos os arquivos .xlsx e .xls que contêm MSCI na pasta data/raw
padrao_busca = os.path.join(pasta_raw, "*MSCI*.xls*")
lista_arquivos = glob.glob(padrao_busca)

# Se não encontrar pelo filtro específico, busca todos os .xlsx da pasta raw
if not lista_arquivos:
    lista_arquivos = [
        f for f in glob.glob(os.path.join(pasta_raw, "*.xlsx"))
        if "consolidado" not in f.lower() and "msci_europe_11" not in f.lower()
    ]

if not lista_arquivos:
    raise FileNotFoundError(
        f"Nenhum arquivo Excel encontrado na pasta:\n{pasta_raw}\n"
        "Certifique-se de salvar as planilhas baixadas da MSCI dentro de 'data/raw/'."
    )

print(f"-> Arquivos localizados em 'data/raw/': {len(lista_arquivos)}")

# 2. Leitura e extração de cada país
dfs_paises = []

for caminho in lista_arquivos:
    nome_arquivo = os.path.basename(caminho)
    try:
        # Lê a primeira aba sem assumir cabeçalhos fixos
        df_bruto = pd.read_excel(caminho, sheet_name=0, header=None)
        
        # Localiza a linha onde está a palavra 'Date' na coluna A
        linha_cabecalho = df_bruto[df_bruto[0].astype(str).str.strip().str.lower() == 'date'].index
        if len(linha_cabecalho) == 0:
            print(f"  [Aviso] Ignorando '{nome_arquivo}': cabeçalho 'Date' não identificado.")
            continue
        idx_header = linha_cabecalho[0]
        
        # Extrai o nome do país a partir da célula da coluna B (ex: 'MSCI Switzerland Index' -> 'Switzerland')
        texto_indice = str(df_bruto.iloc[idx_header, 1]).strip()
        match_pais = re.search(r'MSCI\s+(.*?)\s+Index', texto_indice, re.IGNORECASE)
        if match_pais:
            nome_pais = match_pais.group(1).strip()
        else:
            nome_pais = texto_indice.replace('MSCI', '').replace('Index', '').strip()
            
        # Extrai apenas as colunas A (Data) e B (Cotação) a partir da linha de dados
        df_dados = df_bruto.iloc[idx_header + 1:, :2].copy()
        df_dados.columns = ['Date', nome_pais]
        
        # Tratamento de tipagem
        df_dados = df_dados.dropna(subset=['Date', nome_pais])
        df_dados['Date'] = pd.to_datetime(df_dados['Date'], errors='coerce')
        df_dados[nome_pais] = pd.to_numeric(df_dados[nome_pais], errors='coerce')
        df_dados = df_dados.dropna()
        
        # Cria uma chave canônica Ano-Mês para alinhamento perfeito entre países
        df_dados['Ano_Mes'] = df_dados['Date'].dt.to_period('M')
        
        # Mantém a última cotação do mês (elimina duplicatas eventuais)
        df_dados = df_dados.sort_values('Date').drop_duplicates(subset=['Ano_Mes'], keep='last')
        
        dfs_paises.append(df_dados[['Ano_Mes', nome_pais]])
        print(f"  ✓ {nome_pais:<15} ({len(df_dados)} meses processados) <- {nome_arquivo}")
        
    except Exception as e:
        print(f"  [Erro] Falha ao processar {nome_arquivo}: {e}")

if not dfs_paises:
    raise ValueError("Nenhum dado válido de país pôde ser extraído.")

# 3. Consolidação via Merge Externo pela chave Ano-Mês
df_consolidado = dfs_paises[0]
for df_prox in dfs_paises[1:]:
    df_consolidado = pd.merge(df_consolidado, df_prox, on='Ano_Mes', how='outer')

# Ordena cronologicamente
df_consolidado = df_consolidado.sort_values('Ano_Mes').reset_index(drop=True)

# Converte o período 'Ano_Mes' para a data padrão do último dia de cada mês
df_consolidado['Date'] = df_consolidado['Ano_Mes'].dt.to_timestamp(how='end').dt.strftime('%Y-%m-%d')
df_consolidado = df_consolidado.drop(columns=['Ano_Mes'])

# Posiciona a coluna 'Date' em primeiro lugar
colunas = ['Date'] + [col for col in df_consolidado.columns if col != 'Date']
df_consolidado = df_consolidado[colunas]

# 4. Tratamento de período comum (sem dados faltantes para a otimização)
df_limpo = df_consolidado.dropna()

# 5. Exportação dos Arquivos para data/processed/
caminho_csv = os.path.join(pasta_processed, 'msci_europe_11_countries_usd.csv')
caminho_xlsx = os.path.join(pasta_processed, 'msci_europe_11_countries_usd.xlsx')

df_limpo.to_csv(caminho_csv, index=False)
df_limpo.to_excel(caminho_xlsx, index=False)

print("=" * 70)
print("✓ Base consolidada com sucesso!")
print(f"-> Arquivo CSV gerado:   {caminho_csv}")
print(f"-> Arquivo Excel gerado: {caminho_xlsx}")
print(f"-> Dimensões da base: {df_limpo.shape[0]} meses e {len(df_limpo.columns)-1} países.")
print(f"-> Janela temporal completa: de {df_limpo['Date'].iloc[0]} até {df_limpo['Date'].iloc[-1]}")
print("=" * 70)