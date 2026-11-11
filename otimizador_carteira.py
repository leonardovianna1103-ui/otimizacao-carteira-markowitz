import os
import glob
import numpy as np
import pandas as pd
import scipy.optimize as sco
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =============================================================================
# 1. CARREGAMENTO DOS DADOS (DATA/PROCESSED)
# =============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
pasta_processed = os.path.join(BASE_DIR, "data", "processed")
os.makedirs(pasta_processed, exist_ok=True)

arquivos_candidatos = [
    os.path.join(pasta_processed, 'msci_europe_11_countries_usd.xlsx'),
    os.path.join(pasta_processed, 'msci_europe_11_countries_usd.csv'),
    os.path.join(pasta_processed, 'msci_europe_consolidado.xlsx')
]

caminho_dados = next((f for f in arquivos_candidatos if os.path.exists(f)), None)

if caminho_dados is None:
    busca_generica = (
        glob.glob(os.path.join(pasta_processed, "*msci*.xlsx")) +
        glob.glob(os.path.join(pasta_processed, "*europe*.xlsx")) +
        glob.glob(os.path.join(pasta_processed, "*msci*.csv"))
    )
    if busca_generica:
        caminho_dados = busca_generica[0]

if caminho_dados is None:
    raise FileNotFoundError(
        f"\n[ERRO] Base de dados não encontrada em:\n{pasta_processed}\n"
        "Execute primeiro o script 'consolidar_dados.py' para gerar a base em 'data/processed/'."
    )

print(f"-> Carregando base de dados de 'data/processed/': {os.path.basename(caminho_dados)}")

if caminho_dados.endswith(('.xlsx', '.xls')):
    df_bruto = pd.read_excel(caminho_dados)
else:
    df_bruto = pd.read_csv(caminho_dados)

# Padronização do índice temporal
col_dt = [c for c in df_bruto.columns if any(p in str(c).lower() for p in ['date', 'data', 'mes', 'month'])]
if col_dt:
    df_bruto[col_dt[0]] = pd.to_datetime(df_bruto[col_dt[0]])
    df_precos = df_bruto.set_index(col_dt[0])
else:
    df_bruto.iloc[:, 0] = pd.to_datetime(df_bruto.iloc[:, 0])
    df_precos = df_bruto.set_index(df_bruto.columns[0])

df_precos = df_precos.sort_index().select_dtypes(include=[np.number])
df_precos = df_precos.loc[:'2025-12-31'].dropna()

paises = list(df_precos.columns)
n_ativos = len(paises)
taxa_livre_risco = 0.03  # 3.0% a.a.

# Retornos logarítmicos mensais e anualizados
retornos_mensais = np.log(df_precos / df_precos.shift(1)).dropna()
retorno_mercado = retornos_mensais.mean(axis=1)  # Proxy do Índice de Mercado (MSCI Europe)

mu_anual = retornos_mensais.mean().values * 12
sigma_matriz = retornos_mensais.cov().values * 12

print(f"✓ Base carregada com {df_precos.shape[0]} meses e {n_ativos} países.")

# =============================================================================
# 2. MOTOR DE OTIMIZAÇÃO (MARKOWITZ)
# =============================================================================
def calc_perf(w):
    ret = np.dot(w, mu_anual)
    vol = np.sqrt(np.dot(w.T, np.dot(sigma_matriz, w)))
    sharpe = (ret - taxa_livre_risco) / vol if vol > 0 else 0
    return ret, vol, sharpe

restricao_soma = {'type': 'eq', 'fun': lambda w: np.sum(w) - 1.0}
limites_pesos = tuple((0.0, 1.0) for _ in range(n_ativos))
w_inicial = np.full(n_ativos, 1.0 / n_ativos)

# GMV
res_gmv = sco.minimize(lambda w: np.sqrt(np.dot(w.T, np.dot(sigma_matriz, w))), w_inicial, method='SLSQP', bounds=limites_pesos, constraints=restricao_soma)
w_gmv = res_gmv.x
ret_gmv, vol_gmv, sharpe_gmv = calc_perf(w_gmv)

# Tangência (Máximo Sharpe)
res_sharpe = sco.minimize(lambda w: -calc_perf(w)[2], w_inicial, method='SLSQP', bounds=limites_pesos, constraints=restricao_soma)
w_tang = res_sharpe.x
ret_tang, vol_tang, sharpe_tang = calc_perf(w_tang)

# 1/N
w_1n = np.full(n_ativos, 1.0 / n_ativos)
ret_1n, vol_1n, sharpe_1n = calc_perf(w_1n)

# Curva da Fronteira Eficiente
retornos_fronteira = np.linspace(ret_gmv, max(mu_anual) * 0.999, 50)
volatilidades_fronteira = []
rets_validos = []

for r_alvo in retornos_fronteira:
    res = sco.minimize(
        lambda w: np.sqrt(np.dot(w.T, np.dot(sigma_matriz, w))),
        w_inicial, method='SLSQP', bounds=limites_pesos,
        constraints=[restricao_soma, {'type': 'eq', 'fun': lambda w, r=r_alvo: np.dot(w, mu_anual) - r}]
    )
    if res.success:
        volatilidades_fronteira.append(res.fun)
        rets_validos.append(r_alvo)

# =============================================================================
# 3. GERADOR DO EXCEL COM DUAS ABAS: MARKOWITZ E SHARPE (SEM SOBREPOSIÇÃO)
# =============================================================================
def exportar_planilha_duas_abas(caminho_saida=None):
    if caminho_saida is None:
        caminho_saida = os.path.join(pasta_processed, 'Planilha_Markowitz_e_Sharpe_Europa.xlsx')
        
    wb = openpyxl.Workbook()
    
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    accent_fill = PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid")
    font_bold = Font(bold=True)
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )

    n_obs = len(retornos_mensais)
    row_end_data = 3 + n_obs
    col_mkt_num = 2 + n_ativos  # Coluna M (13)
    col_mkt_let = get_column_letter(col_mkt_num)
    
    start_mat_col = 16  # Início na Coluna P, deixando N e O para respiro visual

    def preencher_tabela_retornos(ws):
        ws.cell(1, 1, "RETORNOS MENSAIS EM US$ DOLAR (FECHAMENTO)").font = font_bold
        ws.cell(3, 1, "Data").font = font_bold
        
        for i, p in enumerate(paises):
            ws.cell(2, i + 2, i + 1).alignment = Alignment(horizontal="center")
            c = ws.cell(3, i + 2, p)
            c.font, c.fill = font_bold, header_fill
            
        ws.cell(2, col_mkt_num, "Mkt").alignment = Alignment(horizontal="center")
        c_mkt = ws.cell(3, col_mkt_num, "IBVSP_Proxy")
        c_mkt.font, c_mkt.fill = font_bold, accent_fill

        for r_idx, (dt, row_vals) in enumerate(retornos_mensais.iterrows()):
            row_num = 4 + r_idx
            ws.cell(row_num, 1, dt.strftime('%Y-%m-%d'))
            for c_idx, val in enumerate(row_vals):
                c = ws.cell(row_num, c_idx + 2, val)
                c.number_format = "0.0000"
                c.border = thin_border
            c_m = ws.cell(row_num, col_mkt_num, retorno_mercado.iloc[r_idx])
            c_m.number_format = "0.0000"
            c_m.border = thin_border

    # =========================================================================
    # ABA 1: MARKOWITZ
    # =========================================================================
    ws_m = wb.active
    ws_m.title = "Markowitz"
    preencher_tabela_retornos(ws_m)
    
    # Matriz 1: Covariância Plena (Linhas 3 a 13)
    start_cov_row = 3
    ws_m.cell(start_cov_row - 1, start_mat_col - 1, "Matriz Covariância").font = font_bold
    for j, p in enumerate(paises):
        c = ws_m.cell(start_cov_row - 1, start_mat_col + j, p)
        c.font, c.fill = font_bold, header_fill
        c.alignment = Alignment(horizontal="center")
        
    for i, p_row in enumerate(paises):
        r_num = start_cov_row + i
        c_label = ws_m.cell(r_num, start_mat_col - 1, p_row)
        c_label.font, c_label.fill = font_bold, header_fill
        col_i_let = get_column_letter(2 + i)
        
        for j, p_col in enumerate(paises):
            col_j_let = get_column_letter(2 + j)
            cell = ws_m.cell(r_num, start_mat_col + j)
            cell.value = f"=_xlfn.COVARIANCE.S(${col_i_let}$4:${col_i_let}${row_end_data}, {col_j_let}$4:{col_j_let}${row_end_data})"
            cell.number_format = "0.00000"
            cell.border = thin_border

    # Matriz 2: Covariância com Pesos (Linhas 16 a 26)
    start_wcov_row = start_cov_row + n_ativos + 2  # 3 + 11 + 2 = 16
    ws_m.cell(start_wcov_row - 1, start_mat_col - 1, "Matriz Covariância c/ Pesos").font = font_bold
    for j, p in enumerate(paises):
        c = ws_m.cell(start_wcov_row - 1, start_mat_col + j, p)
        c.font, c.fill = font_bold, header_fill
        c.alignment = Alignment(horizontal="center")
        
    c_somat = ws_m.cell(start_wcov_row - 1, start_mat_col + n_ativos, "Somatório linhas")
    c_somat.font, c_somat.fill = font_bold, header_fill
    
    ret_start_row = 32
    weight_start_row = 45
    weight_col_let = get_column_letter(start_mat_col)
    
    for i in range(n_ativos):
        r_num = start_wcov_row + i
        c_label = ws_m.cell(r_num, start_mat_col - 1, paises[i])
        c_label.font, c_label.fill = font_bold, header_fill
        
        row_w_cell = f"${weight_col_let}${weight_start_row + i}"
        cov_r_num = start_cov_row + i
        
        for j in range(n_ativos):
            col_w_cell = f"${weight_col_let}${weight_start_row + j}"
            cov_c_let = get_column_letter(start_mat_col + j)
            cov_cell = f"{cov_c_let}{cov_r_num}"
            
            cell = ws_m.cell(r_num, start_mat_col + j)
            cell.value = f"={col_w_cell}*{row_w_cell}*{cov_cell}"
            cell.number_format = "0.00000"
            cell.border = thin_border
            
        c_first = get_column_letter(start_mat_col)
        c_last = get_column_letter(start_mat_col + n_ativos - 1)
        c_sum = ws_m.cell(r_num, start_mat_col + n_ativos)
        c_sum.value = f"=SUM({c_first}{r_num}:{c_last}{r_num})"
        c_sum.number_format = "0.00000"
        c_sum.font, c_sum.border = font_bold, thin_border

    # Célula Var real na soma da matriz (Linha 28)
    var_sum_row = start_wcov_row + n_ativos  # 16 + 11 = 27
    ws_m.cell(var_sum_row + 1, start_mat_col + n_ativos - 1, "soma").font = font_bold
    sum_col_let = get_column_letter(start_mat_col + n_ativos)
    c_var_total = ws_m.cell(var_sum_row + 1, start_mat_col + n_ativos)
    c_var_total.value = f"=SUM({sum_col_let}{start_wcov_row}:{sum_col_let}{var_sum_row})"
    c_var_total.font, c_var_total.number_format = font_bold, "0.00000"
    ws_m.cell(var_sum_row + 2, start_mat_col + n_ativos - 1, "xixj cov(Ri;Rj)").font = font_bold

    # Retornos Médios (Linhas 32 a 42)
    ws_m.cell(ret_start_row - 1, start_mat_col - 1, "Retornos Médios").font = font_bold
    for i, p in enumerate(paises):
        r_num = ret_start_row + i
        ws_m.cell(r_num, start_mat_col - 1, f"R{i+1} ({p})")
        col_let = get_column_letter(2 + i)
        c_r = ws_m.cell(r_num, start_mat_col)
        c_r.value = f"=AVERAGE({col_let}4:{col_let}{row_end_data})"
        c_r.number_format = "0.0000%"
        
    # Pesos da Carteira (Linhas 45 a 55)
    ws_m.cell(weight_start_row - 1, start_mat_col - 1, "Pesos da Carteira").font = font_bold
    ws_m.cell(weight_start_row - 1, start_mat_col, "Variáveis Solver").font = font_bold
    for i, p in enumerate(paises):
        r_num = weight_start_row + i
        ws_m.cell(r_num, start_mat_col - 1, f"X{i+1} ({p})")
        c_w = ws_m.cell(r_num, start_mat_col)
        c_w.value = 1.0 / n_ativos
        c_w.number_format = "0.0000%"
        c_w.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        
    # Controles do Solver (Linhas 56 a 60)
    solver_row = weight_start_row + n_ativos
    ws_m.cell(solver_row, start_mat_col - 1, "SOMA X").font = font_bold
    c_sx = ws_m.cell(solver_row, start_mat_col)
    c_sx.value = f"=SUM({weight_col_let}{weight_start_row}:{weight_col_let}{weight_start_row + n_ativos - 1})"
    c_sx.font, c_sx.number_format = font_bold, "0.0000%"
    
    ws_m.cell(solver_row + 1, start_mat_col - 1, "Rp real").font = font_bold
    c_rp = ws_m.cell(solver_row + 1, start_mat_col)
    c_rp.value = f"=SUMPRODUCT({weight_col_let}{ret_start_row}:{weight_col_let}{ret_start_row + n_ativos - 1}, {weight_col_let}{weight_start_row}:{weight_col_let}{weight_start_row + n_ativos - 1})"
    c_rp.font, c_rp.number_format = font_bold, "0.0000%"
    
    ws_m.cell(solver_row + 2, start_mat_col - 1, "Rp desejado").font = font_bold
    c_des = ws_m.cell(solver_row + 2, start_mat_col)
    c_des.value = 0.0065
    c_des.number_format = "0.0000%"
    c_des.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    
    ws_m.cell(solver_row + 3, start_mat_col - 1, "Var real").font = font_bold
    c_vr = ws_m.cell(solver_row + 3, start_mat_col)
    c_vr.value = f"={sum_col_let}{var_sum_row + 1}"
    c_vr.font, c_vr.number_format = font_bold, "0.00000"
    
    ws_m.cell(solver_row + 4, start_mat_col - 1, "Desvio real").font = font_bold
    c_dr = ws_m.cell(solver_row + 4, start_mat_col)
    c_dr.value = f"=SQRT({weight_col_let}{solver_row + 3})"
    c_dr.font, c_dr.number_format = font_bold, "0.0000%"

    # =========================================================================
    # ABA 2: SHARPE (ORGANIZAÇÃO MODULAR SEM SOBREPOSIÇÃO)
    # =========================================================================
    ws_s = wb.create_sheet(title="Sharpe")
    preencher_tabela_retornos(ws_s)
    
    # BLOCO 1: Parâmetros de Mercado e Ativos (Linhas 2 a 14)
    ws_s.cell(2, 15, "Rm").font = font_bold
    c_rm = ws_s.cell(2, 16)  # P2
    c_rm.value = f"=AVERAGE({col_mkt_let}4:{col_mkt_let}{row_end_data})"
    c_rm.number_format, c_rm.font = "0.0000%", font_bold
    
    ws_s.cell(2, 18, "VarM").font = font_bold
    c_varm = ws_s.cell(2, 19)  # S2
    c_varm.value = f"=_xlfn.VAR.S({col_mkt_let}4:{col_mkt_let}{row_end_data})"
    c_varm.number_format, c_varm.font = "0.00000", font_bold
    
    ws_s.cell(3, 15, "Ativo").font = font_bold
    ws_s.cell(3, 16, "Retorno (Ri)").font = font_bold
    ws_s.cell(3, 18, "Var (Ri)").font = font_bold
    ws_s.cell(3, 20, "Cov (Ri, M)").font = font_bold
    ws_s.cell(3, 22, "Beta (Bi)").font = font_bold
    ws_s.cell(3, 24, "Alfa (Ai)").font = font_bold
    ws_s.cell(3, 26, "Resíduo (σ²ei)").font = font_bold

    for i, p in enumerate(paises):
        r_num = 4 + i  # Linhas 4 a 14
        col_let = get_column_letter(2 + i)
        
        ws_s.cell(r_num, 15, f"R{i+1} ({p})")
        # Ri
        c_r = ws_s.cell(r_num, 16, f"=AVERAGE({col_let}4:{col_let}{row_end_data})")
        c_r.number_format = "0.0000%"
        # Var(Ri)
        c_v = ws_s.cell(r_num, 18, f"=_xlfn.VAR.S({col_let}4:{col_let}{row_end_data})")
        c_v.number_format = "0.00000"
        # Cov(Ri, M)
        c_cov = ws_s.cell(r_num, 20, f"=_xlfn.COVARIANCE.S({col_let}4:{col_let}{row_end_data}, ${col_mkt_let}$4:${col_mkt_let}${row_end_data})")
        c_cov.number_format = "0.00000"
        # Beta (Bi) = Cov / VarM
        c_b = ws_s.cell(r_num, 22, f"=T{r_num}/$S$2")
        c_b.number_format, c_b.font = "0.0000", font_bold
        # Alfa (Ai) = Ri - Beta * Rm
        c_a = ws_s.cell(r_num, 24, f"=P{r_num}-V{r_num}*$P$2")
        c_a.number_format = "0.0000%"
        # Resíduo = Var(Ri) - Beta^2 * VarM
        c_res = ws_s.cell(r_num, 26, f"=R{r_num}-(V{r_num}^2)*$S$2")
        c_res.number_format = "0.00000"

    # BLOCO 2: Pesos da Carteira e Solver (Linhas 16 a 27)
    ws_s.cell(16, 15, "Pesos da Carteira").font = font_bold
    ws_s.cell(16, 16, "Variáveis Solver").font = font_bold
    for i, p in enumerate(paises):
        r_num = 17 + i  # Linhas 17 a 27
        ws_s.cell(r_num, 15, f"X{i+1} ({p})")
        c_w = ws_s.cell(r_num, 16, 1.0 / n_ativos)
        c_w.number_format = "0.0000%"
        c_w.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # Controles do Solver posicionados ao lado (Cols R e S)
    ws_s.cell(17, 18, "Beta Carteira").font = font_bold
    c_bp = ws_s.cell(17, 19, "=SUMPRODUCT(P17:P27, V4:V14)")
    c_bp.font, c_bp.number_format = font_bold, "0.0000"
    
    ws_s.cell(18, 18, "Alfa Carteira").font = font_bold
    c_ap = ws_s.cell(18, 19, "=SUMPRODUCT(P17:P27, X4:X14)")
    c_ap.font, c_ap.number_format = font_bold, "0.0000%"
    
    ws_s.cell(19, 18, "SOMA X").font = font_bold
    c_sx_s = ws_s.cell(19, 19, "=SUM(P17:P27)")
    c_sx_s.font, c_sx_s.number_format = font_bold, "0.0000%"
    
    ws_s.cell(20, 18, "Rp real").font = font_bold
    c_rp_s = ws_s.cell(20, 19, "=S18+S17*$P$2")
    c_rp_s.font, c_rp_s.number_format = font_bold, "0.0000%"
    
    ws_s.cell(21, 18, "Rp desejado").font = font_bold
    c_des_s = ws_s.cell(21, 19, 0.0065)
    c_des_s.number_format = "0.0000%"
    c_des_s.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    
    ws_s.cell(22, 18, "Var real").font = font_bold
    c_vr_s = ws_s.cell(22, 19, "=AA56")  # Vinculada à soma total da matriz ponderada abaixo
    c_vr_s.font, c_vr_s.number_format = font_bold, "0.00000"
    
    ws_s.cell(23, 18, "Desvio real").font = font_bold
    c_dr_s = ws_s.cell(23, 19, "=SQRT(S22)")
    c_dr_s.font, c_dr_s.number_format = font_bold, "0.0000%"

    # BLOCO 3: Matriz de Covariância de Sharpe (Linhas 29 a 40)
    sh_cov_start = 30
    ws_s.cell(sh_cov_start - 1, 15, "Matriz Covariância (Sharpe SIM)").font = font_bold
    for j, p in enumerate(paises):
        c = ws_s.cell(sh_cov_start - 1, 16 + j, p)
        c.font, c.fill = font_bold, header_fill
        c.alignment = Alignment(horizontal="center")
        
    for i in range(n_ativos):
        r_num = sh_cov_start + i  # Linhas 30 a 40
        c_label = ws_s.cell(r_num, 15, paises[i])
        c_label.font, c_label.fill = font_bold, header_fill
        bi_cell = f"$V${4 + i}"
        
        for j in range(n_ativos):
            bj_cell = f"$V${4 + j}"
            cell = ws_s.cell(r_num, 16 + j)
            cell.value = f"={bi_cell}*{bj_cell}*$S$2"
            cell.number_format = "0.00000"
            cell.border = thin_border

    # BLOCO 4: Matriz de Covariância com Pesos (Linhas 43 a 54)
    sh_wcov_start = 44
    ws_s.cell(sh_wcov_start - 1, 15, "Matriz Covariância c/ Pesos (Sharpe)").font = font_bold
    for j, p in enumerate(paises):
        c = ws_s.cell(sh_wcov_start - 1, 16 + j, p)
        c.font, c.fill = font_bold, header_fill
        c.alignment = Alignment(horizontal="center")
        
    c_somat_s = ws_s.cell(sh_wcov_start - 1, 16 + n_ativos, "Somatório linhas")
    c_somat_s.font, c_somat_s.fill = font_bold, header_fill
    
    for i in range(n_ativos):
        r_num = sh_wcov_start + i  # Linhas 44 a 54
        c_label = ws_s.cell(r_num, 15, paises[i])
        c_label.font, c_label.fill = font_bold, header_fill
        
        row_w_cell = f"$P${17 + i}"
        cov_r_num = sh_cov_start + i
        
        for j in range(n_ativos):
            col_w_cell = f"$P${17 + j}"
            cov_c_let = get_column_letter(16 + j)
            cov_cell = f"{cov_c_let}{cov_r_num}"
            
            cell = ws_s.cell(r_num, 16 + j)
            cell.value = f"={cov_cell}*{row_w_cell}*{col_w_cell}"
            cell.number_format = "0.00000"
            cell.border = thin_border
            
        c_first = get_column_letter(16)
        c_last = get_column_letter(16 + n_ativos - 1)
        c_sum = ws_s.cell(r_num, 16 + n_ativos)
        c_sum.value = f"=SUM({c_first}{r_num}:{c_last}{r_num})"
        c_sum.number_format = "0.00000"
        c_sum.font, c_sum.border = font_bold, thin_border
        
    # Linha da Soma Final da Variância (Linha 56)
    var_sum_row_s = 56
    ws_s.cell(var_sum_row_s, 16 + n_ativos - 1, "soma").font = font_bold
    sum_col_let_s = get_column_letter(16 + n_ativos)
    c_vtot_s = ws_s.cell(var_sum_row_s, 16 + n_ativos)
    c_vtot_s.value = f"=SUM({sum_col_let_s}44:{sum_col_let_s}54)"
    c_vtot_s.font, c_vtot_s.number_format = font_bold, "0.00000"
    ws_s.cell(var_sum_row_s + 1, 16 + n_ativos - 1, "xixj cov(Ri;Rj)").font = font_bold

    # Formatação de larguras das colunas
    for ws_cur in [ws_m, ws_s]:
        for col in ws_cur.columns:
            col_letter = get_column_letter(col[0].column)
            if col[0].column == 1:
                ws_cur.column_dimensions[col_letter].width = 12
            elif col[0].column <= col_mkt_num:
                ws_cur.column_dimensions[col_letter].width = 13
            elif col[0].column in [14, 15]:
                ws_cur.column_dimensions[col_letter].width = 16
            else:
                ws_cur.column_dimensions[col_letter].width = 14
                
    wb.save(caminho_saida)
    return caminho_saida

# Gera a planilha inicialmente ao executar
caminho_planilha_gerada = exportar_planilha_duas_abas()
print(f"✓ Planilha Markowitz + Sharpe gerada com sucesso em:\n{caminho_planilha_gerada}")

# =============================================================================
# 4. INTERFACE INTERATIVA PYTHON (TKINTER & MATPLOTLIB)
# =============================================================================
class PainelMarkowitzInterativo:
    def __init__(self, root):
        self.root = root
        self.root.title("Modelo Markowitz & Sharpe Interativo - 11 Países da Europa")
        self.root.geometry("1300x790")
        
        self.modo_var = tk.StringVar(value="retorno")
        
        frame_esq = ttk.Frame(root, padding=12)
        frame_esq.pack(side="left", fill="y", padx=8, pady=8)
        
        ttk.Label(frame_esq, text="Otimização de Carteiras", font=("Helvetica", 14, "bold")).pack(anchor="w", pady=(0, 10))
        
        frame_modo = ttk.LabelFrame(frame_esq, text=" 1. Modo de Alvo ", padding=10)
        frame_modo.pack(fill="x", pady=5)
        
        ttk.Radiobutton(frame_modo, text="Fixar Retorno Desejado (Minimizar Risco)", variable=self.modo_var, value="retorno", command=self.atualizar_interface_modo).pack(anchor="w", pady=2)
        ttk.Radiobutton(frame_modo, text="Fixar Risco Desejado (Maximizar Retorno)", variable=self.modo_var, value="risco", command=self.atualizar_interface_modo).pack(anchor="w", pady=2)
        
        frame_inputs = ttk.LabelFrame(frame_esq, text=" 2. Definir Meta ", padding=10)
        frame_inputs.pack(fill="x", pady=8)
        
        self.lbl_slider = ttk.Label(frame_inputs, text="Retorno Desejado (% a.a.):", font=("Helvetica", 10, "bold"))
        self.lbl_slider.pack(anchor="w")
        
        self.slider = ttk.Scale(frame_inputs, from_=ret_gmv * 100, to=max(mu_anual) * 100, orient="horizontal", command=self.ao_deslizar)
        self.slider.set(ret_tang * 100)
        self.slider.pack(fill="x", pady=6)
        
        frame_entry = ttk.Frame(frame_inputs)
        frame_entry.pack(fill="x", pady=2)
        
        self.entry_val = ttk.Entry(frame_entry, width=10, font=("Helvetica", 10, "bold"))
        self.entry_val.insert(0, f"{ret_tang * 100:.2f}")
        self.entry_val.pack(side="left")
        self.entry_val.bind("<Return>", lambda e: self.ao_digitar())
        
        btn_aplicar = ttk.Button(frame_entry, text="Otimizar", command=self.ao_digitar)
        btn_aplicar.pack(side="left", padx=6)
        
        frame_kpi = ttk.LabelFrame(frame_esq, text=" 3. Performance da Carteira ", padding=10)
        frame_kpi.pack(fill="x", pady=8)
        
        self.lbl_ret = ttk.Label(frame_kpi, text="Retorno Esperado: --", font=("Helvetica", 10))
        self.lbl_ret.pack(anchor="w", pady=1)
        self.lbl_vol = ttk.Label(frame_kpi, text="Volatilidade (Risco): --", font=("Helvetica", 10))
        self.lbl_vol.pack(anchor="w", pady=1)
        self.lbl_sharpe = ttk.Label(frame_kpi, text="Índice de Sharpe: --", font=("Helvetica", 10, "bold"), foreground="#0d6efd")
        self.lbl_sharpe.pack(anchor="w", pady=2)
        
        frame_tree = ttk.LabelFrame(frame_esq, text=" 4. Alocação Ótima (% por País) ", padding=8)
        frame_tree.pack(fill="both", expand=True, pady=6)
        
        self.tree = ttk.Treeview(frame_tree, columns=("Pais", "Peso"), show="headings", height=11)
        self.tree.heading("Pais", text="País")
        self.tree.heading("Peso", text="Peso (%)")
        self.tree.column("Pais", width=110, anchor="w")
        self.tree.column("Peso", width=80, anchor="e")
        self.tree.pack(fill="both", expand=True)
        
        btn_export = ttk.Button(frame_esq, text="Atualizar Planilha Excel (Markowitz & Sharpe)", command=self.ao_exportar_excel)
        btn_export.pack(fill="x", pady=6)

        frame_dir = ttk.Frame(root)
        frame_dir.pack(side="right", fill="both", expand=True, padx=8, pady=8)
        
        self.fig, (self.ax_front, self.ax_bar) = plt.subplots(2, 1, figsize=(7.5, 8.5), gridspec_kw={'height_ratios': [1.3, 1]})
        self.fig.tight_layout(pad=3.2)
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=frame_dir)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        
        self.otimizar_e_atualizar(ret_tang, "retorno")
        
    def atualizar_interface_modo(self):
        modo = self.modo_var.get()
        if modo == "retorno":
            self.lbl_slider.config(text="Retorno Desejado (% a.a.):")
            self.slider.config(from_=ret_gmv * 100, to=max(mu_anual) * 100)
            self.slider.set(ret_tang * 100)
            self.entry_val.delete(0, tk.END)
            self.entry_val.insert(0, f"{ret_tang * 100:.2f}")
            self.otimizar_e_atualizar(ret_tang, "retorno")
        else:
            self.lbl_slider.config(text="Risco / Volatilidade Desejada (% a.a.):")
            self.slider.config(from_=vol_gmv * 100, to=max(np.sqrt(np.diag(sigma_matriz))) * 100)
            self.slider.set(vol_tang * 100)
            self.entry_val.delete(0, tk.END)
            self.entry_val.insert(0, f"{vol_tang * 100:.2f}")
            self.otimizar_e_atualizar(vol_tang, "risco")
            
    def ao_deslizar(self, val):
        alvo = float(val) / 100.0
        self.entry_val.delete(0, tk.END)
        self.entry_val.insert(0, f"{float(val):.2f}")
        self.otimizar_e_atualizar(alvo, self.modo_var.get())
        
    def ao_digitar(self):
        try:
            val = float(self.entry_val.get().replace(',', '.')) / 100.0
            self.slider.set(val * 100)
            self.otimizar_e_atualizar(val, self.modo_var.get())
        except ValueError:
            messagebox.showerror("Valor Inválido", "Digite um valor numérico válido (ex: 8.5).")

    def otimizar_e_atualizar(self, alvo, modo):
        if modo == "retorno":
            alvo_ajustado = np.clip(alvo, ret_gmv, max(mu_anual) * 0.999)
            res = sco.minimize(
                lambda w: np.sqrt(np.dot(w.T, np.dot(sigma_matriz, w))),
                w_inicial, method='SLSQP', bounds=limites_pesos,
                constraints=[restricao_soma, {'type': 'eq', 'fun': lambda w: np.dot(w, mu_anual) - alvo_ajustado}]
            )
            w_otimo = res.x
        else:
            alvo_ajustado = max(alvo, vol_gmv)
            res = sco.minimize(
                lambda w: -np.dot(w, mu_anual),
                w_inicial, method='SLSQP', bounds=limites_pesos,
                constraints=[restricao_soma, {'type': 'ineq', 'fun': lambda w: (alvo_ajustado ** 2) - np.dot(w.T, np.dot(sigma_matriz, w))}]
            )
            w_otimo = res.x
            
        r_p, v_p, s_p = calc_perf(w_otimo)
        
        self.lbl_ret.config(text=f"Retorno Esperado:  {r_p:.2%}")
        self.lbl_vol.config(text=f"Volatilidade (Risco):  {v_p:.2%}")
        self.lbl_sharpe.config(text=f"Índice de Sharpe:    {s_p:.3f}")
        
        for row in self.tree.get_children():
            self.tree.delete(row)
        for p, w in zip(paises, w_otimo):
            self.tree.insert("", "end", values=(p, f"{w * 100:.2f}%"))
            
        self.desenhar_graficos(r_p, v_p, w_otimo)
        
    def desenhar_graficos(self, r_p, v_p, w_otimo):
        self.ax_front.clear()
        self.ax_bar.clear()
        
        self.ax_front.plot(volatilidades_fronteira, rets_validos, 'b-', linewidth=2.2, label='Fronteira Eficiente')
        
        x_cml = np.linspace(0.0, max(volatilidades_fronteira) * 1.15, 30)
        self.ax_front.plot(x_cml, taxa_livre_risco + sharpe_tang * x_cml, 'r--', alpha=0.7, label='CML (Rf 3.0%)')
        
        self.ax_front.scatter(vol_gmv, ret_gmv, marker='*', color='green', s=180, zorder=4, label='GMV')
        self.ax_front.scatter(vol_tang, ret_tang, marker='^', color='red', s=150, zorder=4, label='Tangência (Sharpe)')
        self.ax_front.scatter(vol_1n, ret_1n, marker='o', color='purple', s=90, zorder=4, label='1/N')
        self.ax_front.scatter(v_p, r_p, marker='D', color='#FFD700', edgecolor='black', s=160, zorder=5, label=f'Carteira Selecionada (Ret {r_p:.1%})')
        
        self.ax_front.set_title("Fronteira Eficiente de Markowitz & Alocação Selecionada", fontsize=11, fontweight='bold')
        self.ax_front.set_xlabel("Volatilidade Anualizada ($\sigma_p$)", fontsize=9)
        self.ax_front.set_ylabel("Retorno Esperado ($E(R_p)$)", fontsize=9)
        self.ax_front.set_xlim(left=0.0)
        self.ax_front.grid(True, linestyle='--', alpha=0.4)
        self.ax_front.legend(loc='upper left', fontsize=8)
        
        cores = ['#2b5c8f' if w > 0.005 else '#cccccc' for w in w_otimo]
        bars = self.ax_bar.barh(paises, w_otimo * 100, color=cores, edgecolor='#1b3a5b', height=0.6)
        self.ax_bar.set_title("Composição da Carteira Selecionada (%)", fontsize=11, fontweight='bold')
        self.ax_bar.set_xlabel("Peso Alocado (%)", fontsize=9)
        self.ax_bar.set_xlim(0, 105)
        self.ax_bar.grid(True, linestyle='--', alpha=0.4, axis='x')
        
        for bar in bars:
            w_val = bar.get_width()
            if w_val > 1.0:
                self.ax_bar.text(w_val + 1.2, bar.get_y() + bar.get_height()/2, f"{w_val:.1f}%", va='center', fontsize=8, fontweight='bold')
                
        self.canvas.draw()
        
    def ao_exportar_excel(self):
        caminho = exportar_planilha_duas_abas()
        messagebox.showinfo("Sucesso", f"Planilha com as abas 'Markowitz' e 'Sharpe' atualizada com sucesso!\n\nSalva em:\n{caminho}")

# =============================================================================
# 5. EXECUÇÃO
# =============================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = PainelMarkowitzInterativo(root)
    root.mainloop()