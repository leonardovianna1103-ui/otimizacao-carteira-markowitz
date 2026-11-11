# 📈 Otimizador de Carteiras: Modelo de Markowitz & Sharpe (Europa Desenvolvida)

Aplicação quantitativa em Python para extração, consolidação e otimização de portfólios baseada na **Teoria Moderna de Portfólio (Markowitz)** e no **Modelo de Índice Único (Sharpe)**, aplicada a índices MSCI de 11 países desenvolvidos da Europa.

## 🎯 Funcionalidades

- **Pipeline ETL Automático (`consolidar_dados.py`)**:
  - Leitura dinâmica e extração de séries temporais de índices MSCI em formato Excel.
  - Alinhamento de calendário financeiro (fim de mês) e tratamento de períodos comuns para otimização.
  - Exportação da base limpa em `.csv` e `.xlsx`.

- **Motor de Otimização e Interface Interativa (`otimizador_carteira.py`)**:
  - Cálculo de retornos logarítmicos anualizados e matriz de covariância amostral.
  - Otimização numérica via algoritmo `SLSQP` (`scipy.optimize`):
    - **GMV**: Carteira de Mínima Variância Global.
    - **Tangência**: Maximização do Índice de Sharpe ($R_f = 3{,}0\%$ a.a.).
    - **Portfólio 1/N**: Alocação equiponderada como benchmark.
    - **Fronteira Eficiente & Linha de Mercado de Capitais (CML)**.
  - **Interface Gráfica (Tkinter & Matplotlib)**: Controle interativo via sliders para simulação de metas de retorno ou de risco em tempo real.
  - **Exportação Analítica (`openpyxl`)**: Geração automática de planilha formatada com fórmulas nativas do Excel (`COVARIANCE.S`, `SUMPRODUCT`, `VAR.S`) dividida em abas:
    - `Markowitz`: Matrizes de covariância plena e ponderada, retornos esperados e restrições prontas para uso com o Solver.
    - `Sharpe`: Decomposição de risco sistemático e idiossincrático (Alfa, Beta e variância residual).

## 🛠️ Tecnologias Utilizadas

- **Linguagem**: Python 3.10+
- **Bibliotecas**: `pandas`, `numpy`, `scipy`, `matplotlib`, `openpyxl`, `tkinter`

## 🚀 Como Executar

1. **Clone o repositório:**
   ```bash
   git clone [https://github.com/SEU_USUARIO/otimizacao-carteira-markowitz.git](https://github.com/SEU_USUARIO/otimizacao-carteira-markowitz.git)
   cd otimizacao-carteira-markowitz