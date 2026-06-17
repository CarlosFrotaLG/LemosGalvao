import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime
import unicodedata
import io 

# ==========================================
# CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(page_title="Sistema Financeiro | Lemos Galvão", layout="wide", initial_sidebar_state="expanded")

# ==========================================
# IDENTIDADE VISUAL & CSS AVANÇADO
# ==========================================
st.markdown("""
    <style>
        /* Títulos totalmente centralizados e com fonte corporativa */
        h1, h2, h3 { 
            font-family: 'Georgia', serif; 
            color: #0F1C2E; 
            text-align: center !important; 
        }
        
        .st-eb { background-color: #F7F4EF; }
        [data-testid="stMetricValue"] { color: #C9A84C !important; font-weight: bold; text-align: center; }
        [data-testid="stMetricLabel"] { text-align: center; }
        
        /* Botões Principais do Corpo da Página */
        .main .stButton>button {
            background-color: #0F1C2E; color: #F7F4EF !important;
            border-radius: 8px; border: 1px solid #C9A84C; transition: 0.3s; font-weight: bold;
        }
        .main .stButton>button:hover { background-color: #C9A84C; color: #0F1C2E !important; border: 1px solid #0F1C2E; }
        
        /* ==========================================
           MENU LATERAL MINIMALISTA (SEM CAIXAS/BORDAS)
           ========================================== */
        [data-testid="stSidebar"] { background-color: #F7F4EF; border-right: 2px solid #C9A84C; }
        
        /* Remove a caixa/borda do Expander */
        [data-testid="stSidebar"] [data-testid="stExpander"] {
            border: none !important;
            background-color: transparent !important;
            box-shadow: none !important;
        }
        [data-testid="stSidebar"] [data-testid="stExpander"] details {
            border: none !important;
        }
        [data-testid="stSidebar"] [data-testid="stExpander"] summary {
            background-color: transparent !important;
            border: none !important;
            padding-left: 0 !important;
        }
        
        /* Estilo do Título Maior no Menu */
        [data-testid="stSidebar"] [data-testid="stExpander"] details summary p {
            font-size: 15px !important; font-weight: bold !important; color: #0F1C2E !important; text-align: left !important;
        }
        [data-testid="stSidebar"] [data-testid="stExpander"] summary:hover p {
            color: #C9A84C !important;
        }
        
        /* Estilo dos Botões/Itens Menores */
        [data-testid="stSidebar"] .stButton > button {
            border: none !important; background-color: transparent !important;
            color: #3E3F3A !important; text-align: left !important;
            justify-content: flex-start !important; padding-left: 20px !important;
            font-size: 14px !important; box-shadow: none !important;
        }
        [data-testid="stSidebar"] .stButton > button:hover {
            color: #C9A84C !important; font-weight: bold !important; background-color: transparent !important;
        }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# FUNÇÕES DE FORMATAÇÃO E ESTILIZAÇÃO
# ==========================================
def formata_contabil(valor):
    try:
        if pd.isna(valor) or valor == '': return "-"
        return f"R$ {float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except: return valor

def normalizar_nome(nome):
    if pd.isna(nome) or nome == '': return ''
    nfd = unicodedata.normalize('NFD', str(nome))
    return nfd.encode('ascii', 'ignore').decode('utf8').upper().strip()

def converter_df_para_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Relatório')
    return output.getvalue()

def estilizar_tabela(df, func_pintar=None, tudo_centro=False):
    styler = df.style
    if func_pintar: styler = styler.apply(func_pintar, axis=1)
    
    if tudo_centro:
        styler = styler.set_properties(**{'text-align': 'center !important'})
    else:
        col_right = [c for c in df.columns if any(palavra in c.upper() for palavra in ['VALOR', 'PRODUZIDO', 'DIFERENÇA', 'CRÉDITO', 'PAGO', 'AJUSTE', 'FALTA'])]
        col_center = [c for c in df.columns if c not in col_right]
        if col_center: styler = styler.set_properties(subset=col_center, **{'text-align': 'center'})
        if col_right: styler = styler.set_properties(subset=col_right, **{'text-align': 'right'})
        
    styler = styler.set_table_styles([dict(selector='th', props=[('text-align', 'center !important')])])
    return styler

# ==========================================
# CARREGAMENTO DO BANCO DE DADOS
# ==========================================
@st.cache_data(ttl=600) 
def carregar_dados():
    conexao = sqlite3.connect('banco_consorcio.db')
    try: df_vendas = pd.read_sql_query("SELECT * FROM tb_vendas", conexao)
    except: df_vendas = pd.DataFrame()
    try: df_auditoria = pd.read_sql_query("SELECT * FROM tb_auditoria", conexao)
    except: df_auditoria = pd.DataFrame()
    try: df_comissao = pd.read_sql_query("SELECT * FROM tb_comissionamento", conexao)
    except: df_comissao = pd.DataFrame()
    try: df_parceiros = pd.read_sql_query("SELECT * FROM tb_parceiros", conexao)
    except: df_parceiros = pd.DataFrame()
    try: df_promo = pd.read_sql_query("SELECT * FROM tb_promocoes", conexao)
    except: df_promo = pd.DataFrame()
    conexao.close()
    return df_vendas, df_auditoria, df_comissao, df_parceiros, df_promo

df_vendas, df_auditoria, df_comissao, df_parceiros, df_promo = carregar_dados()

map_empresas = {
    "PLANEJAMENTO": "LG PLANEJAMENTO", "SOLUÇÕES": "LG SOLUÇÕES", "SOLUCOES": "LG SOLUÇÕES",
    "INTELIGÊNCIA": "LG INTELIGÊNCIA", "INTELIGENCIA": "LG INTELIGÊNCIA",
    "BELO HORIZONTE": "LG BELO HORIZONTE", "CURITIBA": "LG CURITIBA",
    "PLATAFORMA 1": "LG PLATAFORMA 1", "PLATAFORMA 2": "LG PLATAFORMA 2", "PLATAFORMA 3": "LG PLATAFORMA 3"
}

consultores_oficiais = [
    "ANDERSON", "ANDERSON / DENISE", "AURORA", "BRUNO BORGES", "CARLOS",
    "CÍCERO", "DENISE", "DOUGLAS", "EDUARDO HENRIQUE", "JOSÉ WAGNER",
    "JULIO AVIDUS", "LEANDRO LEMOS", "LEANDRO MÁXIMO", "LUANA", 
    "LUZ DA SERRA", "MÁRCIO PASSOS", "MASTROCOLA", "PEDRO GONTIJO", "RODOLPHO", "RUY", "TIAGO MENDO"
]
vip_list = [normalizar_nome(c) for c in consultores_oficiais]

if not df_auditoria.empty and 'EMPRESA RECEBEDORA' in df_auditoria.columns:
    df_auditoria['EMPRESA RECEBEDORA'] = df_auditoria['EMPRESA RECEBEDORA'].astype(str).str.strip().str.upper().replace(map_empresas)

if not df_vendas.empty and 'Empresa Vendedora' in df_vendas.columns:
    df_vendas['Empresa Vendedora'] = df_vendas['Empresa Vendedora'].astype(str).str.strip().str.upper().replace(map_empresas)

# ==========================================
# MOTORES DE CÁLCULO (INJEÇÃO DE CLIENTE)
# ==========================================
@st.cache_data
def processar_motor_comercial(df_cruzado, df_promo):
    registros = []
    if df_cruzado.empty: return pd.DataFrame()
    col_cliente = next((c for c in df_cruzado.columns if 'CLIENTE' in c.upper()), None)
    
    for index, row in df_cruzado.iterrows():
        contrato = row['Nº Contrato']
        data_venda = row['Data da Venda']
        vendedora = str(row['Empresa Vendedora']).strip()
        cliente_nome = str(row[col_cliente]).strip() if col_cliente and pd.notna(row[col_cliente]) else "Desconhecido"
        
        try: credito = float(row['Crédito Numérico'])
        except: credito = 0.0
        if pd.isnull(data_venda) or credito == 0: continue
            
        is_promo = False
        if not df_promo.empty and 'Nº Contrato' in df_promo.columns:
            match = df_promo[df_promo['Nº Contrato'].astype(str) == str(contrato)]
            if not match.empty and 'promo' in str(match.iloc[0].get('Status Promocional', '')).lower(): is_promo = True

        qtd_parcelas = 24 if is_promo else 10
        for i in range(1, qtd_parcelas + 1):
            data_prevista = data_venda + pd.DateOffset(months=i)
            if is_promo: taxa = 0.0015 
            else:
                try: taxa = float(str(row.get(f'P{i}', 0)).replace('%', '').replace(',', '.'))
                except: taxa = 0.0
            if credito * taxa > 0:
                registros.append({
                    'Nº Contrato': contrato, 'Cliente': cliente_nome, 'Data da Venda': data_venda, 'Parcela': f"{i}ª", 
                    'Data Prevista': data_prevista, 'Empresa Vendedora': vendedora,
                    'Recebedora': vendedora, 'Valor Previsto': credito * taxa, 'Tipo': 'Comercial'
                })
    return pd.DataFrame(registros).reset_index(drop=True)

@st.cache_data
def processar_motor_plataforma(df_cruzado, df_parceiros):
    registros = []
    if df_cruzado.empty: return pd.DataFrame()
    
    col_cliente = next((c for c in df_cruzado.columns if 'CLIENTE' in c.upper()), None)
    dict_p = {}
    
    if not df_parceiros.empty:
        df_p = df_parceiros.copy()
        df_p['DATA INÍCIO'] = pd.to_datetime(df_p['DATA INÍCIO'], errors='coerce')
        df_p['DATA FIM'] = pd.to_datetime(df_p['DATA FIM'], errors='coerce')
        for _, rp in df_p.iterrows():
            emp = str(rp.get('EMPRESA PARCEIRA', '')).strip().upper()
            emp = map_empresas.get(emp, emp)
            if emp not in dict_p: dict_p[emp] = []
            mae = str(rp.get('QUEM INDICOU (MÃE)', '-')).strip().upper()
            mae = map_empresas.get(mae, mae)
            dict_p[emp].append((rp['DATA INÍCIO'], rp['DATA FIM'], str(rp.get('NÍVEL DE PLATAFORMA', '')).strip(), mae))

    def buscar_mae(empresa, data_v):
        if pd.isnull(data_v) or empresa not in dict_p: return 'Sem Plataforma', '-'
        for di, dfim, niv, mae in dict_p[empresa]:
            if pd.notnull(di) and pd.notnull(dfim) and di <= data_v <= dfim: return niv, mae
        return 'Sem Plataforma', '-'

    for index, row in df_cruzado.iterrows():
        contrato = row['Nº Contrato']
        data_venda = row['Data da Venda']
        vendedora = str(row['Empresa Vendedora']).strip()
        cliente_nome = str(row[col_cliente]).strip() if col_cliente and pd.notna(row[col_cliente]) else "Desconhecido"
        
        try: credito = float(row['Crédito Numérico'])
        except: credito = 0.0
        if pd.isnull(data_venda) or credito == 0: continue
        
        if data_venda < pd.to_datetime('2024-01-01'):
            b1_pool = {13: 0.0050, 18: 0.0050}
            b2_pool = {20: 0.0025, 22: 0.0025, 24: 0.0025, 28: 0.0025}
        elif data_venda < pd.to_datetime('2025-01-01'):
            b1_pool = {p: 0.00125 for p in range(12, 20)}
            b2_pool = {p: 0.00125 for p in range(20, 28)}
        else:
            b1_pool = {p: 0.00125 for p in range(12, 20)}
            b2_pool = {p: 0.00125 for p in range(24, 32)}

        def add_p(p, tx, emp):
            if tx > 0:
                dt_prev = data_venda + pd.DateOffset(months=p)
                registros.append({
                    'Nº Contrato': contrato, 'Cliente': cliente_nome, 'Data da Venda': data_venda, 'Parcela': f"{p}ª", 
                    'Data Prevista': dt_prev, 'Empresa Vendedora': vendedora,
                    'Recebedora': emp, 'Valor Previsto': credito * tx, 'Tipo': 'Plataforma'
                })

        curr_emp = vendedora
        tx_paga_b1 = 0.0
        safe = 0
        while curr_emp and curr_emp != '-' and (b1_pool or b2_pool) and safe < 15:
            safe += 1
            niv, mae = buscar_mae(curr_emp, data_venda)
            if curr_emp == vendedora and niv == 'Sem Plataforma':
                nr = str(row.get('NÍVEL DA VENDEDORA', 'Sem Plataforma')).strip()
                mr = str(row.get('SUBPLATAFORMA MÃE', '-')).strip().upper()
                mr = map_empresas.get(mr, mr)
                if nr and nr != 'nan': niv = nr
                if mr and mr != 'nan': mae = mr
                
            nl = niv.lower()
            tx_perm = 0.0
            if 'sem' in nl or 'nenhuma' in nl: tx_perm = 0.0
            elif '0,25' in nl or '0.25' in nl: tx_perm = 0.0025
            elif '0,50' in nl or '0.50' in nl: tx_perm = 0.0050
            elif '0,75' in nl or '0.75' in nl: tx_perm = 0.0075
            elif '1%' in nl or '1,00' in nl or '1.00' in nl: tx_perm = 0.0100
            elif 'master' in nl or ('plataforma' in nl and 'sub' not in nl): tx_perm = 0.0100
            
            tx_rec_b1 = round(max(0.0, tx_perm - tx_paga_b1), 6)
            vazias = []
            for p in sorted(b1_pool.keys()):
                if tx_rec_b1 <= 0: break
                tx_cons = min(tx_rec_b1, b1_pool[p])
                add_p(p, tx_cons, curr_emp)
                b1_pool[p] = round(b1_pool[p] - tx_cons, 6)
                tx_rec_b1 = round(tx_rec_b1 - tx_cons, 6)
                tx_paga_b1 = round(tx_paga_b1 + tx_cons, 6)
                if b1_pool[p] <= 0: vazias.append(p)
            for p in vazias: del b1_pool[p]
                
            if 'master' in nl:
                for p in sorted(b1_pool.keys()): add_p(p, b1_pool[p], curr_emp)
                b1_pool.clear()
                for p in sorted(b2_pool.keys()): add_p(p, b2_pool[p], curr_emp)
                b2_pool.clear()
            curr_emp = mae
    return pd.DataFrame(registros).reset_index(drop=True)

df_motor = df_vendas.copy()
if not df_motor.empty and not df_comissao.empty:
    df_motor['Crédito Numérico'] = pd.to_numeric(df_motor['Crédito'], errors='coerce').fillna(0)
    df_motor['Data da Venda'] = pd.to_datetime(df_motor['Data da Venda'], errors='coerce')
    df_cruzado = pd.merge(df_motor, df_comissao, left_on='CHAVE (OCULTA)', right_on='CHAVE COMBINADA', how='left')
else:
    df_cruzado = pd.DataFrame()

# ==========================================
# NAVEGAÇÃO LATERAL (MENU MINIMALISTA)
# ==========================================
if 'pagina_atual' not in st.session_state:
    st.session_state['pagina_atual'] = "Página Inicial"

def navegar(pagina):
    st.session_state['pagina_atual'] = pagina

try:
    st.sidebar.image("logo_chaves.png", use_container_width=True) 
except:
    pass

st.sidebar.markdown("### MENU PRINCIPAL")

if st.sidebar.button("Página Inicial", use_container_width=True): navegar("Página Inicial")

with st.sidebar.expander("PRODUÇÃO", expanded=True):
    if st.button("Vendas Cadastradas", use_container_width=True, key="btn_prod1"): navegar("Vendas Cadastradas")
    if st.button("Produção Mensal", use_container_width=True, key="btn_prod2"): navegar("Produção Mensal")
    if st.button("Rank de Produção", use_container_width=True, key="btn_prod3"): navegar("Rank de Produção")

with st.sidebar.expander("CANCELAMENTOS", expanded=False):
    if st.button("Visão Geral", use_container_width=True, key="btn_canc1"): navegar("Visão Geral Cancelamentos")

with st.sidebar.expander("PREVISÃO DE RECEITA", expanded=False):
    if st.button("Motor Comercial", use_container_width=True, key="btn_prev1"): navegar("Motor Comercial")
    if st.button("Motor de Plataforma", use_container_width=True, key="btn_prev2"): navegar("Motor de Plataforma")
    if st.button("Receitas Unificadas", use_container_width=True, key="btn_prev3"): navegar("Receitas Unificadas")

with st.sidebar.expander("RECEBIMENTOS", expanded=False):
    if st.button("Cadastro de Recebimentos", use_container_width=True, key="btn_rec1"): navegar("Cadastro de Recebimentos")
    if st.button("Auditoria Definitiva", use_container_width=True, key="btn_rec2"): navegar("Auditoria Definitiva")

escolha = st.session_state['pagina_atual']

# ==========================================
# BLOCOS REUTILIZÁVEIS (RODAPÉ E FRASE)
# ==========================================
frase_producao = """
    <div style='text-align: center; margin-top: 50px; margin-bottom: 20px;'>
        <p style='color: #0F1C2E; font-style: italic; font-size: 1.1rem; font-family: "Georgia", serif;'>
        "Ser fator de soma ao maior número possível de pessoas que estejam culturalmente prontas para fazer as reflexões patrimoniais necessárias e, a partir delas, transformar sonhos e ambições em legados duradouros."
        </p>
    </div>
"""

rodape_html = """
    <br><br>
    <hr style='border-top: 3px solid #C9A84C;'>
    <div style='display: flex; flex-wrap: wrap; justify-content: space-between; color: #3E3F3A; font-size: 14px; font-weight: bold; padding-top: 10px;'>
        <div style='margin-bottom: 10px;'>Lemos Galvão – 2025 – Todos os direitos reservados</div>
        <div style='text-align: right;'>Edifício Capital Financial Center – SIG Quadra 4, lote 75 a 175 Bloco A Sala 4ME, Brasília – DF, 70610-440</div>
    </div>
"""

# ==========================================
# RENDERIZAÇÃO DA PÁGINA ESCOLHIDA
# ==========================================

# --- PÁGINA INICIAL ---
if escolha == "Página Inicial":
    st.markdown("<br><br>", unsafe_allow_html=True)
    try:
        st.image("capital_financial.png", use_container_width=True)
    except:
        pass
    
    st.markdown("""
        <div style='text-align: center; margin-top: 40px;'>
            <p style='color: #C9A84C; font-style: italic; font-size: 1.4rem; font-weight: bold; margin: 0 auto; font-family: "Georgia", serif; line-height: 1.5;'>
            "Toda jornada existem desafios, e a cada desafio segredos que antes eram desconhecidos,<br>tornam-se conhecimento e experiência."
            </p>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown(rodape_html, unsafe_allow_html=True)

# --- TODAS AS OUTRAS PÁGINAS ---
else:
    try:
        st.image("lemosgalvo_cover.png", use_container_width=True)
    except:
        pass
    
    st.divider()
    
    # --------------------------------------------------
    # MÓDULO PRODUÇÃO
    # --------------------------------------------------
    if escolha == "Vendas Cadastradas":
        st.markdown("<h1>Vendas Cadastradas</h1>", unsafe_allow_html=True)
        if not df_motor.empty: 
            df_view = df_motor.copy()
            
            col1, col2 = st.columns(2)
            empresas_unicas = ["Todas"] + list(df_view['Empresa Vendedora'].dropna().unique())
            f_emp = col1.selectbox("Filtrar Empresa:", empresas_unicas)
            if f_emp != "Todas": df_view = df_view[df_view['Empresa Vendedora'] == f_emp]
            
            cols_exibicao = ['Empresa Vendedora', 'Consultor', 'Data da Venda', 'Cliente', 'Nº Contrato', 'Crédito Numérico']
            colunas_reais = [c for c in cols_exibicao if c in df_view.columns]
            
            df_tabela = df_view[colunas_reais].copy()
            df_tabela['Data da Venda'] = df_tabela['Data da Venda'].dt.strftime('%d/%m/%Y')
            
            if 'Crédito Numérico' in df_tabela.columns:
                df_tabela['Crédito'] = df_tabela['Crédito Numérico'].apply(formata_contabil)
                df_tabela = df_tabela.drop(columns=['Crédito Numérico'])
            
            st.dataframe(estilizar_tabela(df_tabela), use_container_width=True, hide_index=True)
        else:
            st.warning("Nenhum dado de venda encontrado.")
        st.markdown(frase_producao, unsafe_allow_html=True)
        st.markdown(rodape_html, unsafe_allow_html=True)

    elif escolha == "Produção Mensal":
        st.markdown("<h1>Produção Mês a Mês</h1>", unsafe_allow_html=True)
        if not df_motor.empty:
            df_dash = df_motor.copy()
            df_dash['Mês/Ano'] = df_dash['Data da Venda'].dt.strftime('%m/%Y')
            
            c1, c2, c3 = st.columns(3)
            f_inicio = c1.date_input("Data Inicial:", df_dash['Data da Venda'].min().date())
            f_fim = c2.date_input("Data Final:", df_dash['Data da Venda'].max().date())
            f_emp = c3.selectbox("Empresa:", ["Todas"] + list(df_dash['Empresa Vendedora'].dropna().unique()))
            
            df_dash = df_dash[(df_dash['Data da Venda'].dt.date >= f_inicio) & (df_dash['Data da Venda'].dt.date <= f_fim)]
            if f_emp != "Todas": df_dash = df_dash[df_dash['Empresa Vendedora'] == f_emp]
            
            tabela_producao = pd.crosstab(index=df_dash['Empresa Vendedora'], columns=df_dash['Mês/Ano'], values=df_dash['Crédito Numérico'], aggfunc='sum').fillna(0)
            tabela_producao['TOTAL ACUMULADO'] = tabela_producao.sum(axis=1)
            tabela_producao = tabela_producao.sort_values(by='TOTAL ACUMULADO', ascending=False)
            st.dataframe(estilizar_tabela(tabela_producao.map(formata_contabil).reset_index()), use_container_width=True, hide_index=True)
        st.markdown(frase_producao, unsafe_allow_html=True)
        st.markdown(rodape_html, unsafe_allow_html=True)

    elif escolha == "Rank de Produção":
        st.markdown("<h1>Rank de Produção</h1>", unsafe_allow_html=True)
        if not df_motor.empty:
            df_rank = df_motor.copy()
            c1, c2, c3 = st.columns(3)
            f_inicio = c1.date_input("Data Inicial:", df_rank['Data da Venda'].min().date())
            f_fim = c2.date_input("Data Final:", df_rank['Data da Venda'].max().date())
            
            df_rank = df_rank[(df_rank['Data da Venda'].dt.date >= f_inicio) & (df_rank['Data da Venda'].dt.date <= f_fim)]
            
            col_consultor = next((col for col in df_rank.columns if 'CONSULT' in col.upper()), None)
            if col_consultor:
                f_cons = c3.selectbox("Consultor Específico:", ["Todos"] + list(df_rank[col_consultor].dropna().unique()))
                if f_cons != "Todos": df_rank = df_rank[df_rank[col_consultor] == f_cons]
            
            colA, colB = st.columns(2)
            with colA:
                st.markdown("<h3>Top Empresas</h3>", unsafe_allow_html=True)
                rank_emp = df_rank.groupby('Empresa Vendedora')['Crédito Numérico'].sum().reset_index().sort_values(by='Crédito Numérico', ascending=False)
                if not rank_emp.empty:
                    rank_emp.insert(0, 'Posição', range(1, len(rank_emp) + 1))
                    rank_emp['Posição'] = rank_emp['Posição'].astype(str) + "º"
                    rank_emp['Volume Produzido'] = rank_emp['Crédito Numérico'].apply(formata_contabil)
                    
                    cA1, cA2, cA3 = st.columns([1, 4, 1])
                    with cA2:
                        st.dataframe(estilizar_tabela(rank_emp[['Posição', 'Empresa Vendedora', 'Volume Produzido']]), use_container_width=True, hide_index=True)
                
            with colB:
                st.markdown("<h3>Top Consultores</h3>", unsafe_allow_html=True)
                if col_consultor:
                    rank_cons = df_rank.groupby(col_consultor)['Crédito Numérico'].sum().reset_index().sort_values(by='Crédito Numérico', ascending=False)
                    if not rank_cons.empty:
                        rank_cons.insert(0, 'Posição', range(1, len(rank_cons) + 1))
                        rank_cons['Posição'] = rank_cons['Posição'].astype(str) + "º"
                        rank_cons['Volume Produzido'] = rank_cons['Crédito Numérico'].apply(formata_contabil)
                        
                        cB1, cB2, cB3 = st.columns([1, 4, 1])
                        with cB2:
                            st.dataframe(estilizar_tabela(rank_cons[['Posição', col_consultor, 'Volume Produzido']]), use_container_width=True, hide_index=True)
        st.markdown(frase_producao, unsafe_allow_html=True)
        st.markdown(rodape_html, unsafe_allow_html=True)

    # --------------------------------------------------
    # MÓDULO CANCELAMENTOS
    # --------------------------------------------------
    elif escolha == "Visão Geral Cancelamentos":
        st.markdown("<h1>Gestão de Cancelamentos</h1>", unsafe_allow_html=True)
        col_status = next((col for col in df_motor.columns if 'STATUS' in col.upper() or 'SITUAÇÃO' in col.upper()), None)
        
        if not df_motor.empty and col_status:
            df_canc = df_motor[df_motor[col_status].astype(str).str.contains('Cancel|Distrat', case=False, na=False)].copy()
            
            resumo_cancelados = df_canc.groupby('Empresa Vendedora')['Crédito Numérico'].agg(['count', 'sum']).reset_index()
            resumo_cancelados.columns = ['Empresa', 'Qtd Cancelados', 'Volume Perdido']
            resumo_cancelados['Volume Perdido'] = resumo_cancelados['Volume Perdido'].apply(formata_contabil)
            
            c_canc1, c_canc2, c_canc3 = st.columns([1, 2, 1])
            with c_canc2: 
                st.dataframe(estilizar_tabela(resumo_cancelados), use_container_width=True, hide_index=True)
            
            st.divider()
            st.markdown("<h3>Listagem Detalhada</h3>", unsafe_allow_html=True)
            cols_canc = ['Data da Venda', 'Empresa Vendedora', 'Consultor', 'Cliente', 'Nº Contrato', 'Crédito Numérico']
            cols_existentes = [c for c in cols_canc if c in df_canc.columns]
            df_canc_tela = df_canc[cols_existentes].copy()
            if 'Data da Venda' in df_canc_tela: df_canc_tela['Data da Venda'] = df_canc_tela['Data da Venda'].dt.strftime('%d/%m/%Y')
            if 'Crédito Numérico' in df_canc_tela:
                df_canc_tela['Crédito'] = df_canc_tela['Crédito Numérico'].apply(formata_contabil)
                df_canc_tela = df_canc_tela.drop(columns=['Crédito Numérico'])
                
            st.dataframe(estilizar_tabela(df_canc_tela), use_container_width=True, hide_index=True)
        else:
            st.success("Nenhuma informação de cancelamento encontrada.")
        st.markdown(rodape_html, unsafe_allow_html=True)

    # --------------------------------------------------
    # MÓDULO PREVISÃO DE RECEITA
    # --------------------------------------------------
    elif escolha == "Motor Comercial":
        st.markdown("<h1>Motor Comercial Projetado</h1>", unsafe_allow_html=True)
        if st.button("Processar Base Comercial"): st.session_state['motor_com_ativo'] = True 
        if st.session_state.get('motor_com_ativo', False):
            with st.spinner("Processando..."):
                df_com = processar_motor_comercial(df_cruzado, df_promo)
                if not df_com.empty:
                    df_com['Data Prevista'] = pd.to_datetime(df_com['Data Prevista']).dt.strftime('%d/%m/%Y')
                    df_com['Data da Venda'] = pd.to_datetime(df_com['Data da Venda']).dt.strftime('%d/%m/%Y')
                    df_tela = df_com.copy()
                    df_tela['Valor Previsto'] = df_tela['Valor Previsto'].apply(formata_contabil)
                    st.dataframe(estilizar_tabela(df_tela.head(500)), use_container_width=True, hide_index=True, height=400)
        st.markdown(rodape_html, unsafe_allow_html=True)

    elif escolha == "Motor de Plataforma":
        st.markdown("<h1>Motor de Plataforma Projetado</h1>", unsafe_allow_html=True)
        if st.button("Processar Base Plataforma"): st.session_state['motor_plat_ativo'] = True 
        if st.session_state.get('motor_plat_ativo', False):
            with st.spinner("Processando..."):
                df_plat = processar_motor_plataforma(df_cruzado, df_parceiros)
                if not df_plat.empty:
                    df_plat['Data Prevista'] = pd.to_datetime(df_plat['Data Prevista']).dt.strftime('%d/%m/%Y')
                    df_plat['Data da Venda'] = pd.to_datetime(df_plat['Data da Venda']).dt.strftime('%d/%m/%Y')
                    df_tela = df_plat.copy()
                    df_tela['Valor Previsto'] = df_tela['Valor Previsto'].apply(formata_contabil)
                    st.dataframe(estilizar_tabela(df_tela.head(500)), use_container_width=True, hide_index=True, height=400)
        st.markdown(rodape_html, unsafe_allow_html=True)

    elif escolha == "Receitas Unificadas":
        st.markdown("<h1>Previsão Unificada (Comercial + Plataforma)</h1>", unsafe_allow_html=True)
        if st.button("Processar Base Unificada"): st.session_state['motor_uni_ativo'] = True 
        if st.session_state.get('motor_uni_ativo', False):
            with st.spinner("Unificando bases..."):
                df_com = processar_motor_comercial(df_cruzado, df_promo)
                df_plat = processar_motor_plataforma(df_cruzado, df_parceiros)
                df_uni = pd.concat([df_com, df_plat]).reset_index(drop=True)
                
                if not df_uni.empty:
                    df_uni['Mês Ref'] = pd.to_datetime(df_uni['Data Prevista']).dt.strftime('%m/%Y')
                    
                    c1, c2, c3 = st.columns(3)
                    f_mes = c1.selectbox("Período (Mês Previsto):", ["Todos"] + sorted(list(df_uni['Mês Ref'].dropna().unique())))
                    f_vendedora = c2.selectbox("Empresa Vendedora:", ["Todas"] + list(df_uni['Empresa Vendedora'].dropna().unique()))
                    f_recebedora = c3.selectbox("Empresa Recebedora:", ["Todas"] + list(df_uni['Recebedora'].dropna().unique()))
                    
                    if f_mes != "Todos": df_uni = df_uni[df_uni['Mês Ref'] == f_mes]
                    if f_vendedora != "Todas": df_uni = df_uni[df_uni['Empresa Vendedora'] == f_vendedora]
                    if f_recebedora != "Todas": df_uni = df_uni[df_uni['Recebedora'] == f_recebedora]
                    
                    df_tela = df_uni.drop(columns=['Mês Ref']).copy()
                    df_tela['Data Prevista'] = pd.to_datetime(df_tela['Data Prevista']).dt.strftime('%d/%m/%Y')
                    df_tela['Data da Venda'] = pd.to_datetime(df_tela['Data da Venda']).dt.strftime('%d/%m/%Y')
                    df_tela['Valor Previsto'] = df_tela['Valor Previsto'].apply(formata_contabil)
                    
                    st.dataframe(estilizar_tabela(df_tela.head(500)), use_container_width=True, hide_index=True)
        st.markdown(rodape_html, unsafe_allow_html=True)

    # --------------------------------------------------
    # MÓDULO RECEBIMENTOS E AUDITORIA
    # --------------------------------------------------
    elif escolha == "Cadastro de Recebimentos":
        st.markdown("<h1>Recebimentos Registrados (Bancorbrás)</h1>", unsafe_allow_html=True)
        if not df_auditoria.empty:
            df_tela = df_auditoria.tail(500).copy()
            
            cols_dinheiro = ['VALOR PAGO', 'VALOR CRÉDITO', 'VALOR ESPERADO', 'DIFERENÇA', 'DIFERENÇA (FALTA)']
            for c in cols_dinheiro:
                if c in df_tela.columns:
                    df_tela[c] = pd.to_numeric(df_tela[c], errors='coerce').fillna(0)
                    df_tela[c] = df_tela[c].apply(formata_contabil)
                    
            st.dataframe(estilizar_tabela(df_tela), use_container_width=True, hide_index=True)
        st.markdown(rodape_html, unsafe_allow_html=True)

    elif escolha == "Auditoria Definitiva":
        st.markdown("<h1>Auditoria: Esperado vs Recebido</h1>", unsafe_allow_html=True)
        
        if st.button("Executar Cruzamento de Auditoria", type="primary"):
            st.session_state['cruzamento_ativo'] = True 
            
        if st.session_state.get('cruzamento_ativo', False):
            with st.spinner("A processar auditoria financeira avançada..."):
                df_exp_com = processar_motor_comercial(df_cruzado, df_promo)
                df_exp_plat = processar_motor_plataforma(df_cruzado, df_parceiros)
                df_esperado = pd.concat([df_exp_com, df_exp_plat]).reset_index(drop=True)
                
                if df_esperado.empty or df_auditoria.empty:
                    st.warning("Faltam dados para auditar.")
                else:
                    df_esperado['Nº Contrato'] = df_esperado['Nº Contrato'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.upper()
                    df_esperado['Empresa'] = df_esperado['Recebedora'].astype(str).str.strip().str.upper()
                    df_esperado['Parcela_Num'] = pd.to_numeric(df_esperado['Parcela'].astype(str).str.replace('ª', ''), errors='coerce')
                    
                    # Prevenção 1: Resgatar Vendedora
                    map_vendedora = df_esperado.set_index('Nº Contrato')['Empresa Vendedora'].to_dict()
                    
                    # Prevenção 2: Resgatar o Cliente a partir da base crua de auditoria
                    df_aud_clean = df_auditoria.dropna(subset=['Nº CONTRATO']).drop_duplicates(subset=['Nº CONTRATO']).copy()
                    df_aud_clean['Nº CONTRATO'] = df_aud_clean['Nº CONTRATO'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.upper()
                    col_cli_aud = next((c for c in df_aud_clean.columns if 'CLIENTE' in c.upper()), None)
                    map_cliente_a = df_aud_clean.set_index('Nº CONTRATO')[col_cli_aud].to_dict() if col_cli_aud else {}

                    map_data_prev = df_esperado.groupby(['Nº Contrato', 'Parcela_Num'])['Data Prevista'].first().to_dict()
                    
                    # Agrupamento inteligente (Agrega o Valor Previsto, mas "segura" o nome do Cliente na memória)
                    df_esperado['Cliente'] = df_esperado.get('Cliente', "Desconhecido")
                    df_esp_grp = df_esperado.groupby(['Nº Contrato', 'Parcela_Num', 'Empresa']).agg({'Valor Previsto': 'sum', 'Cliente': 'first'}).reset_index()
                    
                    df_aud = df_auditoria.copy()
                    df_aud['Nº Contrato'] = df_aud['Nº CONTRATO'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.upper()
                    df_aud['Parcela_Num'] = pd.to_numeric(df_aud['PARCELA ATUAL'], errors='coerce')
                    df_aud['Empresa'] = df_aud['EMPRESA RECEBEDORA'].astype(str).str.strip().str.upper()
                    df_aud_grp = df_aud.groupby(['Nº Contrato', 'Parcela_Num', 'Empresa'])['VALOR PAGO'].sum().reset_index()
                    df_aud_grp.rename(columns={'VALOR PAGO': 'Valor Pago'}, inplace=True)
                    
                    # Cruzamento Otimizado
                    df_final = pd.merge(df_aud_grp, df_esp_grp, on=['Nº Contrato', 'Parcela_Num', 'Empresa'], how='outer').reset_index(drop=True)
                    df_final['Valor Pago'] = df_final['Valor Pago'].fillna(0)
                    df_final['Valor Previsto'] = df_final['Valor Previsto'].fillna(0)
                    df_final['Diferença (R$)'] = df_final['Valor Pago'] - df_final['Valor Previsto']
                    
                    # Injeção de Segurança
                    df_final['Empresa Vendedora'] = df_final['Nº Contrato'].map(map_vendedora).fillna("N/A")
                    df_final['Cliente'] = df_final.apply(
                        lambda x: map_cliente_a.get(x['Nº Contrato'], "Desconhecido") if pd.isna(x.get('Cliente')) else x['Cliente'], axis=1
                    )
                    
                    df_final['Data Prevista'] = df_final.set_index(['Nº Contrato', 'Parcela_Num']).index.map(map_data_prev)
                    df_final['Data Prevista'] = pd.to_datetime(df_final['Data Prevista'], errors='coerce')
                    df_final['Mês Ref'] = df_final['Data Prevista'].dt.strftime('%m/%Y').fillna('Sem Previsão')
                    
                    def classificar_auditoria(row):
                        p = round(row['Valor Pago'], 2)
                        pr = round(row['Valor Previsto'], 2)
                        d = round(row['Diferença (R$)'], 2)
                        
                        if p == pr: return "TUDO CERTO"
                        if p > 0 and pr == 0: return "RECEBEDOR ERRADO"
                        if p == 0 and pr > 0: return "FALTA RECEBER"
                        if d < 0: return "PAGO A MENOR"
                        return "PAGO A MAIOR"
                    
                    df_final['Status'] = df_final.apply(classificar_auditoria, axis=1)
                    
                    st.divider()
                    
                    # --- FILTROS SUPERIORES ---
                    c1, c2, c3 = st.columns(3)
                    
                    vendedoras_lista = sorted([str(v) for v in df_final['Empresa Vendedora'].unique() if pd.notna(v)])
                    recebedoras_lista = sorted([str(r) for r in df_final['Empresa'].unique() if pd.notna(r)])
                    status_lista = sorted([str(s) for s in df_final['Status'].unique() if pd.notna(s)])
                    
                    f_vend = c1.selectbox("Empresa que Vendeu:", ["Todas"] + vendedoras_lista)
                    f_rec = c2.selectbox("Empresa que Recebe:", ["Todas"] + recebedoras_lista)
                    f_status = c3.selectbox("Status da Auditoria:", ["Todos", "Apenas Divergências"] + status_lista)
                    
                    # --- FILTROS INFERIORES ---
                    c4, c5, c6 = st.columns(3) 
                    meses_disponiveis = [m for m in df_final['Mês Ref'].unique() if m != 'Sem Previsão']
                    meses_disponiveis = sorted(meses_disponiveis, key=lambda x: datetime.strptime(x, '%m/%Y'), reverse=True)
                    opcoes_mes = ["Todos", "Sem Previsão"] + meses_disponiveis
                    
                    f_mes = c4.selectbox("Período (Mês Previsto):", opcoes_mes)
                    f_cliente = c5.text_input("Buscar por Cliente (Digite parte do nome):")
                    
                    # Aplicando os filtros
                    if f_vend != "Todas": df_final = df_final[df_final['Empresa Vendedora'] == f_vend]
                    if f_rec != "Todas": df_final = df_final[df_final['Empresa'] == f_rec]
                    
                    if f_status == "Apenas Divergências": 
                        df_final = df_final[df_final['Status'] != "TUDO CERTO"]
                    elif f_status != "Todos": 
                        df_final = df_final[df_final['Status'] == f_status]
                        
                    if f_mes != "Todos": df_final = df_final[df_final['Mês Ref'] == f_mes]
                    if f_cliente: 
                        termo_busca = normalizar_nome(f_cliente)
                        df_final['Cliente_Norm'] = df_final['Cliente'].apply(normalizar_nome)
                        df_final = df_final[df_final['Cliente_Norm'].str.contains(termo_busca, na=False, regex=False)]
                        df_final = df_final.drop(columns=['Cliente_Norm'])
                    
                    # --- RESUMO LG ---
                    df_lg = df_final[df_final['Empresa'].str.contains('LG', na=False) & (df_final['Diferença (R$)'].round(2) != 0)]
                    if not df_lg.empty:
                        st.markdown("<h3>Resumo de Ajustes - Grupo LG</h3>", unsafe_allow_html=True)
                        resumo_lg = df_lg.groupby('Empresa')['Diferença (R$)'].sum().reset_index()
                        resumo_lg.columns = ['Empresa do Grupo LG', 'Ajuste Total Pendente']
                        resumo_lg['Ajuste Total Pendente'] = resumo_lg['Ajuste Total Pendente'].apply(formata_contabil)
                        
                        c_lg1, c_lg2, c_lg3 = st.columns([1, 2, 1])
                        with c_lg2: st.dataframe(estilizar_tabela(resumo_lg, tudo_centro=True), use_container_width=True, hide_index=True)
                    
                    # --- TABELA ANALÍTICA DE CRUZAMENTO ---
                    st.markdown("<h3>Tabela Analítica de Cruzamento</h3>", unsafe_allow_html=True)
                    df_tela = df_final[['Nº Contrato', 'Cliente', 'Parcela_Num', 'Mês Ref', 'Empresa Vendedora', 'Empresa', 'Valor Previsto', 'Valor Pago', 'Diferença (R$)', 'Status']].copy()
                    df_tela.columns = ['Nº Contrato', 'Cliente', 'Parcela', 'Mês Previsão', 'Vendedora Originária', 'Recebedora', 'Valor Previsto', 'Valor Pago', 'Diferença', 'Status']
                    
                    df_tela['Valor Previsto'] = df_tela['Valor Previsto'].apply(formata_contabil)
                    df_tela['Valor Pago'] = df_tela['Valor Pago'].apply(formata_contabil)
                    df_tela['Diferença'] = df_tela['Diferença'].apply(formata_contabil)
                    
                    def pintar_status(linha):
                        s = linha['Status']
                        estilo_base = 'text-align: center !important; '
                        if 'TUDO CERTO' in s: return [estilo_base + 'color: #155724;'] * len(linha)
                        if 'ERRADO' in s or 'MENOR' in s: return [estilo_base + 'background-color: #F8D7DA; color: #721C24; font-weight: bold;'] * len(linha)
                        if 'FALTA' in s: return [estilo_base + 'background-color: #FFF3CD; color: #856404; font-weight: bold;'] * len(linha)
                        return [estilo_base + 'color: #856404;'] * len(linha)
                    
                    st.dataframe(estilizar_tabela(df_tela.head(500), func_pintar=pintar_status, tudo_centro=True), use_container_width=True, hide_index=True, height=500)
        st.markdown(rodape_html, unsafe_allow_html=True)