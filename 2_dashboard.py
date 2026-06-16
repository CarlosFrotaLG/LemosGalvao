import streamlit as st
import pandas as pd
import sqlite3
import numpy as np
from datetime import datetime
import unicodedata
import io # 🚀 NOVA BIBLIOTECA PARA GERAR O EXCEL NA HORA

# ==========================================
# ⚙️ CONFIGURAÇÃO E FUNÇÕES DE FORMATAÇÃO
# ==========================================
st.set_page_config(page_title="Sistema Financeiro Mestre", layout="wide")
st.title("📊 Painel Financeiro Mestre")

def formata_moeda(valor):
    try:
        if pd.isna(valor) or valor == '': return "-"
        return f"R$ {float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except: return valor

def normalizar_nome(nome):
    if pd.isna(nome) or nome == '': return ''
    nfd = unicodedata.normalize('NFD', str(nome))
    return nfd.encode('ascii', 'ignore').decode('utf8').upper().strip()

# 🚀 FUNÇÃO MÁGICA PARA GERAR EXCEL
def converter_df_para_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Relatório Financeiro')
    processed_data = output.getvalue()
    return processed_data

# ==========================================
# 📥 CARREGAMENTO DO BANCO DE DADOS
# ==========================================
@st.cache_data
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
    "PLANEJAMENTO": "LG PLANEJAMENTO",
    "SOLUÇÕES": "LG SOLUÇÕES",
    "SOLUCOES": "LG SOLUÇÕES",
    "INTELIGÊNCIA": "LG INTELIGÊNCIA",
    "INTELIGENCIA": "LG INTELIGÊNCIA",
    "BELO HORIZONTE": "LG BELO HORIZONTE",
    "CURITIBA": "LG CURITIBA",
    "PLATAFORMA 1": "LG PLATAFORMA 1",
    "PLATAFORMA 2": "LG PLATAFORMA 2",
    "PLATAFORMA 3": "LG PLATAFORMA 3"
}

consultores_oficiais = [
    "ANDERSON", "ANDERSON / DENISE", "AURORA", "BRUNO BORGES", "CARLOS",
    "CÍCERO", "DENISE", "DOUGLAS", "EDUARDO HENRIQUE", "JOSÉ WAGNER",
    "JULIO AVIDUS", "LEANDRO LEMOS", "LEANDRO MÁXIMO", "LUANA", 
    "LUZ DA SERRA", "MÁRCIO PASSOS", "MASTROCOLA", "PEDRO GONTIJO", 
    "RODOLPHO", "RUY", "TIAGO MENDO"
]

vip_list = [normalizar_nome(c) for c in consultores_oficiais]

if not df_auditoria.empty and 'EMPRESA RECEBEDORA' in df_auditoria.columns:
    df_auditoria['EMPRESA RECEBEDORA'] = df_auditoria['EMPRESA RECEBEDORA'].astype(str).str.strip().str.upper()
    df_auditoria['EMPRESA RECEBEDORA'] = df_auditoria['EMPRESA RECEBEDORA'].replace(map_empresas)

if not df_vendas.empty and 'Empresa Vendedora' in df_vendas.columns:
    df_vendas['Empresa Vendedora'] = df_vendas['Empresa Vendedora'].astype(str).str.strip().str.upper()
    df_vendas['Empresa Vendedora'] = df_vendas['Empresa Vendedora'].replace(map_empresas)

# ==========================================
# 🧠 MOTORES COM CACHE (MÁXIMA VELOCIDADE)
# ==========================================
@st.cache_data
def processar_motor_comercial(df_cruzado, df_promo):
    registros = []
    if df_cruzado.empty: return pd.DataFrame()
    for index, row in df_cruzado.iterrows():
        contrato = row['Nº Contrato']
        data_venda = row['Data da Venda']
        vendedora = str(row['Empresa Vendedora']).strip()
        try: credito = float(row['Crédito Numérico'])
        except: credito = 0.0
        if pd.isnull(data_venda) or credito == 0: continue
            
        is_promo = False
        if not df_promo.empty:
            if 'Nº Contrato' in df_promo.columns:
                match = df_promo[df_promo['Nº Contrato'].astype(str) == str(contrato)]
                if not match.empty:
                    st_p = str(match.iloc[0].get('Status Promocional', '')).lower()
                    if 'promo' in st_p: is_promo = True
            if not is_promo and 'Grupo' in df_promo.columns and 'Grupo' in row:
                if str(row['Grupo']) in df_promo['Grupo'].astype(str).values: is_promo = True

        qtd_parcelas = 24 if is_promo else 10
        for i in range(1, qtd_parcelas + 1):
            data_prevista = data_venda + pd.DateOffset(months=i)
            if is_promo: taxa = 0.0015 
            else:
                v_bruto = row.get(f'P{i}', 0)
                try: taxa = float(str(v_bruto).replace('%', '').replace(',', '.'))
                except: taxa = 0.0
                if pd.isna(taxa): taxa = 0.0
            val = credito * taxa
            if val > 0:
                registros.append({
                    'Nº Contrato': contrato, 'Data da Venda': data_venda.strftime('%d/%m/%Y'),
                    'Parcela': f"{i}ª", 'Data Prevista': data_prevista.strftime('%d/%m/%Y'),
                    'Data Calc': data_prevista, 
                    'Recebedora': vendedora, 'Valor Previsto': val, 'Tipo': 'Comercial'
                })
    return pd.DataFrame(registros)

@st.cache_data
def processar_motor_plataforma(df_cruzado, df_parceiros):
    registros = []
    if df_cruzado.empty: return pd.DataFrame()
    dict_p = {}
    if not df_parceiros.empty:
        df_p = df_parceiros.copy()
        df_p['DATA INÍCIO'] = pd.to_datetime(df_p['DATA INÍCIO'], errors='coerce')
        df_p['DATA FIM'] = pd.to_datetime(df_p['DATA FIM'], errors='coerce')
        for _, rp in df_p.iterrows():
            emp = str(rp.get('EMPRESA PARCEIRA', '')).strip().upper()
            emp = map_empresas.get(emp, emp)
            if emp not in dict_p: dict_p[emp] = []
            mae_traduzida = str(rp.get('QUEM INDICOU (MÃE)', '-')).strip().upper()
            mae_traduzida = map_empresas.get(mae_traduzida, mae_traduzida)
            dict_p[emp].append((rp['DATA INÍCIO'], rp['DATA FIM'], str(rp.get('NÍVEL DE PLATAFORMA', '')).strip(), mae_traduzida))

    def buscar_mae(empresa, data_v):
        if pd.isnull(data_v) or empresa not in dict_p: return 'Sem Plataforma', '-'
        for di, dfim, niv, mae in dict_p[empresa]:
            if pd.notnull(di) and pd.notnull(dfim) and di <= data_v <= dfim: return niv, mae
        return 'Sem Plataforma', '-'

    for index, row in df_cruzado.iterrows():
        contrato = row['Nº Contrato']
        data_venda = row['Data da Venda']
        vendedora = str(row['Empresa Vendedora']).strip()
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
                    'Nº Contrato': contrato, 'Data da Venda': data_venda.strftime('%d/%m/%Y'),
                    'Parcela': f"{p}ª", 'Data Prevista': dt_prev.strftime('%d/%m/%Y'),
                    'Data Calc': dt_prev, 
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
    return pd.DataFrame(registros)

df_motor = df_vendas.copy()
if not df_motor.empty and not df_comissao.empty:
    df_motor['Crédito Numérico'] = pd.to_numeric(df_motor['Crédito'], errors='coerce').fillna(0)
    df_motor['Data da Venda'] = pd.to_datetime(df_motor['Data da Venda'], errors='coerce')
    df_cruzado = pd.merge(df_motor, df_comissao, left_on='CHAVE (OCULTA)', right_on='CHAVE COMBINADA', how='left')
else:
    df_cruzado = pd.DataFrame()

# ==========================================
# 🌟 AS 5 ABAS DO SISTEMA
# ==========================================
aba1, aba2, aba3, aba4, aba5 = st.tabs([
    "📊 Dashboard Executivo (BI)", "📄 Auditoria (Bruta)", "💰 Motor Comercial", "🏢 Motor Plataforma", "🔍 ELO FINAL"
])

# ------------------------------------------
# ABA 1: DASHBOARD EXECUTIVO
# ------------------------------------------
with aba1:
    st.header("📊 Visão de Negócio & BI")
    if not df_vendas.empty:
        df_dash = df_motor.copy()
        df_dash['Mês/Ano Venda'] = df_dash['Data da Venda'].dt.to_period('M').astype(str)
        
        st.markdown("### 🔎 Filtros Globais")
        colF1, colF2, colF3 = st.columns(3)
        lista_empresas = ["Todas as Empresas"] + list(df_dash['Empresa Vendedora'].dropna().unique())
        f_empresa = colF1.selectbox("Filtrar por Empresa:", lista_empresas)
        
        if not df_dash['Data da Venda'].dropna().empty:
            min_date = df_dash['Data da Venda'].dropna().min().date()
            max_date = df_dash['Data da Venda'].dropna().max().date()
        else:
            min_date = datetime.today().date()
            max_date = datetime.today().date()
            
        f_data_inicio = colF2.date_input("Período Inicial:", min_date)
        f_data_fim = colF3.date_input("Período Final:", max_date)
        
        if f_empresa != "Todas as Empresas":
            df_dash = df_dash[df_dash['Empresa Vendedora'] == f_empresa]
        df_dash = df_dash[(df_dash['Data da Venda'].dt.date >= f_data_inicio) & (df_dash['Data da Venda'].dt.date <= f_data_fim)]

        st.divider()

        st.subheader("💰 1. Previsão de Caixa e Receitas (Mês Atual)")
        mes_atual = pd.Timestamp.today().to_period('M')
        df_prev_com = processar_motor_comercial(df_cruzado, df_promo)
        df_prev_plat = processar_motor_plataforma(df_cruzado, df_parceiros)
        df_prev_total = pd.concat([df_prev_com, df_prev_plat])
        
        valor_previsto_mes = 0.0
        valor_recebido_mes = 0.0
        
        if not df_prev_total.empty:
            df_prev_total['Mês Calc'] = df_prev_total['Data Calc'].dt.to_period('M')
            df_prev_mes = df_prev_total[df_prev_total['Mês Calc'] == mes_atual]
            if f_empresa != "Todas as Empresas":
                df_prev_mes = df_prev_mes[df_prev_mes['Recebedora'].str.upper() == f_empresa]
            valor_previsto_mes = df_prev_mes['Valor Previsto'].sum()

        if not df_auditoria.empty:
            df_aud = df_auditoria.copy()
            df_aud['VALOR PAGO'] = pd.to_numeric(df_aud['VALOR PAGO'], errors='coerce').fillna(0)
            if 'DATA PAGAMENTO' in df_aud.columns: 
                df_aud['Mês Pag'] = pd.to_datetime(df_aud['DATA PAGAMENTO'], errors='coerce').dt.to_period('M')
                df_aud_mes = df_aud[df_aud['Mês Pag'] == mes_atual]
            else: 
                df_aud_mes = df_aud
            if f_empresa != "Todas as Empresas":
                df_aud_mes = df_aud_mes[df_aud_mes['EMPRESA RECEBEDORA'].str.upper() == f_empresa]
            valor_recebido_mes = df_aud_mes['VALOR PAGO'].sum()

        m1, m2, m3 = st.columns(3)
        m1.metric(f"Receita Prevista ({mes_atual})", formata_moeda(valor_previsto_mes))
        m2.metric(f"Receita Confirmada/Paga ({mes_atual})", formata_moeda(valor_recebido_mes))
        m3.metric("Diferença (Falta Pagar)", formata_moeda(valor_recebido_mes - valor_previsto_mes))
        
        st.divider()

        col1, col2 = st.columns([2, 1])
        with col1:
            st.subheader("📈 2. Produção Mês a Mês")
            if not df_dash.empty:
                tabela_producao = pd.crosstab(index=df_dash['Empresa Vendedora'], columns=df_dash['Mês/Ano Venda'], values=df_dash['Crédito Numérico'], aggfunc='sum').fillna(0)
                tabela_producao['TOTAL ACUMULADO'] = tabela_producao.sum(axis=1)
                tabela_producao = tabela_producao.sort_values(by='TOTAL ACUMULADO', ascending=False)
                tabela_producao_formatada = tabela_producao.map(formata_moeda)
                st.dataframe(tabela_producao_formatada, use_container_width=True)

        with col2:
            st.subheader("⚠️ 3. Cancelamentos Registados")
            col_status = next((col for col in ['Status', 'Situação', 'SITUACAO', 'STATUS'] if col in df_dash.columns), None)
            if col_status:
                df_cancelados = df_dash[df_dash[col_status].astype(str).str.contains('Cancel|Distrat', case=False, na=False)]
                if not df_cancelados.empty:
                    resumo_cancelados = df_cancelados.groupby('Empresa Vendedora')['Crédito Numérico'].agg(['count', 'sum']).reset_index()
                    resumo_cancelados.columns = ['Empresa', 'Qtd Cancelados', 'Volume Perdido']
                    resumo_cancelados['Volume Perdido'] = resumo_cancelados['Volume Perdido'].apply(formata_moeda)
                    st.dataframe(resumo_cancelados, use_container_width=True, hide_index=True)
                else: st.success("Nenhum cancelamento no período!")
            else: st.info("Coluna de 'Status' não encontrada.")

        st.divider()

        st.subheader("🏆 4. Rankings de Produção")
        colA, colB = st.columns(2)
        
        with colA:
            st.markdown("#### 🏢 Top Empresas")
            if not df_dash.empty:
                rank_emp = df_dash.groupby('Empresa Vendedora')['Crédito Numérico'].sum().reset_index()
                rank_emp.columns = ['Empresa', 'Volume Produzido']
                rank_emp = rank_emp.sort_values(by='Volume Produzido', ascending=False).reset_index(drop=True)
                rank_emp.index = rank_emp.index + 1 
                rank_emp_format = rank_emp.copy()
                rank_emp_format['Volume Produzido'] = rank_emp_format['Volume Produzido'].apply(formata_moeda)
                st.dataframe(rank_emp_format, use_container_width=True)

        with colB:
            col_consultor = next((col for col in df_dash.columns if 'CONSULT' in col.upper()), None)
            if not col_consultor:
                col_consultor = next((col for col in df_dash.columns if ('VENDED' in col.upper() or 'CORRET' in col.upper()) and 'EMPRESA' not in col.upper() and 'NÍVEL' not in col.upper() and 'NIVEL' not in col.upper()), None)
            
            st.markdown(f"#### 👤 Top Consultores VIP")
            if col_consultor and not df_dash.empty:
                df_cons = df_dash.copy()
                df_cons['NOME_FILTRADO'] = df_cons[col_consultor].apply(normalizar_nome)
                df_cons = df_cons[df_cons['NOME_FILTRADO'].isin(vip_list)]
                
                if not df_cons.empty:
                    def restaurar_nome(nome_feio): return next((nome_vip for nome_vip in consultores_oficiais if normalizar_nome(nome_vip) == nome_feio), nome_feio)
                    df_cons['Consultor Oficial'] = df_cons['NOME_FILTRADO'].apply(restaurar_nome)
                    rank_cons = df_cons.groupby('Consultor Oficial')['Crédito Numérico'].sum().reset_index()
                    rank_cons.columns = ['Consultor', 'Volume Produzido']
                    rank_cons = rank_cons.sort_values(by='Volume Produzido', ascending=False).reset_index(drop=True)
                    rank_cons.index = rank_cons.index + 1 
                    rank_formatado = rank_cons.copy()
                    rank_formatado['Volume Produzido'] = rank_formatado['Volume Produzido'].apply(formata_moeda)
                    st.dataframe(rank_formatado, use_container_width=True)
                else: st.info("Sem vendas vinculadas aos VIPs.")
    else: st.warning("Sem dados.")

# ------------------------------------------
# ABA 2 A 5
# ------------------------------------------
with aba2:
    st.subheader("Auditoria Bruta")
    if not df_auditoria.empty: st.dataframe(df_auditoria, use_container_width=True, height=400)

with aba3:
    st.subheader("Fluxo Comercial Projetado")
    if st.button("🚀 PROCESSAR COMERCIAL", key="btn_com"): st.session_state['motor_com_ativo'] = True 
    if st.session_state.get('motor_com_ativo', False):
        with st.spinner("Processando..."):
            df_com = processar_motor_comercial(df_cruzado, df_promo)
            if not df_com.empty:
                df_com_visual = df_com.drop(columns=['Data Calc']) 
                df_com_visual['Valor Previsto'] = df_com_visual['Valor Previsto'].apply(formata_moeda)
                st.dataframe(df_com_visual, use_container_width=True, height=500)
                
                # 🚀 BOTÃO DE DOWNLOAD EXCEL (COMERCIAL)
                excel_comercial = converter_df_para_excel(df_com_visual)
                st.download_button(label="📥 Exportar Comercial para Excel", data=excel_comercial, file_name="Previsao_Comercial.xlsx", mime="application/vnd.ms-excel", type="primary")

with aba4:
    st.subheader("Fluxo Plataforma Projetado")
    if st.button("🚀 PROCESSAR PLATAFORMA", key="btn_plat"): st.session_state['motor_plat_ativo'] = True 
    if st.session_state.get('motor_plat_ativo', False):
        with st.spinner("Processando..."):
            df_plat = processar_motor_plataforma(df_cruzado, df_parceiros)
            if not df_plat.empty:
                df_plat_visual = df_plat.drop(columns=['Data Calc'])
                df_plat_visual['Valor Previsto'] = df_plat_visual['Valor Previsto'].apply(formata_moeda)
                st.dataframe(df_plat_visual, use_container_width=True, height=500)
                
                # 🚀 BOTÃO DE DOWNLOAD EXCEL (PLATAFORMA)
                excel_plataforma = converter_df_para_excel(df_plat_visual)
                st.download_button(label="📥 Exportar Plataforma para Excel", data=excel_plataforma, file_name="Previsao_Plataforma.xlsx", mime="application/vnd.ms-excel", type="primary")

with aba5:
    st.subheader("🔍 Auditoria Definitiva (Bancorbrás vs. O Nosso Motor)")
    mostrar = st.radio("Filtro de Resultados:", ["🚨 Mostrar Apenas Divergências", "✅ Mostrar Toda a Auditoria"], horizontal=True, key="filtro_auditoria")
    
    if st.button("⚡ EXECUTAR CRUZAMENTO FINAL", use_container_width=True, type="primary"):
        st.session_state['cruzamento_ativo'] = True 
        
    if st.session_state.get('cruzamento_ativo', False):
        with st.spinner("A fundir o Motor Matemático com a Realidade..."):
            df_exp_com = processar_motor_comercial(df_cruzado, df_promo)
            df_exp_plat = processar_motor_plataforma(df_cruzado, df_parceiros)
            df_esperado = pd.concat([df_exp_com, df_exp_plat])
            
            if df_esperado.empty or df_auditoria.empty:
                st.warning("Faltam dados.")
            else:
                df_esperado['Nº Contrato'] = df_esperado['Nº Contrato'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                df_esperado['Empresa'] = df_esperado['Recebedora'].astype(str).str.strip().str.upper()
                df_esperado['Parcela_Num'] = pd.to_numeric(df_esperado['Parcela'].astype(str).str.replace('ª', ''), errors='coerce')
                df_esp_grp = df_esperado.groupby(['Nº Contrato', 'Parcela_Num', 'Empresa'])['Valor Previsto'].sum().reset_index()
                
                df_aud = df_auditoria.copy()
                df_aud['Nº Contrato'] = df_aud['Nº CONTRATO'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                df_aud['Parcela_Num'] = pd.to_numeric(df_aud['PARCELA ATUAL'], errors='coerce')
                df_aud['Empresa'] = df_aud['EMPRESA RECEBEDORA'].astype(str).str.strip().str.upper()
                df_aud_grp = df_aud.groupby(['Nº Contrato', 'Parcela_Num', 'Empresa'])['VALOR PAGO'].sum().reset_index()
                df_aud_grp.rename(columns={'VALOR PAGO': 'Valor Pago'}, inplace=True)
                
                chaves_auditadas = df_aud_grp[['Nº Contrato', 'Parcela_Num']].drop_duplicates()
                df_esp_filtrado = pd.merge(chaves_auditadas, df_esp_grp, on=['Nº Contrato', 'Parcela_Num'], how='left')
                df_final = pd.merge(df_aud_grp, df_esp_filtrado, on=['Nº Contrato', 'Parcela_Num', 'Empresa'], how='outer')
                
                df_final['Valor Pago'] = df_final['Valor Pago'].fillna(0)
                df_final['Valor Previsto'] = df_final['Valor Previsto'].fillna(0)
                df_final['Diferença (R$)'] = df_final['Valor Pago'] - df_final['Valor Previsto']
                
                def classificar_auditoria(row):
                    p = round(row['Valor Pago'], 2)
                    pr = round(row['Valor Previsto'], 2)
                    d = round(row['Diferença (R$)'], 2)
                    if p == pr: return "✅ TUDO CERTO"
                    if p > 0 and pr == 0: return "🚨 RECEBEDOR ERRADO (Ou Maior)"
                    if p == 0 and pr > 0: return "❌ FALTA RECEBER"
                    if d < 0: return "⚠️ PAGO A MENOR"
                    return "⚠️ PAGO A MAIOR"
                
                df_final['Status do Repasse'] = df_final.apply(classificar_auditoria, axis=1)
                df_final = df_final.rename(columns={'Parcela_Num': 'Parcela'}).sort_values(by=['Nº Contrato', 'Parcela'])
                
                if "Divergências" in mostrar: df_final = df_final[df_final['Status do Repasse'] != "✅ TUDO CERTO"]
                
                def pintar_status(linha):
                    s = linha['Status do Repasse']
                    if 'TUDO CERTO' in s: return ['background-color: #D4EDDA; color: #155724;'] * len(linha)
                    if 'ERRADO' in s: return ['background-color: #F8D7DA; color: #721C24; font-weight: bold;'] * len(linha)
                    if 'FALTA' in s: return ['background-color: #FFF3CD; color: #856404; font-weight: bold;'] * len(linha)
                    return ['background-color: #FFEeba; color: #856404;'] * len(linha)
                
                # Salvamos uma cópia limpa do DataFrame para o Excel (antes de virar texto de moeda)
                df_final_excel = df_final.copy()
                
                df_final['Valor Pago'] = df_final['Valor Pago'].apply(formata_moeda)
                df_final['Valor Previsto'] = df_final['Valor Previsto'].apply(formata_moeda)
                df_final['Diferença (R$)'] = df_final['Diferença (R$)'].apply(formata_moeda)
                
                st.success(f"🎯 Cruzamento concluído! O Motor auditou {len(df_final)} registos.")
                st.dataframe(df_final.style.apply(pintar_status, axis=1), use_container_width=True, height=600)
                
                # 🚀 BOTÃO DE DOWNLOAD EXCEL (AUDITORIA / ELO FINAL)
                excel_auditoria = converter_df_para_excel(df_final_excel)
                st.download_button(label="📥 Exportar Relatório de Auditoria para Excel", data=excel_auditoria, file_name="Auditoria_Bancorbras.xlsx", mime="application/vnd.ms-excel", type="primary")