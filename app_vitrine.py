# -*- coding: utf-8 -*-
# ═══════════════════════════════════════════════════════════════════
#  LEMOS GALVÃO · PORTAL DE VISUALIZAÇÃO (STREAMLIT CLOUD)
#  Lê as planilhas do GitHub · Sem dados sensíveis · Só dashboards
#  Deploy: lemosgalvaofinanceiro.streamlit.app
# ═══════════════════════════════════════════════════════════════════
import streamlit as st
import pandas as pd
import io, os
from datetime import datetime, date
from motor import (
    fmt_brl, fmt_m, sf,
    preparar_vendas, preparar_comissionamento,
    preparar_promocoes, preparar_parceiros,
    motor_comercial, motor_plataforma,
    split_consultores
)

st.set_page_config(
    page_title="LG · Portal Financeiro",
    page_icon="🏛️", layout="wide",
    initial_sidebar_state="expanded"
)

# Credenciais (portal de visualização — sem dados sensíveis de clientes)
USUARIOS = {
    "Carlos.Frota":     "Lg#2026",
    "Valeira.Oliveira": "Lg@2026",
    "Leandro.Lemos":    "Lg$2026",
    "Klaus.Meirose":    "Lg&2026",
}

# ── CSS (idêntico ao local) ────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;600;700&family=Inter:wght@300;400;500;600&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif!important}
[data-testid="stSidebar"]{background:#0F1C2E!important;border-right:1px solid rgba(201,168,76,.2)}
[data-testid="stSidebar"] *{color:rgba(255,255,255,.75)!important}
[data-testid="stSidebar"] .stButton>button{
    background:transparent!important;border:none!important;color:rgba(255,255,255,.7)!important;
    text-align:left!important;padding:8px 16px!important;width:100%!important;
    font-size:13px!important;border-radius:6px!important;transition:all .15s!important}
[data-testid="stSidebar"] .stButton>button:hover{background:rgba(201,168,76,.15)!important;color:#E2C97E!important}
[data-testid="stSidebar"] hr{border-color:rgba(201,168,76,.2)!important}
.main{background:#F7F4EF!important}
.block-container{padding:1.5rem 2rem!important}
h1{font-family:'Cormorant Garamond',serif!important;color:#0F1C2E!important;font-size:1.9rem!important;font-weight:700!important}
h2{font-family:'Cormorant Garamond',serif!important;color:#0F1C2E!important;font-size:1.4rem!important}
h3{font-family:'Cormorant Garamond',serif!important;color:#C9A84C!important;font-size:1.1rem!important}
[data-testid="stMetricValue"]{color:#0F1C2E!important;font-family:'Cormorant Garamond',serif!important;font-size:1.7rem!important;font-weight:700!important}
[data-testid="stMetricLabel"]{color:#5A5A6E!important;font-size:.72rem!important;text-transform:uppercase;letter-spacing:.06em}
[data-testid="metric-container"]{background:white!important;border:1px solid #DDD9D0!important;border-radius:10px!important;padding:16px 20px!important;border-left:3px solid #C9A84C!important}
.stButton>button[kind="primary"]{background:#C9A84C!important;color:#0F1C2E!important;border:none!important;font-weight:600!important;border-radius:6px!important}
.stTabs [data-baseweb="tab-list"]{border-bottom:2px solid #C9A84C!important}
.stTabs [aria-selected="true"]{color:#0F1C2E!important;font-weight:600!important}
.rodape{border-top:2px solid #C9A84C;margin-top:48px;padding-top:14px;color:#9A9AAA;font-size:11px;display:flex;justify-content:space-between}
.aviso-nuvem{background:#E8F5EE;border:1px solid rgba(26,107,60,.2);border-radius:8px;padding:10px 14px;font-size:12px;color:#1A6B3C;margin-bottom:16px}
</style>
""", unsafe_allow_html=True)

RODAPE = """<div class="rodape">
<span>Lemos Galvão · Portal de Visualização · 2026</span>
<span>Capital Financial Center · SIG Quadra 4 · Brasília-DF</span>
</div>"""

# ── LOGIN ──────────────────────────────────────────────────────────
if 'auth' not in st.session_state: st.session_state['auth'] = False
if 'user' not in st.session_state: st.session_state['user'] = ''

if not st.session_state['auth']:
    st.markdown("<br><br>", unsafe_allow_html=True)
    _, col, _ = st.columns([1.5,1,1.5])
    with col:
        st.markdown("""
        <div style="text-align:center;margin-bottom:20px">
            <div style="font-family:'Cormorant Garamond',serif;font-size:28px;font-weight:700;color:#0F1C2E">
                Lemos Galvão
            </div>
            <div style="font-size:11px;color:#C9A84C;letter-spacing:.1em;text-transform:uppercase">
                Portal Financeiro
            </div>
        </div>""", unsafe_allow_html=True)
        usr = st.text_input("Usuário")
        pwd = st.text_input("Senha", type="password")
        if st.button("Entrar", type="primary", use_container_width=True):
            if usr in USUARIOS and USUARIOS[usr] == pwd:
                st.session_state['auth'] = True
                st.session_state['user'] = usr
                st.rerun()
            else:
                st.error("Credenciais inválidas.")
    st.stop()

# ── CARREGAR PLANILHAS DO REPOSITÓRIO ─────────────────────────────
# No Streamlit Cloud, os arquivos ficam no mesmo repositório do GitHub
# Coloque as planilhas Tb_Vendas.xlsx, Tb_Parceiros.xlsx,
# Tb_Comissionamento.xlsx e Tb_Auditoria.xlsx na raiz do repositório

BASE = os.path.dirname(os.path.abspath(__file__))

@st.cache_data(ttl=600, show_spinner=False)
def carregar_planilhas():
    """Carrega as planilhas da raiz do repositório"""
    dados = {}
    arquivos = {
        'vendas':         'Tb_Vendas.xlsx',
        'parceiros':      'Tb_Parceiros.xlsx',
        'comissionamento':'Tb_Comissionamento.xlsx',
        'auditoria':      'Tb_Auditoria.xlsx',
    }
    for chave, nome in arquivos.items():
        caminho = os.path.join(BASE, nome)
        if os.path.exists(caminho):
            try:
                skip = 4 if 'auditoria' in chave.lower() else 0
                df = pd.read_excel(caminho, skiprows=skip)
                df.columns = [str(c).strip() for c in df.columns]
                dados[chave] = df
            except Exception as e:
                st.warning(f"Erro ao ler {nome}: {e}")
                dados[chave] = pd.DataFrame()
        else:
            dados[chave] = pd.DataFrame()
    return dados

@st.cache_data(ttl=600, show_spinner=False)
def calcular_tudo(n_v, n_c, n_p):
    """Cache dos motores"""
    dados = carregar_planilhas()
    df_v  = preparar_vendas(dados['vendas']) if not dados['vendas'].empty else pd.DataFrame()
    idx_c = preparar_comissionamento(dados['comissionamento']) if not dados['comissionamento'].empty else {}
    gp    = preparar_promocoes()
    lp    = preparar_parceiros(dados['parceiros']) if not dados['parceiros'].empty else {}
    df_mc, _ = motor_comercial(df_v, idx_c, gp) if not df_v.empty and idx_c else (pd.DataFrame(), set())
    df_mp    = motor_plataforma(df_v, lp) if not df_v.empty and lp else pd.DataFrame()
    df_all   = pd.concat([df_mc, df_mp], ignore_index=True)
    return df_v, df_mc, df_mp, df_all

# ── SIDEBAR ────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div style="padding:16px 16px 12px;border-bottom:1px solid rgba(201,168,76,.2)">
        <div style="font-family:'Cormorant Garamond',serif;font-size:18px;font-weight:700;color:#E2C97E">Lemos Galvão</div>
        <div style="font-size:9px;color:rgba(255,255,255,.3);text-transform:uppercase;letter-spacing:.1em">Portal de Visualização</div>
        <div style="font-size:11px;color:rgba(255,255,255,.4);margin-top:5px">Olá, {st.session_state['user'].split('.')[0]}</div>
    </div>""", unsafe_allow_html=True)

    if 'pag' not in st.session_state: st.session_state['pag'] = 'dashboard'
    def nav(p): st.session_state['pag'] = p
    def sb(label, pag):
        if st.button(label, use_container_width=True): nav(pag)

    st.markdown("<div style='padding:8px 8px 3px;font-size:9px;color:rgba(255,255,255,.3);text-transform:uppercase;letter-spacing:.1em'>Dashboards</div>", unsafe_allow_html=True)
    sb("📊  Dashboard",        'dashboard')
    sb("📈  Produção",         'producao')
    sb("🏆  Rankings",         'rankings')

    st.markdown("<div style='padding:10px 8px 3px;font-size:9px;color:rgba(255,255,255,.3);text-transform:uppercase;letter-spacing:.1em'>Receitas</div>", unsafe_allow_html=True)
    sb("⚙️  Motor Comercial",  'motor_com')
    sb("🔗  Motor Plataforma", 'motor_plat')
    sb("📦  Receitas Unif.",   'motor_all')

    st.divider()
    st.markdown("""<div style='padding:10px 14px;font-size:10px;color:rgba(255,255,255,.3);line-height:1.5'>
    🔒 Portal somente leitura<br>
    Para lançamentos, use o sistema local
    </div>""", unsafe_allow_html=True)
    if st.button("🔄 Recarregar dados", use_container_width=True):
        st.cache_data.clear(); st.rerun()
    if st.button("🚪 Sair", use_container_width=True):
        st.session_state['auth'] = False; st.rerun()
    st.markdown(f"<div style='padding:8px 16px;font-size:9px;color:rgba(255,255,255,.2);text-align:center'>v10.1 Cloud · {date.today().strftime('%d/%m/%Y')}</div>", unsafe_allow_html=True)

# ── DADOS ──────────────────────────────────────────────────────────
with st.spinner("Carregando dados..."):
    dados_raw = carregar_planilhas()
    df_v_raw  = dados_raw.get('vendas', pd.DataFrame())
    df_a      = dados_raw.get('auditoria', pd.DataFrame())

    n_v = len(df_v_raw); n_c = len(dados_raw.get('comissionamento',pd.DataFrame()))
    n_p = len(dados_raw.get('parceiros',pd.DataFrame()))
    df_v, df_mc, df_mp, df_mall = calcular_tudo(n_v, n_c, n_p)

import plotly.graph_objects as go
import plotly.express as px

pag = st.session_state.get('pag','dashboard')

st.markdown("""<div class="aviso-nuvem">
🌐 Portal de visualização · Dados atualizados quando as planilhas são enviadas ao GitHub ·
Para lançamentos e auditoria, use o sistema local
</div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# DASHBOARD
# ══════════════════════════════════════════════════════════════════
if pag == 'dashboard':
    st.markdown("## Dashboard · Lemos Galvão")
    st.markdown("<div style='color:#9A9AAA;font-size:13px;margin:-8px 0 20px'>Visão geral da operação · Bancorbrás</div>", unsafe_allow_html=True)

    ativos = df_v[df_v['Ativo']] if not df_v.empty else pd.DataFrame()
    total_cred = ativos['Crédito_num'].sum() if not ativos.empty else 0
    total_con  = len(ativos)
    cancelados = len(df_v) - total_con if not df_v.empty else 0
    tv_prev = df_mall['Valor Previsto'].sum() if not df_mall.empty else 0
    tv_pago = sf(df_a['VALOR PAGO'].sum()) if not df_a.empty and 'VALOR PAGO' in df_a.columns else 0

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Carteira Ativa",    fmt_m(total_cred),  f"{total_con:,} contratos · {cancelados} cancelados")
    c2.metric("Comissão Prevista", fmt_m(tv_prev),     f"{len(df_mall):,} parcelas")
    c3.metric("Total Recebido",    fmt_m(tv_pago),     f"Saldo: {fmt_m(tv_prev - tv_pago)}")
    c4.metric("Ticket Médio",      fmt_m(total_cred/total_con) if total_con else "—", "por contrato ativo")

    st.markdown("<br>", unsafe_allow_html=True)
    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown("### Comercial vs Plataforma")
        if not df_mall.empty:
            pt = df_mall.groupby('Tipo')['Valor Previsto'].sum().reset_index()
            fig = go.Figure(go.Pie(
                labels=pt['Tipo'], values=pt['Valor Previsto'],
                hole=.65, marker_colors=['#0F1C2E','#C9A84C'], textinfo='label+percent',
            ))
            fig.update_layout(margin=dict(t=10,b=0,l=0,r=0), height=260,
                paper_bgcolor='rgba(0,0,0,0)', showlegend=False,
                annotations=[dict(text=fmt_m(tv_prev), x=.5, y=.5, font_size=13, showarrow=False)])
            st.plotly_chart(fig, use_container_width=True)

    with col_b:
        st.markdown("### Comissão por empresa (Plataforma)")
        if not df_mp.empty:
            by_e = df_mp.groupby('Recebedora')['Valor Previsto'].sum().sort_values(ascending=False).head(10).reset_index()
            fig2 = px.bar(by_e, y='Recebedora', x='Valor Previsto', orientation='h',
                color_discrete_sequence=['#C9A84C'],
                text=by_e['Valor Previsto'].apply(fmt_m))
            fig2.update_traces(textposition='outside')
            fig2.update_layout(margin=dict(t=10,b=0,l=0,r=40), height=260,
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                yaxis=dict(autorange='reversed'), xaxis_title=None, yaxis_title=None)
            st.plotly_chart(fig2, use_container_width=True)

    # Projeção por ano
    st.markdown("### Projeção de Comissões por Ano")
    if not df_mall.empty:
        df_mall['Ano'] = pd.to_datetime(df_mall['Data Prevista'], errors='coerce').dt.year
        pa = df_mall.groupby('Ano')['Valor Previsto'].sum().reset_index().dropna()
        pa['Ano'] = pa['Ano'].astype(int)
        ano_at = date.today().year
        pa['Cor'] = pa['Ano'].apply(lambda x: '#C9A84C' if x==ano_at else '#0F1C2E')
        fig3 = go.Figure(go.Bar(
            x=pa['Ano'].astype(str), y=pa['Valor Previsto'],
            marker_color=pa['Cor'],
            text=pa['Valor Previsto'].apply(fmt_m), textposition='outside',
        ))
        fig3.update_layout(margin=dict(t=20,b=0,l=0,r=0), height=200,
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
            yaxis=dict(gridcolor='rgba(0,0,0,.04)'), xaxis=dict(tickfont=dict(size=12)))
        st.plotly_chart(fig3, use_container_width=True)

    # Auditoria
    if not df_a.empty and 'STATUS AUDITORIA' in df_a.columns:
        st.markdown("### Auditoria · Resumo")
        res = df_a.groupby('STATUS AUDITORIA').agg(Qtd=('STATUS AUDITORIA','count'), Total=('DIFERENÇA (FALTA)','sum')).reset_index()
        ca,cb,cc = st.columns(3)
        mapa = {
            'PROCESSADO CORRETAMENTE':(ca,'#E8F5EE','#1A6B3C','✅ Correto'),
            'COBRAR DIVERGÊNCIA':     (cb,'#FEF3E2','#8A6000','⚠️ Cobrar'),
            'REEMBOLSO':              (cc,'#FDECEA','#B03050','↩️ Reembolso'),
        }
        for _, row in res.iterrows():
            s = row['STATUS AUDITORIA']
            col_r,bg,cor,lbl = mapa.get(s,(ca,'#E8F5EE','#0F1C2E',s))
            col_r.markdown(f"""<div style="background:{bg};border-radius:10px;padding:16px 20px">
                <b style="font-size:28px;font-family:'Cormorant Garamond',serif;color:{cor}">{row['Qtd']}</b><br>
                <span style="font-size:12px;color:{cor};font-weight:600">{lbl}</span><br>
                <span style="font-size:12px;color:{cor}">{fmt_m(abs(row['Total']))}</span>
            </div>""", unsafe_allow_html=True)

    st.markdown(RODAPE, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# PRODUÇÃO
# ══════════════════════════════════════════════════════════════════
elif pag == 'producao':
    st.markdown("## Produção por Período")

    if df_v.empty:
        st.warning("Planilha Tb_Vendas.xlsx não encontrada no repositório.")
    else:
        df_prod = df_v[df_v['Ativo']].copy()
        df_prod['AnoMes']    = df_prod['Data da Venda'].dt.strftime('%Y-%m')
        df_prod['Trimestre'] = df_prod['Data da Venda'].dt.year.astype(str)+'-T'+df_prod['Data da Venda'].dt.quarter.astype(str)
        df_prod['Ano']       = df_prod['Data da Venda'].dt.year.astype(str)
        df_sp = split_consultores(df_prod)

        tab1, tab2, tab3 = st.tabs(["📊 Evolução","🏢 Por Empresa","👤 Por Consultor"])

        with tab1:
            c1,c2,c3,c4 = st.columns(4)
            per  = c1.selectbox("Agrupar por",["Mês","Trimestre","Ano"])
            fde  = c2.date_input("De",  value=None)
            fate = c3.date_input("Até", value=None)
            meta = c4.number_input("Meta por período (R$)", min_value=0.0, step=100000.0)

            df_f = df_prod.copy()
            if fde:  df_f = df_f[df_f['Data da Venda']>=pd.Timestamp(fde)]
            if fate: df_f = df_f[df_f['Data da Venda']<=pd.Timestamp(fate)]
            col_p = {'Mês':'AnoMes','Trimestre':'Trimestre','Ano':'Ano'}[per]
            by_p  = df_f.groupby(col_p).agg(Contratos=('Nº Contrato','count'),Volume=('Crédito_num','sum')).reset_index()
            by_p  = by_p.rename(columns={col_p:'Período'}).sort_values('Período')
            tv_t  = df_f['Crédito_num'].sum()
            melhor = by_p.loc[by_p['Volume'].idxmax(),'Período'] if not by_p.empty else "—"

            m1,m2,m3,m4 = st.columns(4)
            m1.metric("Volume total", fmt_m(tv_t))
            m2.metric("Contratos",    f"{len(df_f):,}")
            m3.metric("Melhor",       melhor)
            m4.metric("Ticket médio", fmt_m(tv_t/len(df_f)) if len(df_f) else "—")

            fig = go.Figure()
            fig.add_trace(go.Bar(x=by_p['Período'],y=by_p['Volume'],name='Volume',
                marker_color='#C9A84C',text=by_p['Volume'].apply(fmt_m),textposition='outside'))
            fig.add_trace(go.Scatter(x=by_p['Período'],y=by_p['Contratos'],name='Contratos',
                yaxis='y2',line=dict(color='#0F1C2E',width=2),mode='lines+markers'))
            if meta>0:
                fig.add_hline(y=meta,line_dash='dash',line_color='#B03050',
                    annotation_text=f"Meta: {fmt_m(meta)}")
            fig.update_layout(margin=dict(t=20,b=0,l=0,r=60),height=320,
                plot_bgcolor='rgba(0,0,0,0)',paper_bgcolor='rgba(0,0,0,0)',
                legend=dict(orientation='h',y=1.12),
                yaxis=dict(gridcolor='rgba(0,0,0,.04)'),
                yaxis2=dict(overlaying='y',side='right',showgrid=False))
            st.markdown(f"### Produção por {per.lower()}")
            st.plotly_chart(fig, use_container_width=True)

        with tab2:
            rk_e = df_prod.groupby('Empresa Vendedora').agg(Volume=('Crédito_num','sum'),Contratos=('Nº Contrato','count')).sort_values('Volume',ascending=False).reset_index()
            tot_e = rk_e['Volume'].sum() or 1
            fig_e = px.bar(rk_e.head(15), y='Empresa Vendedora', x='Volume', orientation='h',
                color_discrete_sequence=['#0F1C2E'],
                text=rk_e.head(15)['Volume'].apply(fmt_m))
            fig_e.update_traces(textposition='outside')
            fig_e.update_layout(margin=dict(t=10,b=0,l=0,r=60),height=420,
                plot_bgcolor='rgba(0,0,0,0)',paper_bgcolor='rgba(0,0,0,0)',
                yaxis=dict(autorange='reversed'),xaxis_title=None,yaxis_title=None)
            st.plotly_chart(fig_e, use_container_width=True)

        with tab3:
            rk_c = df_sp.groupby('CONSULTOR').agg(Volume=('Crédito_num','sum'),Contratos=('Nº Contrato','count')).sort_values('Volume',ascending=False).reset_index()
            tot_c = rk_c['Volume'].sum() or 1
            meds  = ['🥇','🥈','🥉']+[f'#{i}' for i in range(4,16)]
            for i,(_, row) in enumerate(rk_c.head(12).iterrows()):
                pct = row['Volume']/tot_c*100
                st.markdown(f"""
                <div style="display:flex;align-items:center;gap:14px;margin-bottom:8px;
                            background:white;border-radius:8px;padding:10px 16px;border:1px solid #DDD9D0">
                    <span style="font-size:18px;min-width:28px">{meds[i]}</span>
                    <div style="flex:1">
                        <div style="font-weight:600;font-size:13px;color:#0F1C2E">{row['CONSULTOR']}</div>
                        <div style="background:#EDE9E2;height:4px;border-radius:3px;overflow:hidden;margin-top:4px">
                            <div style="background:{'#C9A84C' if i==0 else '#0F1C2E'};width:{pct:.1f}%;height:100%;border-radius:3px"></div>
                        </div>
                    </div>
                    <div style="text-align:right;min-width:140px">
                        <div style="font-weight:700;font-size:14px;font-family:'Cormorant Garamond',serif;color:#0F1C2E">{fmt_brl(row['Volume'])}</div>
                        <div style="font-size:11px;color:#9A9AAA">{row['Contratos']} contratos · {pct:.1f}%</div>
                    </div>
                </div>""", unsafe_allow_html=True)

    st.markdown(RODAPE, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# RANKINGS
# ══════════════════════════════════════════════════════════════════
elif pag == 'rankings':
    st.markdown("## Rankings · Premiações")

    if df_v.empty:
        st.warning("Planilha Tb_Vendas.xlsx não encontrada.")
    else:
        ativos  = df_v[df_v['Ativo']].copy()
        df_cons = split_consultores(ativos)
        rk_con  = df_cons.groupby('CONSULTOR').agg(Volume=('Crédito_num','sum'),Contratos=('Nº Contrato','count')).sort_values('Volume',ascending=False).reset_index()
        rk_emp  = ativos.groupby('Empresa Vendedora').agg(Volume=('Crédito_num','sum'),Contratos=('Nº Contrato','count')).sort_values('Volume',ascending=False).reset_index()
        tot_c   = rk_con['Volume'].sum() or 1
        tot_e   = rk_emp['Volume'].sum() or 1
        meds    = ['🥇','🥈','🥉']+[f'#{i}' for i in range(4,21)]

        tab_c, tab_e = st.tabs(["👤 Consultores","🏢 Empresas"])
        for tab, rk, tot, col_n in [(tab_c,rk_con,tot_c,'CONSULTOR'),(tab_e,rk_emp,tot_e,'Empresa Vendedora')]:
            with tab:
                # Pódio top 3
                if len(rk) >= 1:
                    top3 = rk.head(3)
                    cores_pod = ['#C9A84C','#9A9AAA','#CD7F32']
                    cols_pod  = st.columns(min(3,len(top3)))
                    for i,(_, row) in enumerate(top3.iterrows()):
                        pct = row['Volume']/tot*100
                        bg  = cores_pod[i]
                        cols_pod[i].markdown(f"""
                        <div style="background:{bg}15;border:2px solid {bg};border-radius:12px;padding:20px;text-align:center;margin-bottom:16px">
                            <div style="font-size:24px">{meds[i]}</div>
                            <div style="font-weight:700;font-size:14px;color:#0F1C2E;margin:8px 0">{row[col_n]}</div>
                            <div style="font-size:17px;font-weight:700;font-family:'Cormorant Garamond',serif;color:{bg}">{fmt_m(row['Volume'])}</div>
                            <div style="font-size:11px;color:#9A9AAA;margin-top:4px">{row['Contratos']} contratos · {pct:.1f}%</div>
                        </div>""", unsafe_allow_html=True)
                # Lista completa
                for i,(_, row) in enumerate(rk.head(15).iterrows()):
                    pct = row['Volume']/tot*100
                    st.markdown(f"""
                    <div style="display:flex;align-items:center;gap:12px;margin-bottom:8px;
                                background:white;border-radius:8px;padding:10px 14px;border:1px solid #DDD9D0">
                        <span style="font-size:16px;min-width:28px">{meds[i]}</span>
                        <div style="flex:1">
                            <div style="font-weight:600;font-size:13px;color:#0F1C2E">{row[col_n]}</div>
                            <div style="background:#EDE9E2;height:4px;border-radius:3px;overflow:hidden;margin-top:4px">
                                <div style="background:{'#C9A84C' if i==0 else '#0F1C2E'};width:{pct:.1f}%;height:100%;border-radius:3px"></div>
                            </div>
                        </div>
                        <div style="text-align:right;min-width:130px">
                            <div style="font-weight:700;font-family:'Cormorant Garamond',serif;color:#0F1C2E">{fmt_brl(row['Volume'])}</div>
                            <div style="font-size:11px;color:#9A9AAA">{row['Contratos']} contratos · {pct:.1f}%</div>
                        </div>
                    </div>""", unsafe_allow_html=True)

    st.markdown(RODAPE, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# MOTORES (somente leitura)
# ══════════════════════════════════════════════════════════════════
elif pag in ('motor_com','motor_plat','motor_all'):
    titulos = {
        'motor_com':  ("Motor Comercial",     "Bloco 1 · P1-P10 (ou P1-P24 promocional)"),
        'motor_plat': ("Motor de Plataforma", "Blocos 2 e 3 · P12-P19 e P20-P31"),
        'motor_all':  ("Receitas Unificadas", "Comercial + Plataforma · Fluxo completo"),
    }
    titulo, subtitulo = titulos[pag]
    st.markdown(f"## {titulo}")
    st.markdown(f"<div style='color:#9A9AAA;font-size:13px;margin:-8px 0 20px'>{subtitulo}</div>", unsafe_allow_html=True)

    df_exib = {'motor_com': df_mc, 'motor_plat': df_mp, 'motor_all': df_mall}[pag]

    if df_exib.empty:
        st.warning("Dados não disponíveis. Verifique as planilhas no repositório.")
    else:
        tv = df_exib['Valor Previsto'].sum()
        tv_c = df_mc['Valor Previsto'].sum() if not df_mc.empty else 0
        tv_p = df_mp['Valor Previsto'].sum() if not df_mp.empty else 0

        c1,c2,c3,c4 = st.columns(4)
        c1.metric("Valor Previsto", fmt_m(tv))
        c2.metric("Parcelas",       f"{len(df_exib):,}")
        c3.metric("Contratos",      f"{df_exib['Nº Contrato'].nunique():,}")
        if pag == 'motor_all':
            c4.metric("Comercial / Plat.", f"{fmt_m(tv_c)} / {fmt_m(tv_p)}")
        else:
            c4.metric("Empresas", f"{df_exib['Recebedora'].nunique():,}")

        # Por empresa
        if pag != 'motor_com':
            by_e = df_exib.groupby('Recebedora')['Valor Previsto'].sum().sort_values(ascending=False).head(12).reset_index()
            fig = px.bar(by_e, y='Recebedora', x='Valor Previsto', orientation='h',
                color_discrete_sequence=['#C9A84C'],
                text=by_e['Valor Previsto'].apply(fmt_m))
            fig.update_traces(textposition='outside')
            fig.update_layout(margin=dict(t=10,b=0,l=0,r=60),height=320,
                plot_bgcolor='rgba(0,0,0,0)',paper_bgcolor='rgba(0,0,0,0)',
                yaxis=dict(autorange='reversed'),xaxis_title=None,yaxis_title=None)
            st.markdown("### Por empresa recebedora")
            st.plotly_chart(fig, use_container_width=True)

        # Filtros
        f1,f2,f3 = st.columns(3)
        busca_m = f1.text_input("🔍 Contrato ou empresa")
        f_de    = f2.date_input("Data prevista de",  value=None)
        f_ate   = f3.date_input("Data prevista até", value=None)

        df_mf = df_exib.copy()
        if busca_m:
            from motor import normalizar
            t = normalizar(busca_m)
            df_mf = df_mf[
                df_mf['Nº Contrato'].astype(str).apply(normalizar).str.contains(t,na=False) |
                df_mf['Recebedora'].astype(str).apply(normalizar).str.contains(t,na=False)
            ]
        if f_de:  df_mf = df_mf[pd.to_datetime(df_mf['Data Prevista'],errors='coerce')>=pd.Timestamp(f_de)]
        if f_ate: df_mf = df_mf[pd.to_datetime(df_mf['Data Prevista'],errors='coerce')<=pd.Timestamp(f_ate)]

        df_ms = df_mf[['Nº Contrato','Data Prevista','Parcela','Empresa Vendedora','Recebedora','Tipo','Taxa','Valor Previsto']].copy()
        df_ms['Taxa']           = df_ms['Taxa'].apply(lambda x: f"{x*100:.4f}%")
        df_ms['Valor Previsto'] = df_ms['Valor Previsto'].apply(fmt_brl)
        df_ms['Data Prevista']  = pd.to_datetime(df_ms['Data Prevista'],errors='coerce').dt.strftime('%d/%m/%Y')
        st.dataframe(df_ms.head(300), use_container_width=True, hide_index=True, height=400)
        st.caption(f"Exibindo até 300 de {len(df_mf):,} parcelas")

    st.markdown(RODAPE, unsafe_allow_html=True)
