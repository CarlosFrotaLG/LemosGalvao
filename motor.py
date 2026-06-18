# -*- coding: utf-8 -*-
"""
motor.py — Motor de cálculo central da Lemos Galvão
Compartilhado entre app_local.py e app_vitrine.py
"""
import pandas as pd
import unicodedata
from dateutil.relativedelta import relativedelta

TAXA_PARCELA_PLAT = 0.00125  # 0,125% por parcela de plataforma


def normalizar(s):
    if pd.isna(s) or not s:
        return ''
    return unicodedata.normalize('NFD', str(s)).encode('ascii', 'ignore').decode().upper().strip()


def fmt_brl(v):
    try:
        if pd.isna(v):
            return '—'
        return f"R$ {float(v):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
        return '—'


def fmt_m(v):
    try:
        v = float(v)
        if v >= 1e9: return f"R$ {v/1e9:.2f}B".replace('.', ',')
        if v >= 1e6: return f"R$ {v/1e6:.2f}M".replace('.', ',')
        if v >= 1e3: return f"R$ {v/1e3:.0f}K"
        return fmt_brl(v)
    except:
        return '—'


def sf(v, d=0.0):
    try:
        return float(v) if pd.notna(v) else d
    except:
        return d


def preparar_vendas(df_v):
    """Normaliza a tabela de vendas"""
    df = df_v.copy()
    df['Status'] = df['Status'].fillna('Ativo')
    df['Crédito_num'] = pd.to_numeric(df['Crédito'], errors='coerce').fillna(0)
    df['Data da Venda'] = pd.to_datetime(df['Data da Venda'], errors='coerce')
    df['Ativo'] = ~df['Status'].str.strip().str.upper().str.contains('CANCEL', na=False)
    df['Grupo_str'] = df['Grupo'].apply(
        lambda x: str(int(x)) if pd.notna(x) and isinstance(x, (int, float)) else str(x).strip()
    )
    return df


def preparar_comissionamento(df_c):
    """Limpa e indexa a tabela de comissionamento"""
    cols_ok = [c for c in df_c.columns if c in [
        'PRODUTO', 'CHAVE COMBINADA', 'P1', 'P2', 'P3', 'P4', 'P5',
        'P6', 'P7', 'P8', 'P9', 'P10', 'TOTAL'
    ]]
    df = df_c[cols_ok].copy()
    df = df[pd.to_numeric(df['P1'], errors='coerce').notna()]
    df = df.dropna(subset=['CHAVE COMBINADA'])
    # Indexar sem duplicatas
    idx = {}
    for _, row in df.iterrows():
        ch = str(row['CHAVE COMBINADA']).strip()
        if ch not in idx:
            idx[ch] = row.to_dict()
    return idx


def preparar_promocoes(df_c_raw=None, df_pr=None):
    """
    Retorna dict de grupos em promoção: {grupo: [(dt_ini, dt_fim), ...]}
    Aceita tanto a planilha bruta de comissionamento quanto uma tabela de promoções
    """
    grupos = {}

    # Se vier tabela de promoções separada
    if df_pr is not None and not df_pr.empty:
        for _, rp in df_pr.iterrows():
            g = str(rp.get('Grupo Promocional', rp.get('Grupo', ''))).strip()
            if not g or g in ('nan', 'Grupo Promocional', 'None'):
                continue
            try:
                di = pd.Timestamp(rp.get('Data Início', rp.get('Data Inicio', '')))
                df_ = pd.Timestamp(rp.get('Data Fim', ''))
                if pd.isnull(di) or pd.isnull(df_):
                    continue
                if g not in grupos:
                    grupos[g] = []
                grupos[g].append((di, df_))
            except:
                pass
        return grupos

    # Promoções hardcoded baseadas nos dados reais
    promo_padrao = [
        ('2025-11-01', '2025-12-02', '1605'),
        ('2025-11-28', '2025-12-02', '1604'),
        ('2026-03-01', '2026-04-04', '1607'),
    ]
    for di_s, df_s, g in promo_padrao:
        try:
            di = pd.Timestamp(di_s)
            df_ = pd.Timestamp(df_s)
            if g not in grupos:
                grupos[g] = []
            grupos[g].append((di, df_))
        except:
            pass
    return grupos


def preparar_parceiros(df_p):
    """Monta lookup de parceiros com vigência temporal"""
    lookup = {}
    df = df_p.copy()
    df['DATA INÍCIO'] = pd.to_datetime(df['DATA INÍCIO'], errors='coerce')
    df['DATA FIM'] = pd.to_datetime(df['DATA FIM'], errors='coerce')
    for _, rp in df.iterrows():
        emp = str(rp.get('EMPRESA PARCEIRA', '') or '').strip().upper()
        mae = str(rp.get('QUEM INDICOU (MÃE)', '') or '').strip().upper()
        nivel = str(rp.get('NÍVEL DE PLATAFORMA', '') or '').strip()
        di = rp['DATA INÍCIO']
        df_ = rp['DATA FIM']
        if not emp:
            continue
        if emp not in lookup:
            lookup[emp] = []
        lookup[emp].append((di, df_, nivel, mae))
    return lookup


def get_nivel_mae(lookup, empresa, dv):
    key = str(empresa).strip().upper()
    if key not in lookup:
        return 'Sem Plataforma', ''
    for di, df_, nivel, mae in lookup[key]:
        try:
            if pd.notnull(di) and pd.notnull(df_) and di <= dv <= df_:
                return nivel, mae
        except:
            pass
    return 'Sem Plataforma', ''


def nivel_para_taxa(nivel):
    nl = nivel.lower()
    if '0,25' in nl:
        return 0.0025
    if '0,50' in nl or '0,5%' in nl:
        return 0.0050
    if '0,75' in nl:
        return 0.0075
    if '1%' in nl or '1,00' in nl or 'master' in nl:
        return 0.0100
    return 0.0


def motor_comercial(df_v, idx_c, grupos_promo):
    """
    Bloco 1 — Comissão Comercial
    • Padrão:     P1-P10 com taxas da tb_comissionamento
    • Promocional: P1-P24 com taxa fixa 0,15%
    """
    registros = []
    sem_chave = set()

    for _, row in df_v[df_v['Ativo']].iterrows():
        cred = sf(row.get('Crédito_num', 0))
        if cred <= 0:
            continue
        try:
            dv = pd.Timestamp(row['Data da Venda'])
        except:
            continue
        if pd.isnull(dv):
            continue

        contrato = str(row.get('Nº Contrato', '')).strip()
        cliente = str(row.get('CLIENTE', '') or '').strip()
        vendedora = str(row.get('Empresa Vendedora', '') or '').strip()
        grupo = str(row.get('Grupo_str', '')).strip()
        chave = str(row.get('CHAVE (OCULTA)', '') or '').strip()

        # Verificar promoção
        is_promo = False
        if grupo in grupos_promo:
            for di, df_ in grupos_promo[grupo]:
                if di <= dv <= df_:
                    is_promo = True
                    break

        if is_promo:
            for i in range(1, 25):
                registros.append({
                    'Nº Contrato': contrato, 'Cliente': cliente,
                    'Data Venda': dv,
                    'Data Prevista': dv + relativedelta(months=i),
                    'Parcela': i, 'Empresa Vendedora': vendedora,
                    'Recebedora': vendedora,
                    'Valor Previsto': round(cred * 0.0015, 2),
                    'Taxa': 0.0015, 'Tipo': 'Comercial', 'Promocional': True,
                    'Grupo': grupo,
                })
        else:
            tr = idx_c.get(chave)
            if not tr:
                if chave and chave != 'nan':
                    sem_chave.add(chave)
                continue
            for i in range(1, 11):
                tx = sf(tr.get(f'P{i}', 0))
                if tx <= 0:
                    continue
                registros.append({
                    'Nº Contrato': contrato, 'Cliente': cliente,
                    'Data Venda': dv,
                    'Data Prevista': dv + relativedelta(months=i),
                    'Parcela': i, 'Empresa Vendedora': vendedora,
                    'Recebedora': vendedora,
                    'Valor Previsto': round(cred * tx, 2),
                    'Taxa': tx, 'Tipo': 'Comercial', 'Promocional': False,
                    'Grupo': grupo,
                })

    return pd.DataFrame(registros), sem_chave


def motor_plataforma(df_v, lookup_parceiros):
    """
    Blocos 2 e 3 — Comissão de Plataforma

    Bloco 2 (P12-P19, 0,125% cada):
      0,25% → P12,P13 (vendedora) | P14-P19 (LG Plataforma 3)
      0,50% → P12-P15 (vendedora) | P16-P19 (LG Plataforma 2)
      0,75% → P12-P17 (vendedora) | P18,P19 (LG Plataforma 1)
      1,00% → P12-P19 (vendedora, das vendas de quem indicou)
               P24-P31 (ela mesma, das suas próprias)
      Master → recebe Bloco 3

    Bloco 3 (Master = LG PLANEJAMENTO):
      Vendas até 31/12/2024 → P20-P27
      Vendas 01/01/2025+    → P24-P31
    """
    registros = []

    for _, row in df_v[df_v['Ativo']].iterrows():
        cred = sf(row.get('Crédito_num', 0))
        if cred <= 0:
            continue
        try:
            dv = pd.Timestamp(row['Data da Venda'])
        except:
            continue
        if pd.isnull(dv):
            continue

        contrato = str(row.get('Nº Contrato', '')).strip()
        cliente = str(row.get('CLIENTE', '') or '').strip()
        vendedora = str(row.get('Empresa Vendedora', '') or '').strip()
        grupo = str(row.get('Grupo_str', '')).strip()

        # Nível da vendedora (com fallback para colunas da venda)
        niv_v, mae_v = get_nivel_mae(lookup_parceiros, vendedora, dv)
        if niv_v == 'Sem Plataforma':
            niv_v = str(row.get('NÍVEL DA VENDEDORA', '') or '').strip()
            mae_v = str(row.get('SUBPLATAFORMA MÃE', '') or '').strip().upper()

        tx_v = nivel_para_taxa(niv_v)
        n_v = round(tx_v / TAXA_PARCELA_PLAT) if tx_v > 0 else 0
        bloco2 = list(range(12, 20))

        # Vendedora recebe as primeiras n_v parcelas do bloco 2
        for p in bloco2[:n_v]:
            registros.append({
                'Nº Contrato': contrato, 'Cliente': cliente,
                'Data Venda': dv,
                'Data Prevista': dv + relativedelta(months=p),
                'Parcela': p, 'Empresa Vendedora': vendedora,
                'Recebedora': vendedora,
                'Valor Previsto': round(cred * TAXA_PARCELA_PLAT, 2),
                'Taxa': TAXA_PARCELA_PLAT, 'Tipo': 'Plataforma', 'Grupo': grupo,
            })

        # Subir hierarquia para o restante do bloco 2
        rest = bloco2[n_v:]
        curr = mae_v
        tx_ac = tx_v
        safe = 0
        while rest and curr and curr not in ('', 'NAN', 'NONE') and safe < 10:
            safe += 1
            niv_c, mae_c = get_nivel_mae(lookup_parceiros, curr, dv)
            tx_c = nivel_para_taxa(niv_c)

            if tx_c <= tx_ac:
                # Chegou no topo — master absorve o restante
                if 'master' in niv_c.lower() or tx_c >= 0.01:
                    for p in rest:
                        registros.append({
                            'Nº Contrato': contrato, 'Cliente': cliente,
                            'Data Venda': dv,
                            'Data Prevista': dv + relativedelta(months=p),
                            'Parcela': p, 'Empresa Vendedora': vendedora,
                            'Recebedora': curr,
                            'Valor Previsto': round(cred * TAXA_PARCELA_PLAT, 2),
                            'Taxa': TAXA_PARCELA_PLAT, 'Tipo': 'Plataforma', 'Grupo': grupo,
                        })
                    rest = []
                curr = mae_c
                continue

            n_c = round((tx_c - tx_ac) / TAXA_PARCELA_PLAT)
            for p in rest[:n_c]:
                registros.append({
                    'Nº Contrato': contrato, 'Cliente': cliente,
                    'Data Venda': dv,
                    'Data Prevista': dv + relativedelta(months=p),
                    'Parcela': p, 'Empresa Vendedora': vendedora,
                    'Recebedora': curr,
                    'Valor Previsto': round(cred * TAXA_PARCELA_PLAT, 2),
                    'Taxa': TAXA_PARCELA_PLAT, 'Tipo': 'Plataforma', 'Grupo': grupo,
                })
            rest = rest[n_c:]
            tx_ac = tx_c
            curr = mae_c

        # Bloco 3 — Master (LG PLANEJAMENTO)
        bloco3 = list(range(20, 28)) if dv < pd.Timestamp('2025-01-01') else list(range(24, 32))
        for p in bloco3:
            registros.append({
                'Nº Contrato': contrato, 'Cliente': cliente,
                'Data Venda': dv,
                'Data Prevista': dv + relativedelta(months=p),
                'Parcela': p, 'Empresa Vendedora': vendedora,
                'Recebedora': 'LG PLANEJAMENTO',
                'Valor Previsto': round(cred * TAXA_PARCELA_PLAT, 2),
                'Taxa': TAXA_PARCELA_PLAT, 'Tipo': 'Plataforma', 'Grupo': grupo,
            })

    return pd.DataFrame(registros)


def calcular_situacao(valor_previsto, total_pago, data_prevista):
    """Classifica a situação de uma parcela"""
    import datetime
    hoje = pd.Timestamp.now()
    try:
        dp = pd.Timestamp(data_prevista)
    except:
        dp = hoje
    saldo = round((valor_previsto or 0) - (total_pago or 0), 2)
    if (total_pago or 0) > (valor_previsto or 0):
        return 'Reembolso'
    elif saldo == 0 and (total_pago or 0) > 0:
        return 'Quitado'
    elif 0 < (total_pago or 0) < (valor_previsto or 0):
        return 'Parcial'
    elif saldo > 0 and pd.notnull(dp) and dp < hoje:
        return 'Atrasado'
    else:
        return 'A Vencer'


def split_consultores(df_v):
    """Divide vendas de consultores compartilhados (ex: ANDERSON / DENISE)"""
    rows = []
    for _, row in df_v.iterrows():
        c = str(row.get('CONSULTOR', '') or '')
        if '/' in c:
            partes = [p.strip() for p in c.split('/') if p.strip()]
            for pt in partes:
                r = row.copy()
                r['CONSULTOR'] = pt
                r['Crédito_num'] = row['Crédito_num'] / len(partes)
                rows.append(r)
        else:
            rows.append(row)
    df = pd.DataFrame(rows)
    return df[~df['CONSULTOR'].astype(str).str.strip().isin(['N/A', '', 'nan', 'None', 'NAN'])]
