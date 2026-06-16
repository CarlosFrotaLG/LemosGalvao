import streamlit as st
import pandas as pd
import sqlite3
import os

# ==========================================
# ⚙️ CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(page_title="Central de Importação", page_icon="📥", layout="centered")
st.title("📥 Central de Importação Segura")
st.markdown("""
Este é o seu **Cofre Forte**. Arraste os novos ficheiros para atualizar o sistema. 
O motor possui *Memória Incremental*: **registo duplicados serão ignorados automaticamente!**
""")

# ==========================================
# 🧠 REGRAS DE MEMÓRIA (COMO EVITAR DUPLICADOS)
# ==========================================
REGRAS = {
    "Vendas Novas (tb_vendas)": {
        "tabela_db": "tb_vendas",
        "chaves_unicas": ["Nº Contrato"],
        "skiprows": 0
    },
    "Auditoria Bancorbrás (tb_auditoria)": {
        "tabela_db": "tb_auditoria",
        "chaves_unicas": ["Nº CONTRATO", "PARCELA ATUAL", "EMPRESA RECEBEDORA"],
        "skiprows": 4 # Ignora o cabeçalho bagunçado do Excel
    },
    "Matriz de Comissões (tb_comissionamento)": {
        "tabela_db": "tb_comissionamento",
        "chaves_unicas": ["CHAVE COMBINADA"],
        "skiprows": 1
    },
    "Árvore de Parceiros (tb_parceiros)": {
        "tabela_db": "tb_parceiros",
        "chaves_unicas": ["EMPRESA PARCEIRA", "DATA INÍCIO"],
        "skiprows": 0
    },
    "Campanhas e Promoções (tb_promocoes)": {
        "tabela_db": "tb_promocoes",
        "chaves_unicas": ["Nº Contrato", "Grupo"],
        "skiprows": 0
    }
}

# ==========================================
# 🖥️ INTERFACE DE UTILIZADOR
# ==========================================
tipo_importacao = st.selectbox("Qual é a base de dados que deseja atualizar hoje?", list(REGRAS.keys()))

st.info(f"💡 **Dica:** O sistema vai usar as colunas {REGRAS[tipo_importacao]['chaves_unicas']} para garantir que não entram dados repetidos.")

ficheiro = st.file_uploader("Arraste o ficheiro Excel (.xlsx) ou CSV (.csv)", type=['csv', 'xlsx'])

# Opção avançada caso o ficheiro venha com linhas em branco no topo
with st.expander("⚙️ Opções Avançadas de Leitura"):
    linhas_ignorar = st.number_input("Ignorar linhas no topo do ficheiro:", 
                                     value=REGRAS[tipo_importacao]['skiprows'], 
                                     min_value=0)

if st.button("🚀 INJETAR DADOS NO COFRE", type="primary", use_container_width=True):
    if ficheiro is None:
        st.error("⚠️ Por favor, anexe um ficheiro primeiro.")
    else:
        with st.spinner("A ler o ficheiro, a procurar duplicados e a guardar no cofre..."):
            try:
                # 1. Ler o ficheiro novo enviado pelo utilizador
                if ficheiro.name.endswith('.csv'):
                    df_novo = pd.read_csv(ficheiro, skiprows=linhas_ignorar, dtype=str)
                else:
                    df_novo = pd.read_excel(ficheiro, skiprows=linhas_ignorar, dtype=str)
                
                # Limpar espaços em branco dos nomes das colunas
                df_novo.columns = df_novo.columns.str.strip()

                # 2. Ligar ao Cofre (Banco de Dados)
                conexao = sqlite3.connect('banco_consorcio.db')
                nome_tabela = REGRAS[tipo_importacao]['tabela_db']
                chaves = REGRAS[tipo_importacao]['chaves_unicas']
                
                # 3. Puxar os dados antigos (se existirem)
                try:
                    df_antigo = pd.read_sql_query(f"SELECT * FROM {nome_tabela}", conexao, dtype=str)
                except:
                    # Se a tabela não existir ainda, criamos uma vazia
                    df_antigo = pd.DataFrame()

                tamanho_antigo = len(df_antigo)
                
                # 4. A MÁGICA INCREMENTAL (Fundir e Remover Duplicados)
                if not df_antigo.empty:
                    # Verifica se as chaves existem no ficheiro novo
                    chaves_validas = [c for c in chaves if c in df_novo.columns]
                    
                    if chaves_validas:
                        # Junta tudo
                        df_completo = pd.concat([df_antigo, df_novo], ignore_index=True)
                        # Apaga os duplicados, mantendo sempre a versão MAIS RECENTE (última)
                        df_completo = df_completo.drop_duplicates(subset=chaves_validas, keep='last')
                    else:
                        st.warning(f"Não encontrei a coluna {chaves} no seu ficheiro para verificar duplicados. Os dados foram apenas somados.")
                        df_completo = pd.concat([df_antigo, df_novo], ignore_index=True)
                else:
                    df_completo = df_novo

                # 5. Guardar a versão final limpa de volta no cofre
                df_completo.to_sql(nome_tabela, conexao, if_exists='replace', index=False)
                conexao.close()
                
                # 6. Relatório de Sucesso
                tamanho_novo = len(df_completo)
                registos_adicionados = tamanho_novo - tamanho_antigo
                
                st.success(f"✅ Importação Concluída com Sucesso na tabela `{nome_tabela}`!")
                
                col1, col2, col3 = st.columns(3)
                col1.metric("Registos Anteriores", tamanho_antigo)
                col2.metric("Linhas no Ficheiro", len(df_novo))
                col3.metric("Novos Registos Inseridos", max(0, registos_adicionados))
                
                if registos_adicionados < len(df_novo):
                    st.info(f"🛡️ O sistema bloqueou a entrada de {len(df_novo) - max(0, registos_adicionados)} linhas que já existiam no sistema ou eram repetidas.")
                
            except Exception as e:
                st.error(f"❌ Erro ao processar o ficheiro: {e}")