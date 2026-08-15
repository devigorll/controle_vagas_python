import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io
import os
import datetime

# Mapeamento de vagas limite por sala individual
vagas = [
    {"Sala": 11, "Vaga": 2}, {"Sala": 12, "Vaga": 2}, {"Sala": 13, "Vaga": 2},
    {"Sala": 14, "Vaga": 2}, {"Sala": 15, "Vaga": 2}, {"Sala": 16, "Vaga": 1},
    {"Sala": 21, "Vaga": 2}, {"Sala": 22, "Vaga": 2}, {"Sala": 23, "Vaga": 2},
    {"Sala": 24, "Vaga": 2}, {"Sala": 25, "Vaga": 2}, {"Sala": 26, "Vaga": 2},
    {"Sala": 31, "Vaga": 2}, {"Sala": 32, "Vaga": 2}, {"Sala": 33, "Vaga": 2},
    {"Sala": 34, "Vaga": 2}, {"Sala": 35, "Vaga": 2}, {"Sala": 36, "Vaga": 2},
    {"Sala": 41, "Vaga": 2}, {"Sala": 42, "Vaga": 2}, {"Sala": 43, "Vaga": 2},
    {"Sala": 44, "Vaga": 2}, {"Sala": 45, "Vaga": 2}, {"Sala": 46, "Vaga": 2},
    {"Sala": 51, "Vaga": 2}, {"Sala": 52, "Vaga": 2}, {"Sala": 53, "Vaga": 2},
    {"Sala": 54, "Vaga": 1}, {"Sala": 55, "Vaga": 2}, {"Sala": 56, "Vaga": 2},
    {"Sala": 61, "Vaga": 3}, {"Sala": 62, "Vaga": 2}, {"Sala": 63, "Vaga": 2},
    {"Sala": 64, "Vaga": 2}, {"Sala": 65, "Vaga": 2}, {"Sala": 66, "Vaga": 2},
    {"Sala": 71, "Vaga": 2}, {"Sala": 72, "Vaga": 2}, {"Sala": 73, "Vaga": 4},
    {"Sala": 74, "Vaga": 7}, {"Sala": 81, "Vaga": 2}, {"Sala": 82, "Vaga": 2},
    {"Sala": 85, "Vaga": 2}, {"Sala": 86, "Vaga": 2}, {"Sala": 83, "Vaga": 4},
]

limite_vagas_map = {item["Sala"]: item["Vaga"] for item in vagas}

dias_map = {
    'Monday': 'Segunda-feira',
    'Tuesday': 'Terça-feira',
    'Wednesday': 'Quarta-feira',
    'Thursday': 'Quinta-feira',
    'Friday': 'Sexta-feira',
    'Saturday': 'Sábado',
    'Sunday': 'Domingo'
}

# Configuração visual do dashboard
st.set_page_config(
    page_title="Gestão & Auditoria de Ocupação de Vagas", 
    page_icon="🚗", 
    layout="wide"
)

st.title("🚗 Auditoria e Análise de Ocupação de Vagas")
st.caption("Sistema de Auditoria Simultânea e Inteligência Operacional de Estacionamento")
st.divider()

data_dir = "../data/"

uploaded_file = st.file_uploader(
    "📂 Faça o upload da planilha de ocupação (.xlsx ou .csv)", 
    type=["csv", "xlsx"]
)

def parse_sala_nums(val):
    """Retorna uma tupla ordenada com os números das salas presentes no campo (ex: '61/62' -> (61, 62))"""
    val_str = str(val).strip()
    parts = [p.strip() for p in val_str.split('/') if p.strip()]
    nums = []
    for p in parts:
        try:
            nums.append(int(float(p)))
        except:
            pass
    return tuple(sorted(nums)) if nums else None

def calcular_limite_vagas(sala_nums_tuple, limite_map):
    """Calcula o limite total de vagas somando as salas informadas"""
    if not sala_nums_tuple:
        return 2
    return sum(limite_map.get(num, 2) for num in sala_nums_tuple)

def formatar_valor_br(val):
    """Formata valores numéricos para o padrão financeiro brasileiro R$ 0,00"""
    if pd.isna(val) or str(val).strip() == "":
        return val
    try:
        val_float = float(str(val).replace("R$", "").replace(".", "").replace(",", ".").strip())
        return f"R$ {val_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return val

def analisar_excesso_vagas(df_raw, limite_vagas_map):
    # Tratar cabeçalho deslocado na leitura
    header_idx = None
    for idx, row in df_raw.iterrows():
        row_str = row.astype(str).str.upper().tolist()
        if "SALA" in row_str and "HORA ENTRADA" in row_str:
            header_idx = idx
            break
            
    if header_idx is not None:
        df = df_raw.iloc[header_idx + 1:].copy()
        df.columns = df_raw.iloc[header_idx].values
    else:
        df = df_raw.copy()

    df = df.loc[:, df.columns.notna()].copy()
    
    cols_rename = {}
    for col in df.columns:
        c_str = str(col).strip().upper()
        if c_str == "DATA": cols_rename[col] = "DATA"
        elif c_str == "SALA": cols_rename[col] = "SALA"
        elif c_str in ["QUEM", "NOME", "CLIENTE"]: cols_rename[col] = "QUEM"
        elif c_str == "HORA ENTRADA": cols_rename[col] = "HORA ENTRADA"
        elif c_str in ["HORA SAÍDA", "HORA SAIDA"]: cols_rename[col] = "HORA SAIDA"
    
    df.rename(columns=cols_rename, inplace=True)
    
    required_cols = ["DATA", "SALA", "HORA ENTRADA", "HORA SAIDA"]
    for col in required_cols:
        if col not in df.columns:
            st.error(f"Coluna obrigatória não encontrada: `{col}`")
            return None, None

    df_valid = df.dropna(subset=["DATA", "SALA", "HORA ENTRADA", "HORA SAIDA"]).copy()

    df_valid["SALA_KEY"] = df_valid["SALA"].apply(parse_sala_nums)
    
    def make_dt(row, time_col):
        d = pd.to_datetime(row["DATA"]).date()
        t = row[time_col]
        if isinstance(t, str):
            t = pd.to_datetime(t).time()
        elif isinstance(t, pd.Timestamp):
            t = t.time()
        return pd.Timestamp.combine(d, t)

    try:
        df_valid["DT_ENTRADA"] = df_valid.apply(lambda r: make_dt(r, "HORA ENTRADA"), axis=1)
        df_valid["DT_SAIDA"] = df_valid.apply(lambda r: make_dt(r, "HORA SAIDA"), axis=1)
        df_valid["PERMANENCIA_MIN"] = (df_valid["DT_SAIDA"] - df_valid["DT_ENTRADA"]).dt.total_seconds() / 60.0
        df_valid["DIA_SEMANA"] = pd.to_datetime(df_valid["DATA"]).dt.day_name().map(dias_map)
    except Exception as e:
        st.error(f"Erro ao converter datas/horários: {e}")
        return None, None

    excesso_explicacoes = {}

    grouped = df_valid.groupby(["DATA", "SALA_KEY"])

    for (data, sala_key), group in grouped:
        limite = calcular_limite_vagas(sala_key, limite_vagas_map)
        
        events = []
        for idx, row in group.iterrows():
            events.append((row["DT_ENTRADA"], 'entrada', idx))
            events.append((row["DT_SAIDA"], 'saida', idx))
            
        events.sort(key=lambda x: (x[0], 0 if x[1] == 'saida' else 1))
        
        carros_presentes = []
        
        for time_pt, event_type, idx in events:
            if event_type == 'entrada':
                carros_presentes.append(idx)
                if len(carros_presentes) > limite:
                    for extra_idx in carros_presentes[limite:]:
                        if extra_idx not in excesso_explicacoes:
                            carros_ja_estacionados = carros_presentes[:limite]
                            intervalos = []
                            for c_idx in carros_ja_estacionados:
                                ent_str = group.loc[c_idx, "HORA ENTRADA"]
                                sai_str = group.loc[c_idx, "HORA SAIDA"]
                                if hasattr(ent_str, 'strftime'): ent_str = ent_str.strftime('%H:%M:%S')
                                if hasattr(sai_str, 'strftime'): sai_str = sai_str.strftime('%H:%M:%S')
                                intervalos.append(f"{ent_str} às {sai_str}")
                            
                            qtd_ja = len(carros_ja_estacionados)
                            intervalos_texto = " | ".join(intervalos)
                            txt_vagas = f"{limite} vaga autorizada" if limite == 1 else f"{limite} vagas autorizadas"
                            txt_carros = f"{qtd_ja} veículo simultâneo" if qtd_ja == 1 else f"{qtd_ja} veículos simultâneos"
                            
                            explicacao = (f"Capacidade excedida: a sala possui limite de {txt_vagas}, porém no momento deste acesso "
                                          f"já havia {txt_carros} ocupando o local nos horários: [{intervalos_texto}].")
                            excesso_explicacoes[extra_idx] = explicacao
            else:
                if idx in carros_presentes:
                    carros_presentes.remove(idx)

    # DataFrame de infrações
    df_excesso = df.loc[sorted(list(excesso_explicacoes.keys()))].copy()
    
    # 1. Formatar DATA no padrão BR (DD/MM/YYYY)
    df_excesso["DATA"] = pd.to_datetime(df_excesso["DATA"]).dt.strftime('%d/%m/%Y')
    
    # 2. Inserir coluna QTD VAGAS combinando limites se for composto (ex: 61/62)
    df_excesso["SALA_KEY_TEMP"] = df_excesso["SALA"].apply(parse_sala_nums)
    df_excesso["QTD VAGAS"] = df_excesso["SALA_KEY_TEMP"].apply(lambda k: calcular_limite_vagas(k, limite_vagas_map))
    df_excesso.drop(columns=["SALA_KEY_TEMP"], inplace=True)

    # Reordenar QTD VAGAS para ficar logo após SALA
    cols = list(df_excesso.columns)
    if "SALA" in cols and "QTD VAGAS" in cols:
        cols.remove("QTD VAGAS")
        sala_idx = cols.index("SALA")
        cols.insert(sala_idx + 1, "QTD VAGAS")
        df_excesso = df_excesso[cols]

    # 3. Adicionar coluna MOTIVO_EXCESSO no final
    df_excesso["MOTIVO_EXCESSO"] = [excesso_explicacoes[idx] for idx in df_excesso.index]
    
    # Remoção de faixas de horário (00:00 a 23:30)
    def is_time_column(col):
        if isinstance(col, (datetime.time, pd.Timestamp)):
            return True
        col_str = str(col).strip()
        if len(col_str.split(':')) in [2, 3] and col_str.replace(':', '').isdigit():
            return True
        return False

    cols_to_keep = [col for col in df_excesso.columns if not is_time_column(col)]
    df_excesso = df_excesso[cols_to_keep].copy()

    # Remoção explícita de ANO, MÊS e OBS
    cols_to_remove = ["ANO", "MÊS", "MES", "OBS"]
    df_excesso.drop(columns=[c for c in cols_to_remove if c in df_excesso.columns], inplace=True)

    # 4. Formatação de colunas monetárias/financeiras para R$ VALOR,00 se existirem
    for col in df_excesso.columns:
        if any(term in str(col).upper() for term in ["VALOR", "PRECO", "PREÇO", "TAXA", "CUSTO", "PAGO"]):
            df_excesso[col] = df_excesso[col].apply(formatar_valor_br)

    return df_excesso, df_valid

if uploaded_file is not None:
    os.makedirs(data_dir, exist_ok=True)
    file_path = os.path.join(data_dir, uploaded_file.name)
    
    try:
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        if uploaded_file.name.endswith(".csv"):
            df_raw = pd.read_csv(file_path)
        else:
            df_raw = pd.read_excel(file_path)

        df_excesso, df_valid = analisar_excesso_vagas(df_raw, limite_vagas_map)

    finally:
        if os.path.exists(file_path):
            os.remove(file_path)

    if df_excesso is not None and df_valid is not None:
        
        tab_report, tab_kpi = st.tabs(["📋 Relatório de Excesso", "📊 Dashboard & KPIs"])

        # --- ABA 1: RELATÓRIO DE EXCESSO ---
        with tab_report:
            st.subheader("🚨 Relatório de Veículos Excedentes")
            
            if len(df_excesso) == 0:
                st.success("🎉 Nenhum veículo excedeu o limite de vagas simultâneas no período analisado.")
            else:
                st.warning(f"Foi identificado um total de **{len(df_excesso)}** ocorrências de excesso de capacidade.")
                
                # Exibir DataFrame Final
                st.dataframe(df_excesso, use_container_width=True)
                
                # Preparar arquivos para Download
                csv_data = df_excesso.to_csv(index=False).encode('utf-8')
                buffer_xlsx = io.BytesIO()
                with pd.ExcelWriter(buffer_xlsx, engine='openpyxl') as writer:
                    df_excesso.to_excel(writer, index=False, sheet_name='Veiculos_Excedentes')
                xlsx_data = buffer_xlsx.getvalue()

                c1, c2 = st.columns(2)
                with c1:
                    st.download_button(
                        label="📥 Baixar Report Infratores (CSV)",
                        data=csv_data,
                        file_name="report_veiculos_excedentes.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                with c2:
                    st.download_button(
                        label="📥 Baixar Report Infratores (Excel)",
                        data=xlsx_data,
                        file_name="report_veiculos_excedentes.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

        # --- ABA 2: DASHBOARD & KPIS ---
        with tab_kpi:
            st.subheader("📊 Indicadores de Desempenho e Utilização")
            
            total_acessos = len(df_valid)
            total_excesso = len(df_excesso)
            taxa_infracao = (total_excesso / total_acessos * 100) if total_acessos > 0 else 0
            media_permanencia = df_valid["PERMANENCIA_MIN"].mean()

            df_excesso_valid = df_valid.loc[df_valid.index.isin(df_excesso.index)].copy()
            dia_mais_excedido = df_excesso_valid["DIA_SEMANA"].mode()[0] if not df_excesso_valid.empty else "N/A"

            # Cards
            kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
            
            with kpi_col1:
                st.metric(label="Total de Veículos Estacionados", value=f"{total_acessos}")
            with kpi_col2:
                st.metric(label="Total de Infração de Vagas", value=f"{total_excesso}", delta=f"{taxa_infracao:.1f}% do total", delta_color="inverse")
            with kpi_col3:
                st.metric(label="Média Tempo Estacionado", value=f"{media_permanencia:.0f} min")
            with kpi_col4:
                st.metric(label="Dia c/ Maior Estouro de Vagas", value=f"{dia_mais_excedido}")

            st.divider()

            # Gráficos
            g_col1, g_col2 = st.columns(2)
            ordem_dias = ['Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado', 'Domingo']

            with g_col1:
                st.markdown("##### 📅 Volume de Veículos Estacionados por Dia da Semana")
                counts_geral = df_valid["DIA_SEMANA"].value_counts().reindex(ordem_dias).fillna(0).reset_index()
                counts_geral.columns = ["Dia da Semana", "Quantidade"]
                
                fig1 = px.bar(
                    counts_geral, 
                    x="Dia da Semana", 
                    y="Quantidade", 
                    text="Quantidade",
                    color_discrete_sequence=["#1f77b4"]
                )
                fig1.update_layout(xaxis_title="", yaxis_title="Nº de Veículos")
                st.plotly_chart(fig1, use_container_width=True)

            with g_col2:
                st.markdown("##### 🚨 Infrações (Excesso de Vagas) por Dia da Semana")
                if not df_excesso_valid.empty:
                    counts_excesso = df_excesso_valid["DIA_SEMANA"].value_counts().reindex(ordem_dias).fillna(0).reset_index()
                    counts_excesso.columns = ["Dia da Semana", "Infrações"]
                    
                    fig2 = px.bar(
                        counts_excesso, 
                        x="Dia da Semana", 
                        y="Infrações", 
                        text="Infrações",
                        color_discrete_sequence=["#d62728"]
                    )
                    fig2.update_layout(xaxis_title="", yaxis_title="Nº de Excedentes")
                    st.plotly_chart(fig2, use_container_width=True)
                else:
                    st.info("Sem dados de infração para exibir.")

            st.divider()

