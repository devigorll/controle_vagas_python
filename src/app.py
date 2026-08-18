import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io
import datetime
import re

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

uploaded_file = st.file_uploader(
    "📂 Faça o upload da planilha de ocupação (.xlsx ou .csv)", 
    type=["csv", "xlsx"]
)

def parse_sala_nums(val):
    """Retorna uma tupla ordenada com os números das salas presentes no campo (ex: '61/62' -> (61, 62))"""
    if pd.isna(val):
        return None
    val_str = str(val).strip()
    parts = [p.strip() for p in val_str.split('/') if p.strip()]
    nums = []
    for p in parts:
        try:
            nums.append(int(float(p)))
        except (ValueError, TypeError):
            pass
    return tuple(sorted(nums)) if nums else None

def carregar_limite_vagas(file_input):
    """
    Lê a aba 'QTD VAGAS X SALA' do arquivo Excel e constrói dinamicamente
    o dicionário de limite de vagas por chave de sala {sala_key: qtd_vagas}.
    """
    try:
        xls = pd.ExcelFile(file_input)
        
        target_sheet = None
        for sheet in xls.sheet_names:
            sheet_upper = sheet.upper()
            if "QTD VAGAS" in sheet_upper or "VAGAS X SALA" in sheet_upper or "VAGAS" in sheet_upper:
                target_sheet = sheet
                break
        
        if not target_sheet:
            return {}

        df_vagas_raw = pd.read_excel(xls, sheet_name=target_sheet)

        header_idx = None
        col_sala_idx = None
        col_vaga_idx = None

        for idx, row in df_vagas_raw.iterrows():
            row_vals = [str(val).strip().upper() for val in row.values if pd.notna(val)]
            if any("SALA" in v for v in row_vals) and any(v in row_vals for v in ["VAGA", "VAGAS", "QTD VAGAS", "QTD_VAGAS"]):
                header_idx = idx
                for c_i, val in enumerate(row.values):
                    val_str = str(val).strip().upper()
                    if "SALA" in val_str:
                        col_sala_idx = c_i
                    elif val_str in ["VAGA", "VAGAS", "QTD VAGAS", "QTD_VAGAS"]:
                        col_vaga_idx = c_i
                break

        if header_idx is None or col_sala_idx is None or col_vaga_idx is None:
            return {}

        df_vagas = df_vagas_raw.iloc[header_idx + 1:].copy()
        
        limite_map = {}
        for _, row in df_vagas.iterrows():
            sala_val = row.iloc[col_sala_idx]
            vaga_val = row.iloc[col_vaga_idx]

            try:
                if pd.notna(sala_val) and pd.notna(vaga_val):
                    sala_key = parse_sala_nums(sala_val)
                    vaga_num = int(float(str(vaga_val).strip()))
                    if sala_key:
                        limite_map[sala_key] = vaga_num
            except (ValueError, TypeError):
                continue

        return limite_map

    except Exception:
        return {}

def calcular_limite_vagas(sala_nums_tuple, limite_map):
    """Calcula o limite total de vagas considerando salas compostas ou individuais"""
    if not sala_nums_tuple:
        return 2
    
    if sala_nums_tuple in limite_map:
        return limite_map[sala_nums_tuple]
        
    total = 0
    for num in sala_nums_tuple:
        total += limite_map.get((num,), 2)
    return total

def formatar_valor_br(val):
    """Formata valores numéricos para o padrão financeiro brasileiro R$ 0,00"""
    if pd.isna(val) or str(val).strip() == "":
        return val
    try:
        val_float = float(str(val).replace("R$", "").replace(".", "").replace(",", ".").strip())
        return f"R$ {val_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return val

def sanitizar_nome_arquivo(texto):
    """Remove caracteres especiais e espaços, convertendo para minúsculo"""
    texto = texto.lower()
    texto = re.sub(r'[áàâã]', 'a', texto)
    texto = re.sub(r'[éèê]', 'e', texto)
    texto = re.sub(r'[íìî]', 'i', texto)
    texto = re.sub(r'[óòôõ]', 'o', texto)
    texto = re.sub(r'[úùû]', 'u', texto)
    texto = re.sub(r'[ç]', 'c', texto)
    texto = re.sub(r'[^a-z0-9_]', '_', texto)
    texto = re.sub(r'_+', '_', texto)
    return texto.strip('_')

def gerar_nome_arquivo_download(data_range, min_date, max_date, salas_selecionadas, salas_disponiveis):
    """Gera o nome customizado do arquivo conforme os filtros ativos"""
    filtro_data_ativo = False
    str_data = ""
    
    if isinstance(data_range, (tuple, list)) and len(data_range) == 2:
        start_d, end_d = data_range
        if start_d != min_date or end_d != max_date:
            filtro_data_ativo = True
            if start_d == end_d:
                str_data = start_d.strftime('%d-%m-%Y')
            else:
                str_data = f"{start_d.strftime('%d-%m-%Y')}_a_{end_d.strftime('%d-%m-%Y')}"
    elif isinstance(data_range, (tuple, list)) and len(data_range) == 1:
        if data_range[0] != min_date or min_date != max_date:
            filtro_data_ativo = True
            str_data = data_range[0].strftime('%d-%m-%Y')

    filtro_sala_ativo = False
    str_sala = ""
    
    if set(salas_selecionadas) != set(salas_disponiveis):
        filtro_sala_ativo = True
        str_sala = "_".join([str(s) for s in sorted(salas_selecionadas)])

    partes = ["relatorio_excesso"]
    
    if filtro_sala_ativo and str_sala:
        partes.append(f"sala_{str_sala}")
        
    if filtro_data_ativo and str_data:
        partes.append(f"dia_{str_data}")

    nome_base = "_".join(partes)
    return sanitizar_nome_arquivo(nome_base)

def analisar_excesso_vagas(df_raw, limite_vagas_map):
    header_idx = None
    for idx, row in df_raw.iterrows():
        row_str = [str(val).upper().strip() for val in row.values if pd.notna(val)]
        
        has_sala = any("SALA" in s for s in row_str)
        has_entrada = any("HORA ENTRADA" in s or "ENTRADA" in s for s in row_str)
        
        if has_sala and has_entrada:
            header_idx = idx
            break
            
    if header_idx is not None:
        df = df_raw.iloc[header_idx + 1:].copy()
        df.columns = [str(c).strip() for c in df_raw.iloc[header_idx].values]
    else:
        df = df_raw.copy()

    df = df.loc[:, df.columns.notna()].copy()
    
    cols_rename = {}
    for col in df.columns:
        c_str = str(col).strip().upper()
        if c_str == "DATA": cols_rename[col] = "DATA"
        elif c_str == "SALA": cols_rename[col] = "SALA"
        elif c_str in ["QUEM", "NOME", "CLIENTE"]: cols_rename[col] = "QUEM"
        elif "ENTRADA" in c_str: cols_rename[col] = "HORA ENTRADA"
        elif "SAÍDA" in c_str or "SAIDA" in c_str: cols_rename[col] = "HORA SAIDA"
    
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
            t = pd.to_datetime(t, format='mixed').time()
        elif isinstance(t, pd.Timestamp):
            t = t.time()
        elif isinstance(t, datetime.time):
            t = t
        else:
            t = pd.to_datetime(str(t)).time()
        return pd.Timestamp.combine(d, t)

    try:
        df_valid["DT_ENTRADA"] = df_valid.apply(lambda r: make_dt(r, "HORA ENTRADA"), axis=1)
        df_valid["DT_SAIDA"] = df_valid.apply(lambda r: make_dt(r, "HORA SAIDA"), axis=1)
        df_valid["PERMANENCIA_MIN"] = (df_valid["DT_SAIDA"] - df_valid["DT_ENTRADA"]).dt.total_seconds() / 60.0
        df_valid["DIA_SEMANA"] = pd.to_datetime(df_valid["DATA"]).dt.day_name().map(dias_map)
    except Exception as e:
        st.error(f"Erro ao converter datas/horários: {e}")
        return None, None

    registro_status = {}
    envolvidos_indices = set()

    grouped = df_valid.groupby(["DATA", "SALA_KEY"])

    for (data, sala_key), group in grouped:
        limite = calcular_limite_vagas(sala_key, limite_vagas_map)
        
        events = []
        for idx, row in group.iterrows():
            events.append((row["DT_ENTRADA"], 'entrada', idx))
            events.append((row["DT_SAIDA"], 'saida', idx))
            
        events.sort(key=lambda x: (x[0], 0 if x[1] == 'saida' else 1))
        
        carros_presentes = []
        contador_sessao = 0
        sessao_indices = set()
        sessao_teve_excesso = False
        
        for time_pt, event_type, idx in events:
            if event_type == 'entrada':
                carros_presentes.append(idx)
                sessao_indices.add(idx)
                
                # Se o pátio estava vazio, inicia uma nova sessão
                if len(carros_presentes) == 1:
                    contador_sessao = 1
                else:
                    contador_sessao += 1
                
                posicao = contador_sessao
                
                if idx not in registro_status:
                    is_excesso = posicao > limite
                    obs_str = f"{posicao}ª vaga (excesso)" if is_excesso else f"{posicao}ª vaga"
                    excedeu_str = "SIM" if is_excesso else ""
                    
                    registro_status[idx] = {
                        "obs": obs_str,
                        "excedeu": excedeu_str
                    }

                # Marca a sessão como com excesso se estourou a capacidade cumulativa ou física
                if posicao > limite or len(carros_presentes) > limite:
                    sessao_teve_excesso = True

            else:
                if idx in carros_presentes:
                    carros_presentes.remove(idx)
                
                # Quando o pátio esvazia totalmente, consolida a sessão
                if len(carros_presentes) == 0:
                    if sessao_teve_excesso:
                        envolvidos_indices.update(sessao_indices)
                    
                    # Reseta os manipuladores da sessão
                    sessao_indices = set()
                    sessao_teve_excesso = False
                    contador_sessao = 0

        # Se o dia acabou e a sessão ainda estava aberta com excesso
        if sessao_teve_excesso:
            envolvidos_indices.update(sessao_indices)

    df_analisado = df_valid.loc[sorted(list(envolvidos_indices))].copy()
    
    if df_analisado.empty:
        return pd.DataFrame(), df_valid

    df_analisado["QTD VAGAS"] = df_analisado["SALA_KEY"].apply(lambda k: calcular_limite_vagas(k, limite_vagas_map))

    cols = list(df_analisado.columns)
    if "SALA" in cols and "QTD VAGAS" in cols:
        cols.remove("QTD VAGAS")
        sala_idx = cols.index("SALA")
        cols.insert(sala_idx + 1, "QTD VAGAS")
        df_analisado = df_analisado[cols]

    df_analisado["DATA_DT"] = pd.to_datetime(df_analisado["DATA"])
    df_analisado["DATA_FMT"] = df_analisado["DATA_DT"].dt.strftime('%d/%m/%Y')
    df_analisado["DATA"] = df_analisado["DATA_FMT"]

    def is_time_column(col):
        if isinstance(col, (datetime.time, pd.Timestamp)):
            return True
        col_str = str(col).strip()
        if len(col_str.split(':')) in [2, 3] and col_str.replace(':', '').isdigit():
            return True
        return False

    cols_to_keep = [col for col in df_analisado.columns if not is_time_column(col) and col not in ["SALA_KEY", "DT_ENTRADA", "DT_SAIDA", "PERMANENCIA_MIN", "DATA_FMT"]]
    df_analisado = df_analisado[cols_to_keep].copy()

    cols_to_remove = ["ANO", "MÊS", "MES", "OBS"]
    df_analisado.drop(columns=[c for c in cols_to_remove if c in df_analisado.columns], inplace=True)

    for col in df_analisado.columns:
        if any(term in str(col).upper() for term in ["VALOR", "PRECO", "PREÇO", "TAXA", "CUSTO", "PAGO"]):
            df_analisado[col] = df_analisado[col].apply(formatar_valor_br)

    df_analisado["OBS"] = df_analisado.index.map(lambda i: registro_status.get(i, {}).get("obs", ""))
    df_analisado["EXCEDEU"] = df_analisado.index.map(lambda i: registro_status.get(i, {}).get("excedeu", ""))

    return df_analisado, df_valid

if uploaded_file is not None:
    file_bytes = io.BytesIO(uploaded_file.getvalue())

    limite_vagas_map = carregar_limite_vagas(file_bytes)

    file_bytes.seek(0)

    if uploaded_file.name.endswith(".csv"):
        df_raw = pd.read_csv(file_bytes)
    else:
        xls = pd.ExcelFile(file_bytes)
        sheet_detalhe = 'DETALHE' if 'DETALHE' in xls.sheet_names else xls.sheet_names[0]
        df_raw = pd.read_excel(xls, sheet_name=sheet_detalhe)

    df_analisado, df_valid = analisar_excesso_vagas(df_raw, limite_vagas_map)

    if df_analisado is not None and df_valid is not None:
        
        # --- SIDEBAR: FILTROS ---
        st.sidebar.header("🔍 Filtros de Análise")
        
        if not df_analisado.empty:
            # 1. Filtro de Data
            min_date = df_analisado["DATA_DT"].min().date()
            max_date = df_analisado["DATA_DT"].max().date()
            
            data_range = st.sidebar.date_input(
                "📅 Período (Data)",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date
            )

            # 2. Filtro de Sala
            salas_disponiveis = sorted(df_analisado["SALA"].astype(str).unique())
            salas_selecionadas = st.sidebar.multiselect(
                "🏢 Sala",
                options=salas_disponiveis,
                default=salas_disponiveis
            )

            # Aplicação dos Filtros
            df_filtrado = df_analisado.copy()

            # Filtrar por Data
            if isinstance(data_range, (tuple, list)) and len(data_range) == 2:
                start_d, end_d = data_range
                df_filtrado = df_filtrado[
                    (df_filtrado["DATA_DT"].dt.date >= start_d) & 
                    (df_filtrado["DATA_DT"].dt.date <= end_d)
                ]
            elif isinstance(data_range, (tuple, list)) and len(data_range) == 1:
                df_filtrado = df_filtrado[df_filtrado["DATA_DT"].dt.date == data_range[0]]

            # Filtrar por Sala
            if salas_selecionadas:
                df_filtrado = df_filtrado[df_filtrado["SALA"].astype(str).isin(salas_selecionadas)]

            df_exibicao = df_filtrado.drop(columns=["DATA_DT"]).copy()

            # Gerar nome customizado do arquivo de download baseado nos filtros ativos
            nome_arquivo_base = gerar_nome_arquivo_download(
                data_range, min_date, max_date, salas_selecionadas, salas_disponiveis
            )
        else:
            df_exibicao = pd.DataFrame()
            df_filtrado = pd.DataFrame()
            nome_arquivo_base = "relatorio_excesso"

        # --- ABAS DE EXIBIÇÃO ---
        tab_report, tab_kpi = st.tabs(["📋 Relatório de Eventos de Excesso", "📊 Dashboard & KPIs"])

        # --- ABA 1: RELATÓRIO DE OCUPAÇÃO ---
        with tab_report:
            st.subheader("📋 Ocorrências de Ocupação e Excesso")
            
            if df_exibicao.empty:
                st.success("🎉 Nenhuma infração ou excesso de vagas encontrado no período selecionado.")
            else:
                total_reg = len(df_exibicao)
                total_exc = (df_exibicao["EXCEDEU"] == "SIM").sum()
                
                st.info(f"Exibindo **{total_reg}** registros envolvidos em ocorrências de excesso (**{total_exc}** veículos excederam diretamente a capacidade).")
                
                # Exibir DataFrame Final
                st.dataframe(df_exibicao, use_container_width=True)
                
                # Preparar arquivos para Download com Nome Dinâmico
                csv_data = df_exibicao.to_csv(index=False).encode('utf-8')
                buffer_xlsx = io.BytesIO()
                with pd.ExcelWriter(buffer_xlsx, engine='openpyxl') as writer:
                    df_exibicao.to_excel(writer, index=False, sheet_name='Ocupacao_Excedente')
                xlsx_data = buffer_xlsx.getvalue()

                c1, c2 = st.columns(2)
                with c1:
                    st.download_button(
                        label="📥 Baixar Relatório Filtrado (CSV)",
                        data=csv_data,
                        file_name=f"{nome_arquivo_base}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                with c2:
                    st.download_button(
                        label="📥 Baixar Relatório Filtrado (Excel)",
                        data=xlsx_data,
                        file_name=f"{nome_arquivo_base}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

        # --- ABA 2: DASHBOARD & KPIS ---
        with tab_kpi:
            st.subheader("📊 Indicadores de Desempenho e Utilização")
            
            if not df_filtrado.empty:
                df_valid_filtrado = df_valid.loc[df_valid.index.isin(df_filtrado.index)].copy()

                total_acessos = len(df_valid_filtrado)
                total_excesso = (df_filtrado["EXCEDEU"] == "SIM").sum()
                taxa_infracao = (total_excesso / total_acessos * 100) if total_acessos > 0 else 0
                media_permanencia = df_valid_filtrado["PERMANENCIA_MIN"].mean() if not df_valid_filtrado.empty else 0

                df_excesso_valid = df_valid_filtrado[df_filtrado.loc[df_valid_filtrado.index, "EXCEDEU"] == "SIM"]
                dia_mais_excedido = df_excesso_valid["DIA_SEMANA"].mode()[0] if not df_excesso_valid.empty else "N/A"

                # Cards
                kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
                
                with kpi_col1:
                    st.metric(label="Total de Veículos Analisados", value=f"{total_acessos}")
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
                    st.markdown("##### 📅 Volume de Veículos Envolvidos por Dia da Semana")
                    if not df_valid_filtrado.empty:
                        counts_geral = df_valid_filtrado["DIA_SEMANA"].value_counts().reindex(ordem_dias).fillna(0).reset_index()
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
                    else:
                        st.info("Sem dados de volume para exibir.")

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
            else:
                st.info("Nenhum dado para exibir no dashboard.")