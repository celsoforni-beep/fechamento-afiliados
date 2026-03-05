import io
import datetime as dt
import re
import unicodedata

import streamlit as st
import pandas as pd

st.set_page_config(page_title="Fechamento Afiliados", layout="wide")
st.title("Fechamento mensal de comissões de afiliados")
st.caption("Fluxo: subir Plan Base (e-commerce) + Plan Afiliados (afiliados) → baixar Plan Afiliados preenchida.")

# ===================== CONFIG (SEUS CABEÇALHOS) =====================
# Plan Base (e-commerce)
ECOM_ORDER_COL = "Order2"
ECOM_STATUS_COL = "Status"
ECOM_FREIGHT_COL = "Shipping Value"     # Frete RATEADO por item -> SUM por pedido
ECOM_ORDER_VALUE_COL = "Total Value"    # Normalmente repetido por item -> MAX por pedido
ECOM_EMISSION_DATE_COL = "Creation D"   # <<< seu arquivo usa "Creation D"

# Plan Afiliados
AFF_ORDER_COL = "Order ID"

# Status que liberam comissão
FATURADO_VALUES = {"faturado"}  # comparação normalizada

# Colunas alvo no output (Plan Afiliados)
AFF_OUT_STATUS_COL = "Status"
AFF_OUT_ORDER_VALUE_COL = "Valor do Pedidos"
AFF_OUT_FREIGHT_COL = "Valor de Frete"
AFF_OUT_NET_COL = "Valor S/frete"
# ====================================================================

# --------------------- Helpers ---------------------
def read_any(file) -> pd.DataFrame:
    if file.name.lower().endswith(".csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)

def normalize_text(x) -> str:
    """lower + strip + remove acentos + normaliza espaços"""
    s = str(x) if x is not None else ""
    s = s.strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s)
    return s

def ensure_cols(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    for c in cols:
        if c not in df.columns:
            df[c] = pd.NA
    return df

def drop_blank_technical_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Remove colunas Unnamed:*, nomes numéricos (0,1,2...) e colunas 100% vazias."""
    cols_to_drop = []
    for c in df.columns:
        cs = str(c)
        if cs.startswith("Unnamed:"):
            cols_to_drop.append(c)
        elif re.fullmatch(r"\d+", cs):  # colunas com nome "0", "1", etc.
            cols_to_drop.append(c)

    df = df.drop(columns=cols_to_drop, errors="ignore")

    # drop colunas totalmente vazias
    empty_cols = [c for c in df.columns if df[c].isna().all()]
    df = df.drop(columns=empty_cols, errors="ignore")

    return df

# --------------------- Parâmetros do fechamento ---------------------
st.subheader("Parâmetros do fechamento")

ref_month = st.date_input(
    "Mês de referência do fechamento (qualquer dia dentro do mês)",
    value=dt.date.today().replace(day=1)
)
cutoff_day = st.number_input("Dia limite (corte)", min_value=1, max_value=31, value=20, step=1)
cutoff_date = dt.date(ref_month.year, ref_month.month, int(cutoff_day))
st.caption(f"Corte: pedidos com emissão até {cutoff_date} e status selecionado serão marcados como Cancelado.")

status_forced_value = st.text_input(
    "Status aplicado pela regra (valor reportado)",
    value="Cancelado"
).strip()

# --------------------- Upload ---------------------
col1, col2 = st.columns(2)
with col1:
    ecom_file = st.file_uploader("📄 Suba a Plan Base (e-commerce) - .csv ou .xlsx", type=["csv", "xlsx"])
with col2:
    aff_file = st.file_uploader("📄 Suba a Plan Afiliados - .csv ou .xlsx", type=["csv", "xlsx"])

if not (ecom_file and aff_file):
    st.warning("Suba os dois arquivos para processar e gerar o fechamento.")
    st.stop()

# --------------------- Leitura ---------------------
try:
    ecom = read_any(ecom_file)
    aff = read_any(aff_file)
except Exception as e:
    st.error(f"Erro ao ler os arquivos: {e}")
    st.stop()

# Limpa colunas técnicas já na entrada (principalmente na Plan Afiliados)
aff = drop_blank_technical_columns(aff)

# --------------------- Validação ---------------------
needed_ecom = {ECOM_ORDER_COL, ECOM_STATUS_COL, ECOM_FREIGHT_COL, ECOM_ORDER_VALUE_COL, ECOM_EMISSION_DATE_COL}
needed_aff = {AFF_ORDER_COL}

missing_ecom = needed_ecom - set(ecom.columns)
missing_aff = needed_aff - set(aff.columns)

if missing_ecom:
    st.error(
        "Faltam colunas na Plan Base (e-commerce): "
        f"{sorted(list(missing_ecom))}\n\n"
        f"Colunas existentes: {list(ecom.columns)}"
    )
    st.stop()

if missing_aff:
    st.error(
        f"Faltam colunas na Plan Afiliados: {sorted(list(missing_aff))}\n\n"
        f"Colunas existentes: {list(aff.columns)}"
    )
    st.stop()

# --------------------- Normalização / parsing ---------------------
ecom[ECOM_ORDER_COL] = ecom[ECOM_ORDER_COL].astype(str).str.strip()
aff[AFF_ORDER_COL] = aff[AFF_ORDER_COL].astype(str).str.strip()

ecom[ECOM_FREIGHT_COL] = pd.to_numeric(ecom[ECOM_FREIGHT_COL], errors="coerce").fillna(0)
ecom[ECOM_ORDER_VALUE_COL] = pd.to_numeric(ecom[ECOM_ORDER_VALUE_COL], errors="coerce").fillna(0)

# Data (pt-BR geralmente dayfirst=True)
ecom[ECOM_EMISSION_DATE_COL] = pd.to_datetime(
    ecom[ECOM_EMISSION_DATE_COL],
    errors="coerce",
    dayfirst=True
)

# --------------------- Dropdown de status alvo (pega direto da base) ---------------------
unique_status = sorted([s for s in ecom[ECOM_STATUS_COL].dropna().astype(str).unique()])
status_target_value = st.selectbox(
    "Qual status deve virar CANCELADO quando emissão <= corte?",
    options=unique_status,
    index=unique_status.index("Preparando Entrega") if "Preparando Entrega" in unique_status else 0
)
status_target_norm = normalize_text(status_target_value)

# --------------------- Consolidação por pedido ---------------------
freight_by_order = (
    ecom.groupby(ECOM_ORDER_COL, as_index=False)[ECOM_FREIGHT_COL]
    .sum()
    .rename(columns={ECOM_FREIGHT_COL: "frete_consolidado"})
)

value_by_order = (
    ecom.groupby(ECOM_ORDER_COL, as_index=False)[ECOM_ORDER_VALUE_COL]
    .max()
    .rename(columns={ECOM_ORDER_VALUE_COL: "valor_pedido_consolidado"})
)

emissao_by_order = (
    ecom.groupby(ECOM_ORDER_COL, as_index=False)[ECOM_EMISSION_DATE_COL]
    .min()
    .rename(columns={ECOM_EMISSION_DATE_COL: "data_emissao_consolidada"})
)

status_by_order = (
    ecom[[ECOM_ORDER_COL, ECOM_STATUS_COL]]
    .dropna(subset=[ECOM_STATUS_COL])
    .groupby(ECOM_ORDER_COL, as_index=False)[ECOM_STATUS_COL]
    .first()
    .rename(columns={ECOM_STATUS_COL: "status_pedido"})
)

cons = status_by_order.merge(freight_by_order, how="outer", on=ECOM_ORDER_COL)
cons = cons.merge(value_by_order, how="outer", on=ECOM_ORDER_COL)
cons = cons.merge(emissao_by_order, how="outer", on=ECOM_ORDER_COL)

cons["frete_consolidado"] = cons["frete_consolidado"].fillna(0)
cons["valor_pedido_consolidado"] = cons["valor_pedido_consolidado"].fillna(0)
cons["status_pedido"] = cons["status_pedido"].fillna("Pedido não encontrado na base")

# --------------------- Regra de corte (status -> cancelado) ---------------------
cutoff_ts = pd.Timestamp(cutoff_date)

cons["status_norm"] = cons["status_pedido"].map(normalize_text)

mask_regra = (
    (cons["status_norm"] == status_target_norm) &
    (cons["data_emissao_consolidada"].notna()) &
    (cons["data_emissao_consolidada"] <= cutoff_ts)
)

cons.loc[mask_regra, "status_pedido"] = status_forced_value
cons.drop(columns=["status_norm"], inplace=True, errors="ignore")

# --------------------- Merge na planilha de afiliados ---------------------
resultado = ensure_cols(
    aff.copy(),
    [AFF_OUT_STATUS_COL, AFF_OUT_ORDER_VALUE_COL, AFF_OUT_FREIGHT_COL, AFF_OUT_NET_COL],
)

tmp = cons.rename(columns={ECOM_ORDER_COL: AFF_ORDER_COL}).copy()
tmp["status_norm"] = tmp["status_pedido"].map(normalize_text)

tmp[AFF_OUT_NET_COL] = (tmp["valor_pedido_consolidado"] - tmp["frete_consolidado"]).clip(lower=0)
tmp.loc[~tmp["status_norm"].isin(FATURADO_VALUES), AFF_OUT_NET_COL] = 0

tmp.rename(
    columns={
        "status_pedido": AFF_OUT_STATUS_COL,
        "valor_pedido_consolidado": AFF_OUT_ORDER_VALUE_COL,
        "frete_consolidado": AFF_OUT_FREIGHT_COL,
    },
    inplace=True,
)

resultado = resultado.merge(
    tmp[[AFF_ORDER_COL, AFF_OUT_STATUS_COL, AFF_OUT_ORDER_VALUE_COL, AFF_OUT_FREIGHT_COL, AFF_OUT_NET_COL]],
    on=AFF_ORDER_COL,
    how="left",
    suffixes=("", "_novo"),
)

for col in [AFF_OUT_STATUS_COL, AFF_OUT_ORDER_VALUE_COL, AFF_OUT_FREIGHT_COL, AFF_OUT_NET_COL]:
    novo = f"{col}_novo"
    if novo in resultado.columns:
        resultado[col] = resultado[novo]
        resultado.drop(columns=[novo], inplace=True)

# limpeza final de colunas técnicas/vazias
resultado = drop_blank_technical_columns(resultado)

# --------------------- Preview e download ---------------------
st.subheader("Prévia do resultado (Plan Afiliados preenchida)")
st.dataframe(resultado.head(50), use_container_width=True)

total_pedidos = len(resultado)
faturados = resultado[AFF_OUT_STATUS_COL].astype(str).map(normalize_text).isin(FATURADO_VALUES).sum()
aplicados_regra = int(mask_regra.sum())
st.info(
    f"Pedidos na Plan Afiliados: {total_pedidos} | "
    f"Faturados (liberam comissão): {faturados} | "
    f"Pedidos forçados para '{status_forced_value}' pela regra: {aplicados_regra}"
)

output_name = "fechamento_afiliados.xlsx"
buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
    resultado.to_excel(writer, index=False, sheet_name="afiliados_preenchido")
    cons.to_excel(writer, index=False, sheet_name="base_consolidada")
buffer.seek(0)

st.download_button(
    "⬇️ Baixar planilha final (Excel)",
    data=buffer,
    file_name=output_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
