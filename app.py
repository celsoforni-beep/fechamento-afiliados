import io
import datetime as dt

import streamlit as st
import pandas as pd

st.set_page_config(page_title="Fechamento Afiliados", layout="wide")
st.title("Fechamento mensal de comissões de afiliados")
st.caption("Fluxo: subir Plan Base (e-commerce) + Plan Afiliados (afiliados) → baixar Plan Afiliados preenchida.")

# ===================== CONFIG (AJUSTE SOMENTE SE MUDAREM OS CABEÇALHOS) =====================
# Plan Base (e-commerce)
ECOM_ORDER_COL = "Order2"
ECOM_STATUS_COL = "Status"
ECOM_FREIGHT_COL = "Shipping Value"     # Frete RATEADO por item -> SUM por pedido
ECOM_ORDER_VALUE_COL = "Total Value"    # Normalmente repetido por item -> MAX por pedido

# >>> IMPORTANTE: ajuste para o nome exato da coluna de data na sua Plan Base <<<
# Exemplos comuns: "Emission Date", "Order Date", "Created At", "Data Emissão"
ECOM_EMISSION_DATE_COL = "Emission Date"

# Plan Afiliados
AFF_ORDER_COL = "Order ID"

# Status que liberam comissão (normalize lower)
FATURADO_VALUES = {"faturado"}

# Colunas que serão preenchidas na Plan Afiliados (se não existirem, serão criadas)
AFF_OUT_STATUS_COL = "Status"
AFF_OUT_ORDER_VALUE_COL = "Valor do Pedidos"
AFF_OUT_FREIGHT_COL = "Valor de Frete"
AFF_OUT_NET_COL = "Valor S/frete"
# ==========================================================================================

# --------------------- Parâmetros de regra do fechamento ---------------------
st.subheader("Parâmetros do fechamento")

ref_month = st.date_input(
    "Mês de referência do fechamento (escolha qualquer dia dentro do mês)",
    value=dt.date.today().replace(day=1)
)

cutoff_day = st.number_input("Dia limite (corte)", min_value=1, max_value=31, value=20, step=1)

status_target_norm = st.text_input(
    "Status que será tratado como 'cancelado' quando ultrapassar o corte",
    value="preparando para entrega"
).strip().lower()

status_forced_value = st.text_input(
    "Status aplicado pela regra (valor que será reportado)",
    value="cancelado"
).strip()

cutoff_date = dt.date(ref_month.year, ref_month.month, int(cutoff_day))
st.caption(
    f"Regra: pedidos com **emissão até {cutoff_date}** e status **'{status_target_norm}'** "
    f"serão reportados como **'{status_forced_value}'**."
)

with st.expander("Configuração (para referência)", expanded=False):
    st.write(
        {
            "Plan Base - chave": ECOM_ORDER_COL,
            "Plan Base - status": ECOM_STATUS_COL,
            "Plan Base - frete (rateado)": ECOM_FREIGHT_COL,
            "Plan Base - total pedido": ECOM_ORDER_VALUE_COL,
            "Plan Base - data emissão": ECOM_EMISSION_DATE_COL,
            "Plan Afiliados - chave": AFF_ORDER_COL,
            "Status faturado": list(FATURADO_VALUES),
            "Corte": str(cutoff_date),
            "Status alvo (norm)": status_target_norm,
            "Status aplicado": status_forced_value,
        }
    )

# --------------------- Upload dos arquivos ---------------------
col1, col2 = st.columns(2)
with col1:
    ecom_file = st.file_uploader("📄 Suba a Plan Base (e-commerce) - .csv ou .xlsx", type=["csv", "xlsx"])
with col2:
    aff_file = st.file_uploader("📄 Suba a Plan Afiliados - .csv ou .xlsx", type=["csv", "xlsx"])


def read_any(file) -> pd.DataFrame:
    if file.name.lower().endswith(".csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)


def norm_str(x) -> str:
    return str(x).strip().lower()


def ensure_cols(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    for c in cols:
        if c not in df.columns:
            df[c] = pd.NA
    return df


if ecom_file and aff_file:
    try:
        ecom = read_any(ecom_file)
        aff = read_any(aff_file)
    except Exception as e:
        st.error(f"Erro ao ler os arquivos: {e}")
        st.stop()

    # --------------------- Validação ---------------------
    needed_ecom = {
        ECOM_ORDER_COL,
        ECOM_STATUS_COL,
        ECOM_FREIGHT_COL,
        ECOM_ORDER_VALUE_COL,
        ECOM_EMISSION_DATE_COL,
    }
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
        st.error(f"Faltam colunas na Plan Afiliados: {sorted(list(missing_aff))}")
        st.stop()

    # --------------------- Normalização e parsing ---------------------
    ecom[ECOM_ORDER_COL] = ecom[ECOM_ORDER_COL].astype(str).str.strip()
    aff[AFF_ORDER_COL] = aff[AFF_ORDER_COL].astype(str).str.strip()

    ecom[ECOM_FREIGHT_COL] = pd.to_numeric(ecom[ECOM_FREIGHT_COL], errors="coerce").fillna(0)
    ecom[ECOM_ORDER_VALUE_COL] = pd.to_numeric(ecom[ECOM_ORDER_VALUE_COL], errors="coerce").fillna(0)

    # Parse data (pt-BR geralmente dayfirst=True)
    ecom[ECOM_EMISSION_DATE_COL] = pd.to_datetime(
        ecom[ECOM_EMISSION_DATE_COL],
        errors="coerce",
        dayfirst=True
    )

    # --------------------- Consolidação por pedido ---------------------
    # Frete rateado -> soma
    freight_by_order = (
        ecom.groupby(ECOM_ORDER_COL, as_index=False)[ECOM_FREIGHT_COL]
        .sum()
        .rename(columns={ECOM_FREIGHT_COL: "frete_consolidado"})
    )

    # Total do pedido -> max (evita duplicação)
    value_by_order = (
        ecom.groupby(ECOM_ORDER_COL, as_index=False)[ECOM_ORDER_VALUE_COL]
        .max()
        .rename(columns={ECOM_ORDER_VALUE_COL: "valor_pedido_consolidado"})
    )

    # Data de emissão -> min (primeira emissão do pedido)
    emissao_by_order = (
        ecom.groupby(ECOM_ORDER_COL, as_index=False)[ECOM_EMISSION_DATE_COL]
        .min()
        .rename(columns={ECOM_EMISSION_DATE_COL: "data_emissao_consolidada"})
    )

    # Status -> primeiro não nulo
    status_by_order = (
        ecom[[ECOM_ORDER_COL, ECOM_STATUS_COL]]
        .dropna(subset=[ECOM_STATUS_COL])
        .groupby(ECOM_ORDER_COL, as_index=False)[ECOM_STATUS_COL]
        .first()
        .rename(columns={ECOM_STATUS_COL: "status_pedido"})
    )

    # Tabela consolidada
    cons = status_by_order.merge(freight_by_order, how="outer", on=ECOM_ORDER_COL)
    cons = cons.merge(value_by_order, how="outer", on=ECOM_ORDER_COL)
    cons = cons.merge(emissao_by_order, how="outer", on=ECOM_ORDER_COL)

    cons["frete_consolidado"] = cons["frete_consolidado"].fillna(0)
    cons["valor_pedido_consolidado"] = cons["valor_pedido_consolidado"].fillna(0)

    # --------------------- Regra do corte: preparando -> cancelado ---------------------
    # Garante status preenchido para não quebrar normalização
    cons["status_pedido"] = cons["status_pedido"].fillna("Pedido não encontrado na base")

    cutoff_ts = pd.Timestamp(cutoff_date)

    cons["status_norm"] = cons["status_pedido"].map(norm_str)

    mask_regra = (
        (cons["status_norm"] == status_target_norm) &
        (cons["data_emissao_consolidada"].notna()) &
        (cons["data_emissao_consolidada"] <= cutoff_ts)
    )

    cons.loc[mask_regra, "status_pedido"] = status_forced_value
    cons.drop(columns=["status_norm"], inplace=True, errors="ignore")

    # --------------------- Merge na planilha de Afiliados ---------------------
    resultado = ensure_cols(
        aff.copy(),
        [AFF_OUT_STATUS_COL, AFF_OUT_ORDER_VALUE_COL, AFF_OUT_FREIGHT_COL, AFF_OUT_NET_COL],
    )

    tmp = cons.rename(columns={ECOM_ORDER_COL: AFF_ORDER_COL}).copy()
    tmp["status_pedido"] = tmp["status_pedido"].fillna("Pedido não encontrado na base")
    tmp["status_norm"] = tmp["status_pedido"].map(norm_str)

    # Valor sem frete (base comissão) só para faturado
    tmp[AFF_OUT_NET_COL] = (tmp["valor_pedido_consolidado"] - tmp["frete_consolidado"]).clip(lower=0)
    tmp.loc[~tmp["status_norm"].isin(FATURADO_VALUES), AFF_OUT_NET_COL] = 0

    # Renomeia para bater com a Plan Afiliados
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

    # Sobrescreve colunas alvo
    for col in [AFF_OUT_STATUS_COL, AFF_OUT_ORDER_VALUE_COL, AFF_OUT_FREIGHT_COL, AFF_OUT_NET_COL]:
        novo = f"{col}_novo"
        if novo in resultado.columns:
            resultado[col] = resultado[novo]
            resultado.drop(columns=[novo], inplace=True)

    # Limpeza de colunas técnicas (evita colunas em branco extras)
    resultado = resultado.loc[:, ~resultado.columns.astype(str).str.contains(r"(_novo|_x|_y)$", regex=True)]

    # --------------------- Preview e métricas ---------------------
    st.subheader("Prévia do resultado (Plan Afiliados preenchida)")
    st.dataframe(resultado.head(50), use_container_width=True)

    total_pedidos_aff = len(resultado)
    pedidos_com_status = resultado[AFF_OUT_STATUS_COL].notna().sum()

    faturados = resultado[AFF_OUT_STATUS_COL].astype(str).str.strip().str.lower().isin(FATURADO_VALUES).sum()
    aplicados_regra = tmp.loc[tmp[AFF_OUT_STATUS_COL].astype(str).str.strip().str.lower() == status_forced_value.strip().lower()].shape[0]

    st.info(
        f"Pedidos na Plan Afiliados: {total_pedidos_aff} | "
        f"Pedidos com status preenchido: {pedidos_com_status} | "
        f"Faturados (liberam comissão): {faturados}"
    )

    # --------------------- Export Excel ---------------------
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

else:
    st.warning("Suba os dois arquivos para processar e gerar o fechamento.")
