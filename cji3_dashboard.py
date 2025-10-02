"""
Interactive CJI3 DRE Dashboard
--------------------------------

This Streamlit application allows users to upload an Excel export
from SAP's CJI3 transaction (or CJ13) and interactively generate
a Demonstration of Results (DRE) along with key performance
indicators (KPIs). Users can map the relevant columns from their
spreadsheet to the expected fields, filter by date (month and
year) and by object (WBS/PEP), choose the currency (BRL or EUR)
and view aggregated financial metrics.

Usage:
    streamlit run cji3_dashboard.py

Dependencies:
    - streamlit
    - pandas
    - numpy

The application does not rely on any fixed column names. After
uploading a file, the user must map each required field using
dropdown selectors. The core classification logic is based on
commonly used SAP cost element prefixes (RA, RCA, RCC, RCE,
RCJ). Deduction categories (PIS, COFINS, ICMS, ISS) are
identified by keywords in the description fields. Credits and
COGS adjustments follow the D/C (debit/credit) indicator.

Author: OpenAI ChatGPT
"""

import re
import pandas as pd
import numpy as np
import streamlit as st
from typing import Dict, List, Tuple


def normalize_text(text: str) -> str:
    """Normalize text by uppercasing and removing accents and extra spaces."""
    if pd.isna(text):
        return ""
    # Remove accents by encoding to ASCII and ignoring errors
    normalized = (
        str(text)
        .upper()
        .replace("Á", "A")
        .replace("Â", "A")
        .replace("Ã", "A")
        .replace("À", "A")
        .replace("É", "E")
        .replace("Ê", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Ô", "O")
        .replace("Õ", "O")
        .replace("Ú", "U")
        .replace("Ç", "C")
    )
    # Replace multiple spaces with single space
    normalized = re.sub(r"\s+", " ", normalized)
    return normalized.strip()


def classify_transaction(
    class_code: str,
    dc_indicator: str,
    description: str,
    value: float,
) -> Dict[str, float]:
    """
    Classify a single transaction into DRE categories based on the
    cost element prefix, debit/credit indicator and description.

    Parameters
    ----------
    class_code : str
        The cost element or class code (e.g. RAA1TZZS33).
    dc_indicator : str
        Debit/Credit indicator ('D' for debit, 'C' for credit).
    description : str
        Normalized description text from various descriptive fields.
    value : float
        Transaction value to accumulate (absolute values will be used
        and sign will be handled here).

    Returns
    -------
    Dict[str, float]
        A dictionary mapping DRE category names to the contribution
        of this transaction. Unused categories will have zero values.
    """
    # Initialize result dictionary
    categories = {
        "Receita Bruta": 0.0,
        "Dedução": 0.0,
        "PIS": 0.0,
        "COFINS": 0.0,
        "ICMS": 0.0,
        "ISS": 0.0,
        "COGS": 0.0,
        "Despesas Diretas": 0.0,
        "Créditos Tributários": 0.0,
    }

    if not class_code:
        return categories

    class_prefix = class_code[:2].upper()
    detailed_prefix = class_code[:3].upper()

    # Normalize description for keyword detection
    desc = normalize_text(description)

    # Handle revenue and deductions (RA...)
    if class_prefix == "RA":
        if dc_indicator.upper() == "C":
            # Credit: revenue
            # Check if description indicates sale of goods to override sign
            if "VENDA DE BENS" in desc or "OUTRAS RECEITAS VENDA DE BENS" in desc:
                categories["Receita Bruta"] += abs(value)
            else:
                categories["Receita Bruta"] += abs(value)
        elif dc_indicator.upper() == "D":
            # Debit: deduction
            # Identify specific taxes
            if "PIS" in desc:
                categories["PIS"] += abs(value)
            elif "COFINS" in desc:
                categories["COFINS"] += abs(value)
            elif "ICMS" in desc:
                categories["ICMS"] += abs(value)
            elif "ISS" in desc:
                categories["ISS"] += abs(value)
            else:
                categories["Dedução"] += abs(value)

    # Handle cost of goods sold and credits (RCA...)
    elif detailed_prefix == "RCA":
        if dc_indicator.upper() == "D":
            # Debit: cost
            categories["COGS"] += abs(value)
        elif dc_indicator.upper() == "C":
            # Credit: reduction of cost and tax credit
            categories["COGS"] -= abs(value)
            categories["Créditos Tributários"] += abs(value)

    # Handle operating expenses (RCC, RCE, RCJ)
    elif class_code.upper().startswith("RCC") or class_code.upper().startswith("RCE") or class_code.upper().startswith("RCJ"):
        if dc_indicator.upper() == "D":
            categories["Despesas Diretas"] += abs(value)
        elif dc_indicator.upper() == "C":
            categories["Despesas Diretas"] -= abs(value)

    # Other categories like RF, RI, RCZ, RZZ are ignored in EBITDA but could be added if needed

    return categories


def aggregate_dre(
    df: pd.DataFrame,
    cols: Dict[str, str],
    currency: str,
    filter_months: List[int],
    filter_years: List[int],
    filter_objects: List[str],
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Aggregate the DRE metrics by object based on filtered data.

    Parameters
    ----------
    df : pd.DataFrame
        Raw transaction data from CJI3.
    cols : dict
        Mapping for required columns: keys expected are
        'object', 'class', 'value_brl', 'value_euro', 'dc', 'desc', 'date'.
    currency : str
        Desired currency for aggregation ('BRL' or 'EUR').
    filter_months : list of int
        Months to include in the filter (1-12). If empty, include all.
    filter_years : list of int
        Years to include in the filter. If empty, include all.
    filter_objects : list of str
        Object codes to include. If empty, include all.

    Returns
    -------
    Tuple[pd.DataFrame, pd.DataFrame]
        DataFrame of aggregated DRE metrics per object and
        DataFrame of KPI summary across selected objects.
    """
    # Ensure date column is datetime
    df[cols['date']] = pd.to_datetime(df[cols['date']], errors='coerce')

    # Filter by month and year if provided
    if filter_months:
        df = df[df[cols['date']].dt.month.isin(filter_months)]
    if filter_years:
        df = df[df[cols['date']].dt.year.isin(filter_years)]
    # Filter by objects if provided
    if filter_objects:
        df = df[df[cols['object']].astype(str).isin(filter_objects)]

    # Select value column based on currency
    value_col = cols['value_brl'] if currency == 'BRL' else cols['value_euro']

    # Replace NaN values with zero in numeric columns
    df[value_col] = pd.to_numeric(df[value_col], errors='coerce').fillna(0.0)

    # Normalize description
    df['__norm_desc'] = df[cols['desc']].apply(normalize_text)

    # Prepare an empty list for aggregated results
    results = []

    # Group by object and aggregate
    for obj_code, group in df.groupby(df[cols['object']].astype(str)):
        # Initialize totals for this object
        totals = {
            'Object': obj_code,
            'Receita Bruta': 0.0,
            'Dedução de Receita': 0.0,
            'PIS': 0.0,
            'COFINS': 0.0,
            'ICMS': 0.0,
            'ISS': 0.0,
            'Receita Líquida': 0.0,
            'COGS': 0.0,
            'Margem Bruta': 0.0,
            'Despesas Diretas': 0.0,
            'Créditos Tributários': 0.0,
            'EBITDA': 0.0,
        }
        # Classify each transaction and accumulate
        for _, row in group.iterrows():
            categories = classify_transaction(
                class_code=str(row[cols['class']]),
                dc_indicator=str(row[cols['dc']]),
                description=row['__norm_desc'],
                value=float(row[value_col]),
            )
            for key, val in categories.items():
                if key == 'Dedução':
                    totals['Dedução de Receita'] += val
                else:
                    totals[key] += val

        # Compute derived metrics
        totals['Receita Líquida'] = totals['Receita Bruta'] - (
            totals['Dedução de Receita'] + totals['PIS'] + totals['COFINS'] + totals['ICMS'] + totals['ISS']
        )
        totals['Margem Bruta'] = totals['Receita Líquida'] - totals['COGS']
        totals['EBITDA'] = totals['Margem Bruta'] - totals['Despesas Diretas'] + totals['Créditos Tributários']
        results.append(totals)

    result_df = pd.DataFrame(results)

    # Create KPI summary across selected objects
    if not result_df.empty:
        summary = pd.DataFrame({
            'Receita Bruta': [result_df['Receita Bruta'].sum()],
            'Receita Líquida': [result_df['Receita Líquida'].sum()],
            'COGS': [result_df['COGS'].sum()],
            'Despesas Diretas': [result_df['Despesas Diretas'].sum()],
            'Créditos Tributários': [result_df['Créditos Tributários'].sum()],
            'EBITDA': [result_df['EBITDA'].sum()],
        })
    else:
        summary = pd.DataFrame(columns=['Receita Bruta','Receita Líquida','COGS','Despesas Diretas','Créditos Tributários','EBITDA'])

    return result_df, summary


def main() -> None:
    """Main function for Streamlit app."""
    st.set_page_config(page_title="CJI3 DRE Dashboard", layout="wide")
    st.title("CJI3 DRE Dashboard")
    st.write(
        """
        Faça upload do relatório exportado da transação CJI3 do SAP para gerar a
        Demonstração do Resultado (DRE) e KPIs por Objeto (WBS/PEP). O
        aplicativo suporta análise em BRL (coluna M) e EUR (coluna O).
        """
    )

    uploaded_file = st.file_uploader(
        "Selecione o arquivo Excel da CJI3", type=["xlsx", "xls"]
    )
    if uploaded_file is None:
        st.info("Aguardando upload do arquivo...")
        st.stop()

    # Read the Excel file
    try:
        raw_df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
        st.stop()

    st.subheader("Mapeamento de Colunas")
    st.write(
        "Selecione a coluna correspondente a cada campo obrigatório. Isso permite\n"
        "que o aplicativo funcione com diferentes versões da CJI3."
    )

    columns = list(raw_df.columns)

    # Column mapping selectors
    col1, col2, col3 = st.columns(3)
    with col1:
        object_col = st.selectbox("Coluna do Objeto (WBS/PEP)", options=columns)
        value_brl_col = st.selectbox("Coluna de Valor em BRL (M)", options=columns)
        value_euro_col = st.selectbox("Coluna de Valor em EUR (O)", options=columns)
    with col2:
        class_col = st.selectbox("Coluna da Classe de Custo (J)", options=columns)
        dc_col = st.selectbox("Coluna D/C (indicador débito/crédito)", options=columns)
    with col3:
        desc_col = st.selectbox(
            "Coluna da Descrição (pode ser combinação de descrições)", options=columns
        )
        date_col = st.selectbox(
            "Coluna de Data (Data de lançamento)", options=columns
        )

    col_mapping = {
        'object': object_col,
        'class': class_col,
        'value_brl': value_brl_col,
        'value_euro': value_euro_col,
        'dc': dc_col,
        'desc': desc_col,
        'date': date_col,
    }

    # Prepare filters based on date and object
    # Ensure date parsing
    df_temp = raw_df.copy()
    df_temp[date_col] = pd.to_datetime(df_temp[date_col], errors='coerce')

    available_months = sorted(df_temp[date_col].dropna().dt.month.unique())
    available_years = sorted(df_temp[date_col].dropna().dt.year.unique())
    available_objects = sorted(df_temp[object_col].astype(str).unique())

    st.subheader("Filtros de Data e Objeto")
    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        selected_months = st.multiselect(
            "Selecione o(s) mês(es)",
            options=[(i, __import__('datetime').date(1900, i, 1).strftime('%b')) for i in available_months],
            format_func=lambda x: x[1],
        )
        month_list = [m[0] for m in selected_months]
    with col_f2:
        selected_years = st.multiselect(
            "Selecione o(s) ano(s)", options=available_years
        )
    with col_f3:
        selected_objects = st.multiselect(
            "Selecione o(s) objeto(s)", options=available_objects
        )

    # Currency selection
    st.subheader("Moeda de Análise")
    currency = st.radio(
        "Selecione a moeda", options=["BRL", "EUR"], index=0, horizontal=True
    )

    # Process data and display results
    if st.button("Gerar DRE e KPIs"):
        with st.spinner("Processando dados..."):
            dre_df, summary_df = aggregate_dre(
                raw_df,
                cols=col_mapping,
                currency=currency,
                filter_months=month_list,
                filter_years=selected_years,
                filter_objects=selected_objects,
            )
        st.success("DRE e KPIs gerados com sucesso!")

        # Display aggregated DRE per object
        st.subheader(f"DRE por Objeto ({currency})")
        if dre_df.empty:
            st.info("Nenhum dado encontrado para os filtros selecionados.")
        else:
            # Format values as currency
            formatted_dre = dre_df.copy()
            currency_symbol = "R$" if currency == "BRL" else "€"
            value_cols = [
                'Receita Bruta', 'Dedução de Receita', 'PIS', 'COFINS', 'ICMS', 'ISS',
                'Receita Líquida', 'COGS', 'Margem Bruta', 'Despesas Diretas',
                'Créditos Tributários', 'EBITDA'
            ]
            for col in value_cols:
                formatted_dre[col] = formatted_dre[col].apply(
                    lambda x: f"{currency_symbol} {x:,.2f}".replace(",", "_temp_").replace(".", ",").replace("_temp_", ".")
                )
            st.dataframe(formatted_dre)

        # Display KPI summary
        st.subheader(f"KPIs Consolidados ({currency})")
        if summary_df.empty:
            st.info("Sem métricas para exibir.")
        else:
            formatted_summary = summary_df.copy()
            for col in formatted_summary.columns:
                formatted_summary[col] = formatted_summary[col].apply(
                    lambda x: f"{currency_symbol} {x:,.2f}".replace(",", "_temp_").replace(".", ",").replace("_temp_", ".")
                )
            st.dataframe(formatted_summary)

    st.caption(
        "As fórmulas seguem as regras definidas: RA → Receita/Deduções, RCA → COGS/Créditos, "
        "RCC/RCE/RCJ → Despesas. Os tributos (PIS, COFINS, ICMS, ISS) são identificados "
        "por palavras-chave na descrição."
    )


if __name__ == "__main__":
    main()