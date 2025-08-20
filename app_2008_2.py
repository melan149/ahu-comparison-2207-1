import streamlit as st
import pandas as pd
from PIL import Image
import plotly.express as px
import numpy as np

# -----------------------------
# Load data
# -----------------------------
@st.cache_data
def load_data():
    return pd.read_excel("Data_2025.xlsx", sheet_name="data", engine="openpyxl")

df = load_data()

# -----------------------------
# Helpers
# -----------------------------
def get_column_safe(df, name_options):
    for name in name_options:
        if name in df.columns:
            return name
    return None

unit_name_col = get_column_safe(df, ["Unit name", "Unit Name"])
region_col = get_column_safe(df, ["Region"])
year_col = get_column_safe(df, ["Year"])
quarter_col = get_column_safe(df, ["Quarter"])
recovery_col = get_column_safe(df, ["Recovery type", "Recovery Type"])
size_col = get_column_safe(df, ["Unit size", "Unit Size"])
brand_col = get_column_safe(df, ["Brand name", "Brand"])
logo_col = get_column_safe(df, ["Brand logo", "Brand Logo"])
unit_photo_col = get_column_safe(df, ["Unit photo", "Unit Photo"])
type_col = get_column_safe(df, ["Type"])
material_col = get_column_safe(df, ["Material"])

# -----------------------------
# Sidebar filter block per unit
# -----------------------------
def unit_filter_block(unit_idx, df):
    st.subheader(f"Select Unit {unit_idx}")

    year = st.selectbox(f"Year (Unit {unit_idx})", sorted(df[year_col].dropna().unique()), key=f"year_{unit_idx}")
    df_tmp = df[df[year_col] == year]

    quarter = st.selectbox(f"Quarter (Unit {unit_idx})", sorted(df_tmp[quarter_col].dropna().unique()), key=f"quarter_{unit_idx}")
    df_tmp = df_tmp[df_tmp[quarter_col] == quarter]

    region = st.selectbox(f"Region (Unit {unit_idx})", sorted(df_tmp[region_col].dropna().unique()), key=f"region_{unit_idx}")
    df_tmp = df_tmp[df_tmp[region_col] == region]

    brand = st.selectbox(f"Brand (Unit {unit_idx})", sorted(df_tmp[brand_col].dropna().unique()), key=f"brand_{unit_idx}")
    df_tmp = df_tmp[df_tmp[brand_col] == brand]

    unit = st.selectbox(f"Unit name (Unit {unit_idx})", sorted(df_tmp[unit_name_col].dropna().unique()), key=f"unit_{unit_idx}")
    df_tmp = df_tmp[df_tmp[unit_name_col] == unit]

    recovery = st.selectbox(f"Recovery type (Unit {unit_idx})", sorted(df_tmp[recovery_col].dropna().unique()), key=f"recovery_{unit_idx}")
    df_tmp = df_tmp[df_tmp[recovery_col] == recovery]

    size = st.selectbox(f"Unit size (Unit {unit_idx})", sorted(df_tmp[size_col].dropna().unique()), key=f"size_{unit_idx}")
    df_tmp = df_tmp[df_tmp[size_col] == size]

    # conditional type / material filter
    if recovery == "RRG" and type_col:
        types = sorted(df_tmp[type_col].dropna().unique())
        selected_type = st.selectbox(f"Rotary wheel type (Unit {unit_idx})", types, key=f"type_{unit_idx}")
        df_tmp = df_tmp[df_tmp[type_col] == selected_type]
    elif recovery in ["HEX", "PCR"] and material_col:
        materials = sorted(df_tmp[material_col].dropna().unique())
        selected_material = st.selectbox(f"PCR/HEX material (Unit {unit_idx})", materials, key=f"material_{unit_idx}")
        df_tmp = df_tmp[df_tmp[material_col] == selected_material]

    return df_tmp

# -----------------------------
# Sidebar main
# -----------------------------
with st.sidebar:
    n_units = st.slider("How many units to compare?", 2, 10, 2)
    st.markdown("---")
    filtered_units = [unit_filter_block(i, df) for i in range(1, n_units+1)]

# -----------------------------
# Main content: Comparison
# -----------------------------
st.title("Technical Data Comparison")

if all([not fu.empty for fu in filtered_units]):

    # --- Logos ---
    st.subheader("Brand Logos")
    cols = st.columns(n_units)
    for i, fu in enumerate(filtered_units):
        logo_path = str(fu[logo_col].iloc[0]) if logo_col in fu and pd.notna(fu[logo_col].iloc[0]) else None
        if logo_path:
            try:
                img = Image.open(f"images/{logo_path}")
                cols[i].image(img.resize((150, int(150 * img.height / img.width))), caption=f"{fu[brand_col].iloc[0]} Logo")
            except:
                cols[i].write("No logo")
        else:
            cols[i].write("No logo")

    # --- Unit Photos ---
    st.subheader("Unit Photos")
    cols = st.columns(n_units)
    for i, fu in enumerate(filtered_units):
        photo_path = str(fu[unit_photo_col].iloc[0]) if unit_photo_col in fu and pd.notna(fu[unit_photo_col].iloc[0]) else None
        if photo_path:
            try:
                img = Image.open(f"images/{photo_path}")
                cols[i].image(img.resize((250, int(250 * img.height / img.width))), caption=f"{fu[unit_name_col].iloc[0]} Photo")
            except:
                cols[i].write("No photo")
        else:
            cols[i].write("No photo")

    # --- General Data Table ---
    st.subheader("General Data")
    for col_name in df.columns:
        row_cols = st.columns([2] + [3]*n_units)
        row_cols[0].markdown(f"**{col_name}**")
        for i, fu in enumerate(filtered_units):
            val = fu[col_name].iloc[0] if col_name in fu.columns else "-"
            row_cols[i+1].write(val)

    # --- Example Chart (Unit size vs Region) ---
    if size_col and region_col:
        st.subheader("Example Chart: Unit size by Region")
        chart_data = []
        for i, fu in enumerate(filtered_units):
            chart_data.append({
                "Unit": f"Unit {i+1} - {fu[brand_col].iloc[0]}",
                "Size": fu[size_col].iloc[0],
                "Region": fu[region_col].iloc[0]
            })
        chart_df = pd.DataFrame(chart_data)
        fig = px.bar(chart_df, x="Unit", y="Size", color="Region", barmode="group")
        st.plotly_chart(fig, use_container_width=True)

    # --- CSV Export ---
    csv_data = []
    header_row = ["Parameter"] + [f"Unit {i+1}" for i in range(n_units)]
    csv_data.append(header_row)

    for col in df.columns:
        row = [col]
        for fu in filtered_units:
            row.append(fu[col].iloc[0] if col in fu.columns else "-")
        csv_data.append(row)

    csv_df = pd.DataFrame(csv_data)
    csv_string = csv_df.to_csv(index=False, header=False)
    st.download_button("Download Comparison CSV", csv_string, file_name="comparison.csv", mime="text/csv")

else:
    st.warning("At least one selected unit has no data. Please adjust selections.")
