import streamlit as st
import pandas as pd
from PIL import Image
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

# Load data
@st.cache_data
def load_data():
    # Assuming Data_2025.xlsx is in the same directory as app.py
    url = "Data_2025.xlsx"
    return pd.read_excel(url, sheet_name="data", engine='openpyxl')

df = load_data()

# Resolve potential column naming issues for robustness
def get_column_safe(df, name_options):
    for name in name_options:
        if name in df.columns:
            return name
    return None

unit_name_col = get_column_safe(df, ["Unit name", "Unit name", "Unit Name"])
region_col = get_column_safe(df, ["Region"])
year_col = get_column_safe(df, ["Year"])
quarter_col = get_column_safe(df, ["Quarter"])
recovery_col = get_column_safe(df, ["Recovery type", "Recovery Type", "Recovery_type"])
size_col = get_column_safe(df, ["Unit size", "Unit Size"])
brand_col = get_column_safe(df, ["Brand name", "Brand name", "Brand"])
logo_col = get_column_safe(df, ["Brand logo", "Brand logo", "Brand Logo"])
unit_photo_col = get_column_safe(df, ["Unit photo", "Unit photo", "Unit Photo", "Unit Photo Name"])
unit_size_quantity_col = get_column_safe(df, ["Unit size quantity", "Unit Size quantity"])
eurovent_model_box_col = get_column_safe(df, ["Eurovent Model Box"])

# Specific columns to trigger chart or section header display
internal_height_supply_filter_col = get_column_safe(df, ["Internal Height (Supply Filter) [mm]", "Internal Height (Supply Filter)", "Internal Height Supply Filter"])
unit_cross_section_area_supply_fan_col = get_column_safe(df, ["Unit cross section area (Supply Fan) [m2]", "Unit cross section area (Supply Fan)", "Unit cross section area Supply Fan"])
duct_connection_height_col = get_column_safe(df, ["Duct connection Height [mm]", "Duct connection Height", "Duct Connection Height"])
duct_connection_diameter_col = get_column_safe(df, ["Duct connection Diameter [mm]", "Duct connection Diameter", "Duct Connection Diameter"])

# New columns for dropdowns and chart
type_col = get_column_safe(df, ["Type"])
material_col = get_column_safe(df, ["Material"])
capacity_range1_col = get_column_safe(df, ["Capacity range1 [kW]", "Capacity range1", "Capacity Range1"])
capacity_range2_col = get_column_safe(df, ["Capacity range2 [kW]", "Capacity range2", "Capacity Range2"])
capacity_range3_col = get_column_safe(df, ["Capacity range3 [kW]", "Capacity range3", "Capacity Range3"])
heating_elements_type_col = get_column_safe(df, ["Heating elements type", "Heating Elements Type", "Heating_elements_type"])
impeller_efficiency_col = get_column_safe(df, ["Impeller efficiency at nominal airflow [%]", "Impeller efficiency at nominal airflow"])

sens_efficiency_nominal_rrg_col = get_column_safe(df, ["Sens. efficiency at nominal balanced airflows_RRG [%]", "Sens. efficiency at nominal balanced airflows [%]"])
sens_efficiency_opt_rrg_col = get_column_safe(df, ["Sens. efficiency at opt balanced airflows (ErP)_RRG [%]", "Sens. efficiency at opt balanced airflows (ErP) [%]"])
sens_efficiency_nominal_pcr_hex_col = get_column_safe(df, ["Sens. efficiency at nominal balanced airflows_PCR/HEX [%]", "Sens. efficiency at nominal balanced airflows [%].1"])
sens_efficiency_opt_pcr_hex_col = get_column_safe(df, ["Sens. efficiency at opt balanced airflows (ErP)_PCR/HEX [%]", "Sens. efficiency at opt balanced airflows (ErP) [%].1"])

metal_sheet_thickness_external_col = get_column_safe(df, ["Metal sheet thickness (External) [mm]", "Metal sheet thickness (External)"])
air_speed_filter_max_airflow_col = get_column_safe(df, ["Air speed on Filter at opt airflow (ErP) [m/s]", "Air speed on Filter at opt airflow (ErP)"])
final_pd_supply_col = get_column_safe(df, ["Final PD_Supply", "Final PD_typ1"])
final_pd_exhaust_col = get_column_safe(df, ["Final PD_Exhaust", "Final PD_typ2"])

duct_connection_width_col = get_column_safe(df, ["Duct connection Width [mm]", "Duct connection Width"])
water_heater_min_rows_col = get_column_safe(df, ["Water heater_min rows"])
water_cooler_min_rows_col = get_column_safe(df, ["Water cooler_min rows"])
dxh_min_rows_col = get_column_safe(df, ["DXH_min rows"])
filter_type_supply_col = get_column_safe(df, ["Filter type_Supply"])
filter_type_exhaust_col = get_column_safe(df, ["Filter type_Exhaust"])
silencer_casing_col = get_column_safe(df, ["Silencer casing"])
motor_type_col = get_column_safe(df, ["Motor type"])
supply_col = get_column_safe(df, ["Supply"])

# Header trigger columns
header_triggers_map = {
    get_column_safe(df, ["Eurovent Certificate"]): "Certification data",
    supply_col: "Available configurations",
    get_column_safe(df, ["Insulation material"]): "Casing",
    get_column_safe(df, ["Minimum airflow [CMH]"]): "Airflows",
    get_column_safe(df, ["Internal Width (Supply Filter) [mm]"]): "Overall dimensions",
    type_col: "Rotary wheel",
    sens_efficiency_nominal_pcr_hex_col: "PCR/HEX recovery exchanger",
    motor_type_col: "Fan section data",
    heating_elements_type_col: "Electrical heater",
    water_heater_min_rows_col: "Water heater",
    water_cooler_min_rows_col: "Water cooler",
    dxh_min_rows_col: "DX/DXH cooler",
    filter_type_supply_col: "Supply Filter",
    filter_type_exhaust_col: "Exhaust Filter",
    silencer_casing_col: "Silencer data"
}

electrical_heater_chart_trigger_col = heating_elements_type_col

# --- Chart coordinate column names ---
coord_col_pairs_1_5 = [(get_column_safe(df, [f"x{i}", f"X{i}"]), get_column_safe(df, [f"y{i}", f"Y{i}"])) for i in range(1, 6)]
coord_col_pairs_6_10 = [(get_column_safe(df, [f"x{i}", f"X{i}"]), get_column_safe(df, [f"y{i}", f"Y{i}"])) for i in range(6, 11)]
coord_col_pairs_11_15 = [(get_column_safe(df, [f"x{i}", f"X{i}"]), get_column_safe(df, [f"y{i}", f"Y{i}"])) for i in range(11, 16)]

st.title("Technical Data Comparison")

# --- Sidebar Filters ---
with st.sidebar:
    num_units = st.slider("Number of units for comparison", min_value=2, max_value=10, value=2)

    filtered_dfs = []
    selections = []

    for i in range(num_units):
        st.markdown("---")
        # Use an expander for each unit to create the hidden/collapsible menu
        with st.expander(f"Select Unit {i + 1}"):
            # Year filter
            available_years = sorted(df[year_col].dropna().unique())
            selected_year = st.selectbox(f"Year", available_years, key=f"year_{i}")
            df_filtered_by_year = df[df[year_col] == selected_year]

            # Quarter filter
            available_quarters = sorted(df_filtered_by_year[quarter_col].dropna().unique())
            selected_quarter = st.selectbox(f"Quarter", available_quarters, key=f"quarter_{i}")
            df_filtered_by_quarter = df_filtered_by_year[df_filtered_by_year[quarter_col] == selected_quarter]

            # Region filter
            available_regions = sorted(df_filtered_by_quarter[region_col].dropna().unique())
            selected_region = st.selectbox(f"Region", available_regions, key=f"region_{i}")
            df_filtered_by_region = df_filtered_by_quarter[df_filtered_by_quarter[region_col] == selected_region]

            # Brand filter
            available_brands = sorted(df_filtered_by_region[brand_col].dropna().unique())
            selected_brand = st.selectbox(f"Select Brand", available_brands, key=f"brand_{i}")
            df_filtered_by_brand = df_filtered_by_region[df_filtered_by_region[brand_col] == selected_brand]

            # Unit name filter
            available_units = sorted(df_filtered_by_brand[unit_name_col].dropna().unique())
            selected_unit = st.selectbox(f"Unit name", available_units, key=f"unit_{i}")
            df_filtered_by_unit = df_filtered_by_brand[df_filtered_by_brand[unit_name_col] == selected_unit]

            # Recovery type filter
            available_recovery_types = sorted(df_filtered_by_unit[recovery_col].dropna().unique())
            selected_recovery = st.selectbox(f"Recovery type", available_recovery_types, key=f"recovery_{i}")
            df_filtered_by_recovery = df_filtered_by_unit[df_filtered_by_unit[recovery_col] == selected_recovery]

            # Unit size filter
            available_sizes = sorted(df_filtered_by_recovery[size_col].dropna().unique())
            selected_size = st.selectbox(f"Unit size", available_sizes, key=f"size_{i}")
            df_filtered_by_size = df_filtered_by_recovery[df_filtered_by_recovery[size_col] == selected_size]

            # Conditional dropdowns
            selected_type = None
            selected_material = None
            df_temp_filtered = df_filtered_by_size

            if selected_recovery == "RRG" and type_col:
                available_types = sorted(df_filtered_by_size[type_col].dropna().unique())
                selected_type = st.selectbox(f"Rotary wheel type", available_types, key=f"type_{i}")
                df_temp_filtered = df_filtered_by_size[df_filtered_by_size[type_col] == selected_type]
            elif selected_recovery in ["HEX", "PCR"] and material_col:
                available_materials = sorted(df_filtered_by_size[material_col].dropna().unique())
                selected_material = st.selectbox(f"PCR/HEX lamels material", available_materials, key=f"material_{i}")
                df_temp_filtered = df_filtered_by_size[df_filtered_by_size[material_col] == selected_material]

        filtered_dfs.append(df_temp_filtered)
        selections.append({
            "year": selected_year, "quarter": selected_quarter, "region": selected_region,
            "brand": selected_brand, "unit": selected_unit, "size": selected_size,
            "recovery": selected_recovery, "type": selected_type, "material": selected_material
        })

    # --- CSV Download Button ---
    st.markdown("---")
    csv_header = ["Parameter"] + [f"{s['brand']} - {s['unit']} - {s['size']}" for s in selections]
    csv_data = [csv_header]

    raw_excluded_cols_base = [
        brand_col, logo_col, unit_photo_col, year_col, quarter_col, region_col,
        unit_name_col, recovery_col, size_col, type_col, material_col
    ]
    for pair in coord_col_pairs_1_5 + coord_col_pairs_6_10 + coord_col_pairs_11_15:
        if pair[0]: raw_excluded_cols_base.append(pair[0])
        if pair[1]: raw_excluded_cols_base.append(pair[1])

    unit_area_col_name = get_column_safe(df, ["Unit cross section area (Supply Filter) [m2]"])
    if unit_area_col_name: raw_excluded_cols_base.append(unit_area_col_name)

    excluded_cols_from_table = list(set([col for col in raw_excluded_cols_base if col is not None]))
    excluded_headers_from_display = set()

    selected_recoveries = [s['recovery'] for s in selections]

    if all(rec == "HEX" for rec in selected_recoveries):
        if get_column_safe(df, ["Wheel diameter [mm]"]): excluded_cols_from_table.append(get_column_safe(df, ["Wheel diameter [mm]"]))
        if get_column_safe(df, ["Distance between lamels [mm]"]): excluded_cols_from_table.append(get_column_safe(df, ["Distance between lamels [mm]"]))
        if type_col: excluded_cols_from_table.append(type_col)
        excluded_headers_from_display.add("Rotary wheel")
        if sens_efficiency_nominal_rrg_col: excluded_cols_from_table.append(sens_efficiency_nominal_rrg_col)
        if sens_efficiency_opt_rrg_col: excluded_cols_from_table.append(sens_efficiency_opt_rrg_col)

    if all(rec == "RRG" for rec in selected_recoveries):
        if material_col: excluded_cols_from_table.append(material_col)
        if sens_efficiency_nominal_pcr_hex_col: excluded_cols_from_table.append(sens_efficiency_nominal_pcr_hex_col)
        if sens_efficiency_opt_pcr_hex_col: excluded_cols_from_table.append(sens_efficiency_opt_pcr_hex_col)
        excluded_headers_from_display.add("PCR/HEX recovery exchanger")

    excluded_cols_from_table = list(set(excluded_cols_from_table))

    csv_displayed_headers = set()
    csv_data.append(["General data"] + [""] * num_units)

    for col in df.columns:
        triggered_header = header_triggers_map.get(col)
        if triggered_header and triggered_header not in csv_displayed_headers and triggered_header not in excluded_headers_from_display:
            csv_data.append([""] * (num_units + 1))
            csv_data.append([triggered_header] + [""] * num_units)
            csv_displayed_headers.add(triggered_header)

        if col not in excluded_cols_from_table:
            row = [col]
            for i in range(num_units):
                val = filtered_dfs[i][col].values[0] if not filtered_dfs[i].empty and col in filtered_dfs[i].columns else "-"
                row.append(val)
            csv_data.append(row)

    csv_df = pd.DataFrame(csv_data)
    csv_string = csv_df.to_csv(index=False, header=False)

    st.download_button(
        label="Download Comparison as CSV",
        data=csv_string,
        file_name="technical_data_comparison.csv",
        mime="text/csv",
        key="csv_download_sidebar"
    )

# --- Main Content Area ---

# --- Brand Logos ---
st.subheader("Brand Logos")
logo_cols = st.columns(num_units)
loaded_logos = []
for i in range(num_units):
    logo_path = None
    if not filtered_dfs[i].empty and logo_col in filtered_dfs[i].columns:
        logo_path_raw = filtered_dfs[i][logo_col].iloc[0]
        if pd.notna(logo_path_raw):
            logo_path = str(logo_path_raw)

    img = None
    if logo_path and logo_path.strip():
        try:
            img = Image.open(f"images/{logo_path}")
        except FileNotFoundError:
            st.warning(f"Logo not found for Unit {i+1}: images/{logo_path}")
        except Exception as e:
            st.warning(f"Error loading logo for Unit {i+1}: {e}")
    loaded_logos.append(img)

max_logo_height = max([img.height for img in loaded_logos if img] or [0])

for i in range(num_units):
    with logo_cols[i]:
        if loaded_logos[i]:
            img = loaded_logos[i]
            width = 150
            height = int(img.height * (width / img.width))
            if max_logo_height > 0:
                width = int(img.width * (max_logo_height / img.height))
                height = max_logo_height
            st.image(img.resize((width, height)), caption=f"Logo for {selections[i]['brand']}")
        else:
            st.write("No logo available.")

# --- Unit Photos ---
st.subheader("Unit Photo")
photo_cols = st.columns(num_units)
loaded_photos = []
for i in range(num_units):
    photo_path = None
    if not filtered_dfs[i].empty and unit_photo_col in filtered_dfs[i].columns:
        photo_path_raw = filtered_dfs[i][unit_photo_col].iloc[0]
        if pd.notna(photo_path_raw):
            photo_path = str(photo_path_raw)

    img = None
    if photo_path and photo_path.strip():
        try:
            img = Image.open(f"images/{photo_path}")
        except FileNotFoundError:
            st.warning(f"Unit photo not found for Unit {i+1}: images/{photo_path}")
        except Exception as e:
            st.warning(f"Error loading unit photo for Unit {i+1}: {e}")
    loaded_photos.append(img)

max_photo_height = max([img.height for img in loaded_photos if img] or [0])

for i in range(num_units):
    with photo_cols[i]:
        if loaded_photos[i]:
            img = loaded_photos[i]
            width = 250
            height = int(img.height * (width / img.width))
            if max_photo_height > 0:
                width = int(img.width * (max_photo_height / img.height))
                height = max_photo_height
            st.image(img.resize((width, height)), caption=f"{selections[i]['unit']} Photo")
        else:
            st.write("No unit photo available.")


# --- Comparison Table ---
if not all(df.empty for df in filtered_dfs):
    st.subheader("General data")

    # Table header
    # Adjust column widths to make the 'Parameter' column wider
    col_widths = [3] + [2] * num_units
    header_cols = st.columns(col_widths)
    with header_cols[0]:
        st.markdown("**Parameter**")
    for i in range(num_units):
        with header_cols[i + 1]:
            s = selections[i]
            st.markdown(f'<div style="text-align: center;">**{s["brand"]} - {s["unit"]} - {s["size"]}**</div>', unsafe_allow_html=True)

    displayed_headers = set()
    display_items_ordered = []

    for col_name in df.columns:
        if col_name == get_column_safe(df, ["Execution"]):
            if unit_size_quantity_col not in excluded_cols_from_table:
                display_items_ordered.append({"type": "row", "col": unit_size_quantity_col})
            display_items_ordered.append({"type": "chart", "name": "unit_area_chart"})
            continue

        header_title = header_triggers_map.get(col_name)
        if header_title and header_title not in displayed_headers and header_title not in excluded_headers_from_display:
            display_items_ordered.append({"type": "header", "title": header_title})
            displayed_headers.add(header_title)

        if col_name not in excluded_cols_from_table and col_name != unit_size_quantity_col:
            display_items_ordered.append({"type": "row", "col": col_name})

        if col_name == internal_height_supply_filter_col:
            display_items_ordered.append({"type": "chart", "name": "chart1"})
        elif col_name == unit_cross_section_area_supply_fan_col:
            display_items_ordered.append({"type": "chart", "name": "chart2"})
        elif col_name == duct_connection_height_col:
            display_items_ordered.append({"type": "chart", "name": "chart3"})
        elif col_name == electrical_heater_chart_trigger_col:
            if capacity_range1_col: display_items_ordered.append({"type": "row", "col": capacity_range1_col})
            if capacity_range2_col: display_items_ordered.append({"type": "row", "col": capacity_range2_col})
            if capacity_range3_col: display_items_ordered.append({"type": "row", "col": capacity_range3_col})
            display_items_ordered.append({"type": "chart", "name": "electrical_heater_chart"})

    colors = px.colors.qualitative.Plotly

    for item in display_items_ordered:
        if item["type"] == "header":
            st.markdown(f'<h4 style="text-align: center; font-size: 1.2em; margin: 1em 0;">{item["title"]}</h4>', unsafe_allow_html=True)
            # Re-add table headers for the new section with updated column widths
            h_cols = st.columns(col_widths)
            with h_cols[0]: st.markdown("**Parameter**")
            for i in range(num_units):
                with h_cols[i + 1]:
                    s = selections[i]
                    st.markdown(f'<div style="text-align: center;">**{s["brand"]} - {s["unit"]} - {s["size"]}**</div>', unsafe_allow_html=True)

        elif item["type"] == "row":
            col = item["col"]
            # Use updated column widths for the data rows
            row_cols = st.columns(col_widths)
            with row_cols[0]:
                st.write(f'<span style="font-family: sans-serif; font-size: 16px;">{col}</span>', unsafe_allow_html=True)
            for i in range(num_units):
                with row_cols[i + 1]:
                    val = filtered_dfs[i][col].values[0] if not filtered_dfs[i].empty and col in filtered_dfs[i].columns else "-"
                    color = colors[i % len(colors)]
                    st.markdown(f'<div style="text-align: center; font-family: sans-serif; font-size: 16px; color: {color};">{val}</div>', unsafe_allow_html=True)

        elif item["type"] == "chart":
            chart_name = item["name"]
            
            if chart_name == "unit_area_chart":
                chart_data_area = []
                color_map_area = {}
                unit_area_col_name = get_column_safe(df, ["Unit cross section area (Supply Filter) [m2]"])
                if unit_area_col_name and size_col:
                    for i in range(num_units):
                        s = selections[i]
                        label = f"Unit {i+1}: {s['brand']}"
                        color_map_area[label] = colors[i % len(colors)]
                        
                        df_chart_base = df[
                            (df[year_col] == s['year']) & (df[quarter_col] == s['quarter']) &
                            (df[region_col] == s['region']) & (df[brand_col] == s['brand']) &
                            (df[unit_name_col] == s['unit']) & (df[recovery_col] == s['recovery'])
                        ].copy()

                        if s['recovery'] == "RRG" and type_col and s['type']:
                            df_chart_base = df_chart_base[df_chart_base[type_col] == s['type']]
                        elif s['recovery'] in ["HEX", "PCR"] and material_col and s['material']:
                            df_chart_base = df_chart_base[df_chart_base[material_col] == s['material']]

                        if not df_chart_base.empty:
                            for _, row in df_chart_base.iterrows():
                                if pd.notna(row[unit_area_col_name]) and pd.notna(row[size_col]):
                                    chart_data_area.append({
                                        "Brand_UnitSize": f"{s['brand']} - Size {row[size_col]}",
                                        "Unit Cross Section Area (m²)": row[unit_area_col_name],
                                        "Unit Size": str(row[size_col]),
                                        "Selection_Label": label,
                                    })
                if chart_data_area:
                    chart_df_area = pd.DataFrame(chart_data_area)
                    fig_area = px.scatter(chart_df_area, x="Unit Cross Section Area (m²)", y="Brand_UnitSize",
                                          color="Selection_Label", text="Unit Size", 
                                          title='Unit Cross Section Area (Supply Filter) vs. Unit Size',
                                          color_discrete_map=color_map_area)
                    fig_area.update_traces(textposition='top center')
                    fig_area.update_layout(xaxis_title="Unit Cross Section Area (m²)", yaxis_title="Brand and Unit Size")
                    st.plotly_chart(fig_area, use_container_width=True)


            elif chart_name == "chart1":
                chart_data = []
                color_map = {}
                for i in range(num_units):
                    df_unit = filtered_dfs[i]
                    s_unit = selections[i]
                    label = f"Unit {i+1}: {s_unit['brand']} - {s_unit['size']}"
                    color_map[label] = colors[i % len(colors)]
                    can_plot = not df_unit.empty and all(pair[0] in df_unit.columns and pair[1] in df_unit.columns and pd.notna(df_unit[pair[0]].iloc[0]) and pd.notna(df_unit[pair[1]].iloc[0]) for pair in coord_col_pairs_1_5)
                    if can_plot:
                        for j, (x_name, y_name) in enumerate(coord_col_pairs_1_5):
                            chart_data.append({'X': df_unit[x_name].values[0], 'Y': df_unit[y_name].values[0], 'Label': label, 'Order': j})
                if chart_data:
                    chart_df = pd.DataFrame(chart_data).sort_values(by=['Label', 'Order'])
                    fig1 = px.line(chart_df, x="X", y="Y", color="Label", line_group="Label", markers=True, 
                                   title='Internal Cross Section area (Supply Filter)',
                                   color_discrete_map=color_map)
                    fig1.update_layout(xaxis_title="Width (mm)", yaxis_title="Height (mm)", legend_title_text="Selection")
                    fig1.update_yaxes(scaleanchor="x", scaleratio=1)
                    st.plotly_chart(fig1, use_container_width=True)

            elif chart_name == "chart2":
                chart_data = []
                color_map = {}
                for i in range(num_units):
                    df_unit = filtered_dfs[i]
                    s_unit = selections[i]
                    label = f"Unit {i+1}: {s_unit['brand']} - {s_unit['size']}"
                    color_map[label] = colors[i % len(colors)]
                    can_plot = not df_unit.empty and all(pair[0] in df_unit.columns and pair[1] in df_unit.columns and pd.notna(df_unit[pair[0]].iloc[0]) and pd.notna(df_unit[pair[1]].iloc[0]) for pair in coord_col_pairs_6_10)
                    if can_plot:
                        for j, (x_name, y_name) in enumerate(coord_col_pairs_6_10):
                            chart_data.append({'X': df_unit[x_name].values[0], 'Y': df_unit[y_name].values[0], 'Label': label, 'Order': j})
                if chart_data:
                    chart_df = pd.DataFrame(chart_data).sort_values(by=['Label', 'Order'])
                    fig2 = px.line(chart_df, x="X", y="Y", color="Label", line_group="Label", markers=True, 
                                   title='Internal Cross Section area (Supply Fan)',
                                   color_discrete_map=color_map)
                    fig2.update_layout(xaxis_title="Width (mm)", yaxis_title="Height (mm)", legend_title_text="Selection")
                    fig2.update_yaxes(scaleanchor="x", scaleratio=1)
                    st.plotly_chart(fig2, use_container_width=True)

            elif chart_name == "chart3":
                chart_data = []
                color_map = {}
                for i in range(num_units):
                    df_unit = filtered_dfs[i]
                    s_unit = selections[i]
                    label = f"Unit {i+1}: {s_unit['brand']} - {s_unit['size']}"
                    color_map[label] = colors[i % len(colors)]

                    is_rect = not df_unit.empty and any(pd.notna(df_unit[c].iloc[0]) and df_unit[c].iloc[0] != 0 for pair in coord_col_pairs_11_15 if pair[0] and pair[1] for c in pair)
                    is_circ = not df_unit.empty and duct_connection_diameter_col in df_unit.columns and pd.notna(df_unit[duct_connection_diameter_col].iloc[0]) and df_unit[duct_connection_diameter_col].iloc[0] > 0

                    if is_rect:
                        for j, (x_name, y_name) in enumerate(coord_col_pairs_11_15):
                             if x_name in df_unit.columns and y_name in df_unit.columns and pd.notna(df_unit[x_name].iloc[0]) and pd.notna(df_unit[y_name].iloc[0]):
                                chart_data.append({'X': df_unit[x_name].values[0], 'Y': df_unit[y_name].values[0], 'Label': label, 'Order': j})
                    elif is_circ:
                        diameter = df_unit[duct_connection_diameter_col].values[0]
                        radius = diameter / 2.0
                        theta = np.linspace(0, 2 * np.pi, 100)
                        for t in theta:
                            chart_data.append({'X': radius + radius * np.cos(t), 'Y': radius + radius * np.sin(t), 'Label': label, 'Order': 0})

                if chart_data:
                    chart_df = pd.DataFrame(chart_data).sort_values(by=['Label', 'Order'])
                    fig3 = px.line(chart_df, x="X", y="Y", color="Label", line_group="Label", markers=False, 
                                   title='Supply Duct connection, mm',
                                   color_discrete_map=color_map)
                    fig3.update_layout(xaxis_title="Width (mm)", yaxis_title="Height (mm)", legend_title_text="Selection")
                    fig3.update_yaxes(scaleanchor="x", scaleratio=1)
                    st.plotly_chart(fig3, use_container_width=True)

            elif chart_name == "electrical_heater_chart":
                chart_data = []
                color_map = {}
                for i in range(num_units):
                    df_unit = filtered_dfs[i]
                    s_unit = selections[i]
                    label = f"Unit {i+1}: {s_unit['brand']} - {s_unit['size']}"
                    color_map[label] = colors[i % len(colors)]
                    
                    if not df_unit.empty and all(c in df_unit.columns and pd.notna(df_unit[c].iloc[0]) for c in [capacity_range1_col, capacity_range2_col, capacity_range3_col]):
                        chart_data.append({"Capacity Range": "Range 1", "Value (kW)": df_unit[capacity_range1_col].values[0], "Selection": label})
                        chart_data.append({"Capacity Range": "Range 2", "Value (kW)": df_unit[capacity_range2_col].values[0], "Selection": label})
                        chart_data.append({"Capacity Range": "Range 3", "Value (kW)": df_unit[capacity_range3_col].values[0], "Selection": label})

                if chart_data:
                    chart_df = pd.DataFrame(chart_data)
                    fig_heater = px.bar(chart_df, x="Capacity Range", y="Value (kW)", color="Selection", barmode="group", 
                                        title='Electrical Heater Capacity (kW)',
                                        color_discrete_map=color_map)
                    fig_heater.update_layout(legend_title_text="Selection", yaxis_title="Capacity (kW)")
                    st.plotly_chart(fig_heater, use_container_width=True)

else:
    st.warning("Please make valid selections for all units to see a comparison.")
