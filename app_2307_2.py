import streamlit as st
import pandas as pd
from PIL import Image
import plotly.express as px # Import plotly for charting
import plotly.graph_objects as go # Import graph objects for more control if needed
import numpy as np # For generating circle points

# Load data
@st.cache_data
def load_data():
    # Assuming Data_2025.xlsx is in the same directory as app.py
    url = "Data_2025.xlsx" # This path needs to be correct for your Streamlit environment
    return pd.read_excel(url, sheet_name="data", engine='openpyxl')

df = load_data()

# Initialize chart display flags at the beginning of the script
chart1_displayed = False
chart2_displayed = False
chart3_displayed = False
electrical_heater_chart_displayed = False
unit_area_chart_displayed = False # New flag for the new chart

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
unit_photo_col = get_column_safe(df, ["Unit photo", "Unit photo", "Unit Photo", "Unit Photo Name"]) # Added common names for unit photo column
unit_size_quantity_col = get_column_safe(df, ["Unit size quantity", "Unit Size quantity"]) # Explicitly define for relocation
eurovent_model_box_col = get_column_safe(df, ["Eurovent Model Box"]) # New column

# Specific columns to trigger chart or section header display
internal_height_supply_filter_col = get_column_safe(df, ["Internal Height (Supply Filter) [mm]", "Internal Height (Supply Filter)", "Internal Height Supply Filter"])
unit_cross_section_area_supply_fan_col = get_column_safe(df, ["Unit cross section area (Supply Fan) [m2]", "Unit cross section area (Supply Fan)", "Unit cross section area Supply Fan"])
duct_connection_height_col = get_column_safe(df, ["Duct connection Height [mm]", "Duct connection Height", "Duct Connection Height"])
# Column for circle diameter
duct_connection_diameter_col = get_column_safe(df, ["Duct connection Diameter [mm]", "Duct connection Diameter", "Duct Connection Diameter"])

# New columns for dropdowns and chart
type_col = get_column_safe(df, ["Type"]) # This will be "Rotary wheel type"
material_col = get_column_safe(df, ["Material"]) # This will be "PCR/HEX lamels material"
capacity_range1_col = get_column_safe(df, ["Capacity range1 [kW]", "Capacity range1", "Capacity Range1"])
capacity_range2_col = get_column_safe(df, ["Capacity range2 [kW]", "Capacity range2", "Capacity Range2"])
capacity_range3_col = get_column_safe(df, ["Capacity range3 [kW]", "Capacity range3", "Capacity Range3"])
heating_elements_type_col = get_column_safe(df, ["Heating elements type", "Heating Elements Type", "Heating_elements_type"])
impeller_efficiency_col = get_column_safe(df, ["Impeller efficiency at nominal airflow [%]", "Impeller efficiency at nominal airflow"])

# Updated column names based on user's latest input for efficiency
sens_efficiency_nominal_rrg_col = get_column_safe(df, ["Sens. efficiency at nominal balanced airflows_RRG [%]", "Sens. efficiency at nominal balanced airflows [%]"])
sens_efficiency_opt_rrg_col = get_column_safe(df, ["Sens. efficiency at opt balanced airflows (ErP)_RRG [%]", "Sens. efficiency at opt balanced airflows (ErP) [%]"])
sens_efficiency_nominal_pcr_hex_col = get_column_safe(df, ["Sens. efficiency at nominal balanced airflows_PCR/HEX [%]", "Sens. efficiency at nominal balanced airflows [%].1"])
sens_efficiency_opt_pcr_hex_col = get_column_safe(df, ["Sens. efficiency at opt balanced airflows (ErP)_PCR/HEX [%]", "Sens. efficiency at opt balanced airflows (ErP) [%].1"])

metal_sheet_thickness_external_col = get_column_safe(df, ["Metal sheet thickness (External) [mm]", "Metal sheet thickness (External)"])
air_speed_filter_max_airflow_col = get_column_safe(df, ["Air speed on Filter at opt airflow (ErP) [m/s]", "Air speed on Filter at opt airflow (ErP)"])
final_pd_supply_col = get_column_safe(df, ["Final PD_Supply", "Final PD_typ1"]) # Updated name
final_pd_exhaust_col = get_column_safe(df, ["Final PD_Exhaust", "Final PD_typ2"])

# Columns for Duct Connection Width
duct_connection_width_col = get_column_safe(df, ["Duct connection Width [mm]", "Duct connection Width"])
water_cooler_min_rows_col = get_column_safe(df, ["Water cooler_min rows"])
dxh_min_rows_col = get_column_safe(df, ["DXH_min rows"])
filter_type_supply_col = get_column_safe(df, ["Filter type_Supply"])
filter_type_exhaust_col = get_column_safe(df, ["Filter type_Exhaust"])
silencer_casing_col = get_column_safe(df, ["Silencer casing"])
motor_type_col = get_column_safe(df, ["Motor type"])
supply_col = get_column_safe(df, ["Supply"])


# Header trigger columns (mapping first column in a section to its header)
header_triggers_map = {
    get_column_safe(df, ["Eurovent Certificate"]): "Certification data",
    supply_col: "Available configurations", # Changed trigger to 'Supply'
    get_column_safe(df, ["Insulation material"]): "Casing", # Changed trigger to 'Insulation material'
    get_column_safe(df, ["Minimum airflow [CMH]"]): "Airflows",
    get_column_safe(df, ["Internal Width (Supply Filter) [mm]"]): "Overall dimensions", # Changed trigger to 'Internal Width (Supply Filter) [mm]'
    type_col: "Rotary wheel",
    sens_efficiency_nominal_pcr_hex_col: "PCR/HEX recovery exchanger",
    motor_type_col: "Fan section data", # Changed trigger to 'Motor type'
    heating_elements_type_col: "Electrical heater",
    capacity_range3_col: "Water heater",
    water_cooler_min_rows_col: "Water cooler", # Changed trigger to 'Water cooler_min rows'
    dxh_min_rows_col: "DX/DXH cooler", # Changed trigger to 'DXH_min rows'
    filter_type_supply_col: "Supply Filter", # Changed trigger to 'Filter type_Supply'
    filter_type_exhaust_col: "Exhaust Filter", # Changed trigger to 'Filter type_Exhaust'
    silencer_casing_col: "Silencer data" # Changed trigger to 'Silencer casing'
}

# Chart for Electrical Heater Capacity: after the row containing heating_elements_type_col
electrical_heater_chart_trigger_col = heating_elements_type_col


# --- Chart 1 specific coordinate column names (X1-X5, Y1-Y5) ---
coord_col_pairs_1_5 = []
for i in range(1, 6):
    x_col_name = get_column_safe(df, [f"x{i}", f"X{i}", f"X{i}_coord", f"x{i}_coord"])
    y_col_name = get_column_safe(df, [f"y{i}", f"Y{i}", f"Y{i}_coord", f"y{i}_coord"])
    if x_col_name and y_col_name:
        coord_col_pairs_1_5.append((x_col_name, y_col_name))

# --- Chart 2 specific coordinate column names (X6-X10, Y6-Y10) ---
coord_col_pairs_6_10 = []
for i in range(6, 11): # For X6, Y6 to X10, Y10
    x_col_name = get_column_safe(df, [f"x{i}", f"X{i}", f"X{i}_coord", f"x{i}_coord"])
    y_col_name = get_column_safe(df, [f"y{i}", f"Y{i}", f"Y{i}_coord", f"y{i}_coord"])
    if x_col_name and y_col_name:
        coord_col_pairs_6_10.append((x_col_name, y_col_name))

# --- Chart 3 specific coordinate column names (X11-X15, Y11-Y15) ---
coord_col_pairs_11_15 = []
for i in range(11, 16): # For X11, Y11 to X15, Y15
    x_col_name = get_column_safe(df, [f"x{i}", f"X{i}", f"X{i}_coord", f"x{i}_coord"])
    y_col_name = get_column_safe(df, [f"y{i}", f"Y{i}", f"Y{i}_coord", f"y{i}_coord"])
    if x_col_name and y_col_name:
        coord_col_pairs_11_15.append((x_col_name, y_col_name))


# Main layout filters for the comparison interface
st.title("Technical Data Comparison")

# --- Left Column: Filters (Now in Sidebar) ---
with st.sidebar:
    st.header("Select First Unit for Comparison") # Title for the sidebar filter section

    # Year filter
    available_years1 = sorted(df[year_col].dropna().unique())
    selected_year1 = st.selectbox("Year (Left)", available_years1, key="year1_sidebar") # Added _sidebar key
    df_filtered_by_year1 = df[df[year_col] == selected_year1]

    # Quarter filter
    available_quarters1 = sorted(df_filtered_by_year1[quarter_col].dropna().unique())
    selected_quarter1 = st.selectbox("Quarter (Left)", available_quarters1, key="quarter1_sidebar")
    df_filtered_by_quarter1 = df_filtered_by_year1[df_filtered_by_year1[quarter_col] == selected_quarter1]

    # Region filter
    available_regions1 = sorted(df_filtered_by_quarter1[region_col].dropna().unique())
    selected_region1 = st.selectbox("Region (Left)", available_regions1, key="region1_sidebar")
    df_filtered_by_region1 = df_filtered_by_quarter1[df_filtered_by_quarter1[region_col] == selected_region1]

    # Brand filter
    available_brands1 = sorted(df_filtered_by_region1[brand_col].dropna().unique())
    selected_brand1 = st.selectbox("Select Brand (Left)", available_brands1, key="brand1_sidebar")
    df_filtered_by_brand1 = df_filtered_by_region1[df_filtered_by_region1[brand_col] == selected_brand1]

    # Unit name filter
    available_units1 = sorted(df_filtered_by_brand1[unit_name_col].dropna().unique())
    selected_unit1 = st.selectbox("Unit name (Left)", available_units1, key="unit1_sidebar")
    df_filtered_by_unit1 = df_filtered_by_brand1[df_filtered_by_brand1[unit_name_col] == selected_unit1]

    # Recovery type filter
    available_recovery_types1 = sorted(df_filtered_by_unit1[recovery_col].dropna().unique())
    selected_recovery1 = st.selectbox("Recovery type (Left)", available_recovery_types1, key="recovery1_sidebar")
    df_filtered_by_recovery1 = df_filtered_by_unit1[df_filtered_by_unit1[recovery_col] == selected_recovery1]

    # Unit size filter
    available_sizes1 = sorted(df_filtered_by_recovery1[size_col].dropna().unique())
    selected_size1 = st.selectbox("Unit size (Left)", available_sizes1, key="size1_sidebar")
    df_filtered_by_size1 = df_filtered_by_recovery1[df_filtered_by_recovery1[size_col] == selected_size1]

    # Conditional dropdowns based on Recovery type
    selected_type1 = None
    selected_material1 = None
    df_temp_filtered_by_type1 = df_filtered_by_size1 # Temporary dataframe for subsequent filtering

    if selected_recovery1 == "RRG" and type_col:
        available_types1 = sorted(df_filtered_by_size1[type_col].dropna().unique())
        selected_type1 = st.selectbox("Rotary wheel type (Left)", available_types1, key="type1_sidebar")
        df_temp_filtered_by_type1 = df_filtered_by_size1[df_filtered_by_size1[type_col] == selected_type1]
    elif selected_recovery1 in ["HEX", "PCR"] and material_col:
        available_materials1 = sorted(df_filtered_by_size1[material_col].dropna().unique())
        selected_material1 = st.selectbox("PCR/HEX lamels material (Left)", available_materials1, key="material1_sidebar")
        df_temp_filtered_by_type1 = df_filtered_by_size1[df_filtered_by_size1[material_col] == selected_material1]
    
    # This is the final filtered DataFrame for the left column
    filtered_df1 = df_temp_filtered_by_type1

    st.markdown("---") # Separator for the second set of filters in sidebar

    # --- Right Column: Filters (Also in Sidebar for consistency) ---
    st.header("Select Second Unit for Comparison")

    # Year filter
    available_years2 = sorted(df[year_col].dropna().unique())
    selected_year2 = st.selectbox("Year (Right)", available_years2, key="year2_sidebar")
    df_filtered_by_year2 = df[df[year_col] == selected_year2]

    # Quarter filter
    available_quarters2 = sorted(df_filtered_by_year2[quarter_col].dropna().unique())
    selected_quarter2 = st.selectbox("Quarter (Right)", available_quarters2, key="quarter2_sidebar")
    df_filtered_by_quarter2 = df_filtered_by_year2[df_filtered_by_year2[quarter_col] == selected_quarter2]

    # Region filter
    available_regions2 = sorted(df_filtered_by_quarter2[region_col].dropna().unique())
    selected_region2 = st.selectbox("Region (Right)", available_regions2, key="region2_sidebar")
    df_filtered_by_region2 = df_filtered_by_quarter2[df_filtered_by_quarter2[region_col] == selected_region2]

    # Brand filter
    available_brands2 = sorted(df_filtered_by_region2[brand_col].dropna().unique())
    selected_brand2 = st.selectbox("Select Brand (Right)", available_brands2, key="brand2_sidebar")
    df_filtered_by_brand2 = df_filtered_by_region2[df_filtered_by_region2[brand_col] == selected_brand2]

    # Unit name filter
    available_units2 = sorted(df_filtered_by_brand2[unit_name_col].dropna().unique())
    selected_unit2 = st.selectbox("Unit name (Right)", available_units2, key="unit2_sidebar")
    df_filtered_by_unit2 = df_filtered_by_brand2[df_filtered_by_brand2[unit_name_col] == selected_unit2]

    # Recovery type filter
    available_recovery_types2 = sorted(df_filtered_by_unit2[recovery_col].dropna().unique())
    selected_recovery2 = st.selectbox("Recovery type (Right)", available_recovery_types2, key="recovery2_sidebar")
    df_filtered_by_recovery2 = df_filtered_by_unit2[df_filtered_by_unit2[recovery_col] == selected_recovery2]

    # Unit size filter
    available_sizes2 = sorted(df_filtered_by_recovery2[size_col].dropna().unique())
    selected_size2 = st.selectbox("Unit size (Right)", available_sizes2, key="size2_sidebar")
    df_filtered_by_size2 = df_filtered_by_recovery2[df_filtered_by_recovery2[size_col] == selected_size2]

    # Conditional dropdowns based on Recovery type
    selected_type2 = None
    selected_material2 = None
    df_temp_filtered_by_type2 = df_filtered_by_size2 # Temporary dataframe for subsequent filtering

    if selected_recovery2 == "RRG" and type_col:
        available_types2 = sorted(df_filtered_by_size2[type_col].dropna().unique())
        selected_type2 = st.selectbox("Rotary wheel type (Right)", available_types2, key="type2_sidebar")
        df_temp_filtered_by_type2 = df_filtered_by_size2[df_filtered_by_size2[type_col] == selected_type2]
    elif selected_recovery2 in ["HEX", "PCR"] and material_col:
        available_materials2 = sorted(df_filtered_by_size2[material_col].dropna().unique())
        selected_material2 = st.selectbox("PCR/HEX lamels material (Right)", available_materials2, key="material2_sidebar")
        df_temp_filtered_by_type2 = df_filtered_by_size2[df_filtered_by_size2[material_col] == selected_material2]
    
    # This is the final filtered DataFrame for the right column
    filtered_df2 = df_temp_filtered_by_type2

    st.markdown("---") # Separator for the second set of filters in sidebar

    # --- CSV Download Button in Sidebar ---
    # Prepare data for CSV download
    csv_data = []
    # Add headers for CSV
    csv_data.append(["Parameter", f"{selected_brand1} - {selected_unit1} - {selected_size1}", f"{selected_brand2} - {selected_unit2} - {selected_size2}"])

    # List of columns to be excluded from the comparison table display directly as text
    raw_excluded_cols_base = [
        brand_col, logo_col, unit_photo_col, year_col, quarter_col, region_col,
        unit_name_col, recovery_col, size_col, # Exclude dropdowns themselves
        type_col, material_col # Exclude type/material as they are dropdown values
    ]
    # Add all resolved coordinate column names (used only for charts)
    for x_name, y_name in coord_col_pairs_1_5 + coord_col_pairs_6_10 + coord_col_pairs_11_15:
        if x_name: raw_excluded_cols_base.append(x_name)
        if y_name: raw_excluded_cols_base.append(y_name)

    # Unit area column for the new chart should also be excluded from the main table
    unit_area_col_name = get_column_safe(df, ["Unit cross section area (Supply Filter) [m2]"])
    if unit_area_col_name: raw_excluded_cols_base.append(unit_area_col_name)

    # Filter out any None values and ensure uniqueness
    excluded_cols_from_table = list(set([col_name for col_name in raw_excluded_cols_base if col_name is not None]))
    
    # Set of headers to be conditionally excluded from display
    excluded_headers_from_display = set()

    # Conditional Hiding Logic for columns and headers based on selected Recovery types
    # Rotary wheel related columns and header if both HEX
    if selected_recovery1 == "HEX" and selected_recovery2 == "HEX":
        if get_column_safe(df, ["Wheel diameter [mm]"]): excluded_cols_from_table.append(get_column_safe(df, ["Wheel diameter [mm]"]))
        if get_column_safe(df, ["Distance between lamels [mm]"]): excluded_cols_from_table.append(get_column_safe(df, ["Distance between lamels [mm]"]))
        if type_col: excluded_cols_from_table.append(type_col) # Type column is for Rotary wheel
        excluded_headers_from_display.add("Rotary wheel")
        # Also hide RRG efficiencies if both are HEX
        if sens_efficiency_nominal_rrg_col: excluded_cols_from_table.append(sens_efficiency_nominal_rrg_col)
        if sens_efficiency_opt_rrg_col: excluded_cols_from_table.append(sens_efficiency_opt_rrg_col)


    # PCR/HEX related columns and header if both RRG
    if selected_recovery1 == "RRG" and selected_recovery2 == "RRG":
        if material_col: excluded_cols_from_table.append(material_col) # Material column is for PCR/HEX
        if sens_efficiency_nominal_pcr_hex_col: excluded_cols_from_table.append(sens_efficiency_nominal_pcr_hex_col)
        if sens_efficiency_opt_pcr_hex_col: excluded_cols_from_table.append(sens_efficiency_opt_pcr_hex_col)
        excluded_headers_from_display.add("PCR/HEX recovery exchanger")

    # Ensure uniqueness after conditional additions
    excluded_cols_from_table = list(set(excluded_cols_from_table))

    # Collect all comparison data for CSV, respecting conditional exclusions
    csv_displayed_headers = set()
    csv_data.append(["General data", "", ""]) # Initial header for CSV

    # Build CSV data based on the full column order, applying exclusions
    for col in df.columns:
        # Check if this column triggers a section header for CSV
        triggered_header_csv = None
        for trigger_col, header_title in header_triggers_map.items():
            if trigger_col == col and header_title not in csv_displayed_headers and header_title not in excluded_headers_from_display:
                triggered_header_csv = header_title
                csv_displayed_headers.add(header_title)
                break

        if triggered_header_csv:
            csv_data.append(["", "", ""]) # Blank line before new section
            csv_data.append([triggered_header_csv, "", ""]) # Add new section header

        # Add the parameter row to CSV data if not excluded
        if col not in excluded_cols_from_table:
            val1 = filtered_df1[col].values[0] if col in filtered_df1.columns else "-"
            val2 = filtered_df2[col].values[0] if col in filtered_df2.columns else "-"
            csv_data.append([col, val1, val2])


    csv_df = pd.DataFrame(csv_data)
    csv_string = csv_df.to_csv(index=False, header=False) # No header because we manually added it

    st.download_button(
        label="Download Comparison as CSV",
        data=csv_string,
        file_name="technical_data_comparison.csv",
        mime="text/csv",
        key="csv_download_sidebar" # Unique key for sidebar button
    )


# --- Main Content Area ---

# --- Image Height Synchronization (Brand Logos) ---
st.subheader("Brand Logos")
col_logo1, col_logo2 = st.columns(2)

loaded_image1 = None
brand1_logo_path_raw = filtered_df1[logo_col].iloc[0] if not filtered_df1.empty and logo_col and logo_col in filtered_df1.columns else None
brand1_logo_path = str(brand1_logo_path_raw) if pd.notna(brand1_logo_path_raw) else None

if isinstance(brand1_logo_path, str) and brand1_logo_path.strip():
    try:
        loaded_image1 = Image.open(f"images/{brand1_logo_path}")
    except FileNotFoundError:
        st.warning(f"Brand logo image not found for {selected_brand1}: images/{brand1_logo_path}")
    except Exception as e:
        st.warning(f"Error loading brand logo for {selected_brand1}: {e}")

loaded_image2 = None
brand2_logo_path_raw = filtered_df2[logo_col].iloc[0] if not filtered_df2.empty and logo_col and logo_col in filtered_df2.columns else None
brand2_logo_path = str(brand2_logo_path_raw) if pd.notna(brand2_logo_path_raw) else None

if isinstance(brand2_logo_path, str) and brand2_logo_path.strip():
    try:
        loaded_image2 = Image.open(f"images/{brand2_logo_path}")
    except FileNotFoundError:
        st.warning(f"Brand logo image not found for {selected_brand2}: images/{brand2_logo_path}")
    except Exception as e:
        st.warning(f"Error loading brand logo for {selected_brand2}: {e}")

max_logo_height = 0
if loaded_image1:
    max_logo_height = max(max_logo_height, loaded_image1.height)
if loaded_image2:
    max_logo_height = max(max_logo_height, loaded_image2.height)

with col_logo1:
    if loaded_image1:
        width1 = 150
        height1 = int(loaded_image1.height * (width1 / loaded_image1.width))
        if max_logo_height > 0:
            width1 = int(loaded_image1.width * (max_logo_height / loaded_image1.height))
            height1 = max_logo_height
        st.image(loaded_image1.resize((width1, height1)), caption=f"Logo for {selected_brand1}")
    else:
        st.write("No logo available for selected brand.")

with col_logo2:
    if loaded_image2:
        width2 = 150
        height2 = int(loaded_image2.height * (width2 / loaded_image2.width))
        if max_logo_height > 0:
            width2 = int(loaded_image2.width * (max_logo_height / loaded_image2.height))
            height2 = max_logo_height
        st.image(loaded_image2.resize((width2, height2)), caption=f"Logo for {selected_brand2}")
    else:
        st.write("No logo available for selected brand.")


# --- Image Height Synchronization (Unit Photos) ---
st.subheader("Unit Photo")
col_photo1, col_photo2 = st.columns(2)

loaded_unit_image1 = None
unit_photo_path1_raw = filtered_df1[unit_photo_col].values[0] if not filtered_df1.empty and unit_photo_col and unit_photo_col in filtered_df1.columns else None
unit_photo_path1 = str(unit_photo_path1_raw) if pd.notna(unit_photo_path1_raw) else None

if isinstance(unit_photo_path1, str) and unit_photo_path1.strip():
    try:
        loaded_unit_image1 = Image.open(f"images/{unit_photo_path1}")
    except FileNotFoundError:
        st.warning(f"Unit photo image not found for {selected_unit1}: images/{unit_photo_path1}")
    except Exception as e:
        st.warning(f"Error loading unit photo for {selected_unit1}: {e}")

loaded_unit_image2 = None
unit_photo_path2_raw = filtered_df2[unit_photo_col].values[0] if not filtered_df2.empty and unit_photo_col and unit_photo_col in filtered_df2.columns else None
unit_photo_path2 = str(unit_photo_path2_raw) if pd.notna(unit_photo_path2_raw) else None

if isinstance(unit_photo_path2, str) and unit_photo_path2.strip():
    try:
        loaded_unit_image2 = Image.open(f"images/{unit_photo_path2}")
    except FileNotFoundError:
        st.warning(f"Unit photo image not found for {selected_unit2}: images/{unit_photo_path2}")
    except Exception as e:
        st.warning(f"Error loading unit photo for {selected_unit2}: {e}")

max_unit_photo_height = 0
if loaded_unit_image1:
    max_unit_photo_height = max(max_unit_photo_height, loaded_unit_image1.height)
if loaded_unit_image2:
    max_unit_photo_height = max(max_unit_photo_height, loaded_unit_image2.height)

with col_photo1:
    if loaded_unit_image1:
        width1 = 250
        height1 = int(loaded_unit_image1.height * (width1 / loaded_unit_image1.width))
        if max_unit_photo_height > 0:
            width1 = int(loaded_unit_image1.width * (max_unit_photo_height / loaded_unit_image1.height))
            height1 = max_unit_photo_height
        st.image(loaded_unit_image1.resize((width1, height1)), caption=f"{selected_unit1} Photo")
    else:
        st.write("No unit photo available for this selection.")

with col_photo2:
    if loaded_unit_image2:
        width2 = 250
        height2 = int(loaded_unit_image2.height * (width2 / loaded_unit_image2.width))
        if max_unit_photo_height > 0:
            width2 = int(loaded_unit_image2.width * (max_unit_photo_height / loaded_unit_image2.height))
            height2 = max_unit_photo_height
        st.image(loaded_unit_image2.resize((width2, height2)), caption=f"{selected_unit2} Photo")
    else:
        st.write("No unit photo available for this selection.")


if not filtered_df1.empty and not filtered_df2.empty:
    st.subheader("General data")

    # Initial table header for Streamlit display
    col1, col2, col3 = st.columns([2, 3, 3])
    with col1:
        st.markdown("**Parameter**")
    with col2:
        st.markdown(f'<div style="text-align: center;">**{selected_brand1} - {selected_unit1} - {selected_size1}**</div>', unsafe_allow_html=True)
    with col3:
        st.markdown(f'<div style="text-align: center;">**{selected_brand2} - {selected_unit2} - {selected_size2}**</div>', unsafe_allow_html=True)

    displayed_headers = set()

    # Build a list of items to display in order
    display_items_ordered = []

    # Iterate through df.columns to build the ordered list of display items
    for col_name in df.columns:
        # Special handling for 'Unit size quantity' and its chart
        if col_name == get_column_safe(df, ["Execution"]):
            # Add 'Unit size quantity' after 'Execution'
            if unit_size_quantity_col not in excluded_cols_from_table:
                display_items_ordered.append({"type": "row", "col": unit_size_quantity_col})
            # Add the chart right after 'Unit size quantity'
            display_items_ordered.append({"type": "chart", "name": "unit_area_chart"})
            # Skip adding 'Unit size quantity' later if it was already added here
            continue # Continue to next column in df.columns

        # Check if this column starts a new header section
        # This logic ensures headers are placed *before* their associated data rows
        if col_name in header_triggers_map and header_triggers_map[col_name] not in displayed_headers and header_triggers_map[col_name] not in excluded_headers_from_display:
            display_items_ordered.append({"type": "header", "title": header_triggers_map[col_name]})
            displayed_headers.add(header_triggers_map[col_name])

        # Add the column's row if it's not excluded and not already handled (like unit_size_quantity_col)
        if col_name not in excluded_cols_from_table and col_name != unit_size_quantity_col:
            display_items_ordered.append({"type": "row", "col": col_name})

        # Insert charts after their respective trigger columns
        if col_name == internal_height_supply_filter_col:
            display_items_ordered.append({"type": "chart", "name": "chart1"})
        elif col_name == unit_cross_section_area_supply_fan_col:
            display_items_ordered.append({"type": "chart", "name": "chart2"})
        elif col_name == duct_connection_height_col:
            display_items_ordered.append({"type": "chart", "name": "chart3"})
        elif col_name == heating_elements_type_col: # Insert capacity range lines after this
            if capacity_range1_col not in excluded_cols_from_table:
                display_items_ordered.append({"type": "row", "col": capacity_range1_col})
            if capacity_range2_col not in excluded_cols_from_table:
                display_items_ordered.append({"type": "row", "col": capacity_range2_col})
            if capacity_range3_col not in excluded_cols_from_table:
                display_items_ordered.append({"type": "row", "col": capacity_range3_col})
            display_items_ordered.append({"type": "chart", "name": "electrical_heater_chart"})


    # Now iterate through the ordered display_items_ordered list
    for item in display_items_ordered:
        if item["type"] == "header":
            header_col1, header_col2, header_col3 = st.columns([1, 4, 1])
            with header_col2:
                st.markdown(f'<h4 style="text-align: center; font-size: 1.2em; margin-bottom: 0.5em; margin-top: 0.5em;">{item["title"]}</h4>', unsafe_allow_html=True)
            
            # Re-add table headers for the new section
            col1, col2, col3 = st.columns([2, 3, 3])
            with col1: st.markdown("**Parameter**")
            with col2:
                st.markdown(f'<div style="text-align: center;">**{selected_brand1} - {selected_unit1} - {selected_size1}**</div>', unsafe_allow_html=True)
            with col3:
                st.markdown(f'<div style="text-align: center;">**{selected_brand2} - {selected_unit2} - {selected_size2}**</div>', unsafe_allow_html=True)

        elif item["type"] == "row":
            col = item["col"]
            val1 = filtered_df1[col].values[0] if col in filtered_df1.columns else "-"
            val2 = filtered_df2[col].values[0] if col in filtered_df2.columns else "-"
            
            row_col1, row_col2, row_col3 = st.columns([2, 3, 3])
            with row_col1:
                st.write(f'<span style="font-family: sans-serif; font-size: 16px;">{col}</span>', unsafe_allow_html=True)
            with row_col2:
                st.markdown(f'<div style="text-align: center; font-family: sans-serif; font-size: 16px; color: green;">{val1}</div>', unsafe_allow_html=True)
            with row_col3:
                st.markdown(f'<div style="text-align: center; font-family: sans-serif; font-size: 16px; color: blue;">{val2}</div>', unsafe_allow_html=True)

        elif item["type"] == "chart":
            chart_name = item["name"]
            if chart_name == "unit_area_chart" and not unit_area_chart_displayed:
                if unit_area_col_name and unit_area_col_name in df.columns and size_col in df.columns:
                    chart_data_area = []

                    df_chart_base1 = df[
                        (df[year_col] == selected_year1) & (df[quarter_col] == selected_quarter1) &
                        (df[region_col] == selected_region1) & (df[brand_col] == selected_brand1) &
                        (df[unit_name_col] == selected_unit1) & (df[recovery_col] == selected_recovery1)
                    ].copy()
                    if selected_recovery1 == "RRG" and type_col and selected_type1:
                        df_chart_base1 = df_chart_base1[df_chart_base1[type_col] == selected_type1]
                    elif selected_recovery1 in ["HEX", "PCR"] and material_col and selected_material1:
                        df_chart_base1 = df_chart_base1[df_chart_base1[material_col] == selected_material1]

                    df_chart_base2 = df[
                        (df[year_col] == selected_year2) & (df[quarter_col] == selected_quarter2) &
                        (df[region_col] == selected_region2) & (df[brand_col] == selected_brand2) &
                        (df[unit_name_col] == selected_unit2) & (df[recovery_col] == selected_recovery2)
                    ].copy()
                    if selected_recovery2 == "RRG" and type_col and selected_type2:
                        df_chart_base2 = df_chart_base2[df_chart_base2[type_col] == selected_type2]
                    elif selected_recovery2 in ["HEX", "PCR"] and material_col and selected_material2:
                        df_chart_base2 = df_chart_base2[df_chart_base2[material_col] == selected_material2]

                    if not df_chart_base1.empty and unit_area_col_name in df_chart_base1.columns and size_col in df_chart_base1.columns:
                        for index, row in df_chart_base1.iterrows():
                            if pd.notna(row[unit_area_col_name]) and pd.notna(row[size_col]):
                                chart_data_area.append({
                                    "Brand": selected_brand1,
                                    "Unit Cross Section Area (m²)": row[unit_area_col_name],
                                    "Unit Size": row[size_col],
                                    "Selection_Label": f"Left: {selected_brand1}",
                                    "Full_Selection_Details": f"Left: {selected_year1}-{selected_quarter1}-{selected_brand1}-{row[unit_name_col]}-{row[size_col]}"
                                })

                    if not df_chart_base2.empty and unit_area_col_name in df_chart_base2.columns and size_col in df_chart_base2.columns:
                        for index, row in df_chart_base2.iterrows():
                            if pd.notna(row[unit_area_col_name]) and pd.notna(row[size_col]):
                                chart_data_area.append({
                                    "Brand": selected_brand2,
                                    "Unit Cross Section Area (m²)": row[unit_area_col_name],
                                    "Unit Size": row[size_col],
                                    "Selection_Label": f"Right: {selected_brand2}",
                                    "Full_Selection_Details": f"Right: {selected_year2}-{selected_quarter2}-{selected_brand2}-{row[unit_name_col]}-{row[size_col]}"
                                })

                if chart_data_area:
                    chart_df_area = pd.DataFrame(chart_data_area)
                    chart_df_area["Unit Size"] = chart_df_area["Unit Size"].astype(str)

                    st.markdown(f'<h4 style="font-size: 1.2em; margin-bottom: 0.2em; margin-top: 0.2em;">Unit Cross Section Area (Supply Filter) vs. Unit Size</h4>', unsafe_allow_html=True)
                    fig_area = px.scatter(chart_df_area,
                                        x="Unit Cross Section Area (m²)",
                                        y="Brand",
                                        color="Selection_Label",
                                        text="Unit Size",
                                        title=None,
                                        hover_data={"Unit Size": True, "Unit Cross Section Area (m²)": True, "Full_Selection_Details": True},
                                        color_discrete_map={
                                            f"Left: {selected_brand1}": "green",
                                            f"Right: {selected_brand2}": "blue"
                                        })
                    fig_area.update_traces(textposition='top center')
                    fig_area.update_layout(
                        xaxis_title="Unit Cross Section Area (m²)",
                        yaxis_title="Brand",
                        showlegend=True,
                        xaxis=dict(showgrid=True, gridwidth=1, gridcolor='lightgray')
                    )
                    st.plotly_chart(fig_area, use_container_width=True)
                else:
                    st.info("Unit Cross Section Area (Supply Filter) data not available for charting for one or both selections, or no valid unit sizes found for the current filter criteria.")
                unit_area_chart_displayed = True

            elif chart_name == "chart1" and not chart1_displayed:
                chart_data_1 = []
                required_coord_cols_chart1 = [name for pair in coord_col_pairs_1_5 for name in pair]
                can_plot_brand1_chart1 = True
                if not filtered_df1.empty:
                    for coord_col in required_coord_cols_chart1:
                        if coord_col not in filtered_df1.columns or pd.isna(filtered_df1[coord_col].values[0]):
                            can_plot_brand1_chart1 = False
                            break
                else:
                    can_plot_brand1_chart1 = False

                can_plot_brand2_chart1 = True
                if not filtered_df2.empty:
                    for coord_col in required_coord_cols_chart1:
                        if coord_col not in filtered_df2.columns or pd.isna(filtered_df2[coord_col].values[0]):
                            can_plot_brand2_chart1 = False
                            break
                else:
                    can_plot_brand2_chart1 = False

                if can_plot_brand1_chart1:
                    for i, (x_name, y_name) in enumerate(coord_col_pairs_1_5):
                        chart_data_1.append({
                            'X_coord_actual': filtered_df1[x_name].values[0],
                            'Y_coord_actual': filtered_df1[y_name].values[0],
                            'Display_Label': f"Left: {selected_year1}-{selected_quarter1}-{selected_brand1}-{selected_unit1}-{selected_size1}",
                            'Point_Order': i + 1
                        })
                
                if can_plot_brand2_chart1:
                    for i, (x_name, y_name) in enumerate(coord_col_pairs_1_5):
                        chart_data_1.append({
                            'X_coord_actual': filtered_df2[x_name].values[0],
                            'Y_coord_actual': filtered_df2[y_name].values[0],
                            'Display_Label': f"Right: {selected_year2}-{selected_quarter2}-{selected_brand2}-{selected_unit2}-{selected_size2}",
                            'Point_Order': i + 1
                        })
                
                if chart_data_1:
                    st.markdown(f'<h4 style="font-size: 1.2em; margin-bottom: 0.2em; margin-top: 0.2em;">Internal Cross Section area (Supply Filter)</h4>', unsafe_allow_html=True)
                    chart_df_1 = pd.DataFrame(chart_data_1)
                    chart_df_1 = chart_df_1.sort_values(by=['Display_Label', 'Point_Order'])
                    
                    fig1 = px.line(chart_df_1,
                                x="X_coord_actual",
                                y="Y_coord_actual",
                                color="Display_Label",
                                line_group="Display_Label",
                                markers=True,
                                title=None,
                                hover_data={'X_coord_actual': True, 'Y_coord_actual': True},
                                color_discrete_map={
                                    f"Left: {selected_year1}-{selected_quarter1}-{selected_brand1}-{selected_unit1}-{selected_size1}": "green",
                                    f"Right: {selected_year2}-{selected_quarter2}-{selected_brand2}-{selected_unit2}-{selected_size2}": "blue"
                                })
                    
                    fig1.update_layout(
                        xaxis_title="Unit internal width_Supply Filter (mm)",
                        yaxis_title="Unit internal height_Supply Filter (mm)",
                        hovermode="x unified",
                        legend_title_text="Selection - Year-Quarter-Brand-Unit-Size",
                        xaxis_constrain="domain",
                        yaxis_constrain="domain",
                        showlegend=True
                    )
                    fig1.update_yaxes(scaleanchor="x", scaleratio=1)
                    st.plotly_chart(fig1, use_container_width=True)
                elif not can_plot_brand1_chart1 and not can_plot_brand2_chart1:
                    st.warning("No complete coordinate data (X1-X5, Y1-Y5) found for selected units to generate Chart 1. Please ensure data is present and valid for both selections.")
                chart1_displayed = True

            elif chart_name == "chart2" and not chart2_displayed:
                chart_data_2 = []
                required_coord_cols_chart2 = [name for pair in coord_col_pairs_6_10 for name in pair]

                can_plot_brand1_chart2 = True
                if not filtered_df1.empty:
                    for coord_col in required_coord_cols_chart2:
                        if coord_col not in filtered_df1.columns or pd.isna(filtered_df1[coord_col].values[0]):
                            can_plot_brand1_chart2 = False
                            break
                else:
                    can_plot_brand1_chart2 = False

                can_plot_brand2_chart2 = True
                if not filtered_df2.empty:
                    for coord_col in required_coord_cols_chart2:
                        if coord_col not in filtered_df2.columns or pd.isna(filtered_df2[coord_col].values[0]):
                            can_plot_brand2_chart2 = False
                            break
                else:
                    can_plot_brand2_chart2 = False

                if can_plot_brand1_chart2:
                    for i, (x_name, y_name) in enumerate(coord_col_pairs_6_10):
                        chart_data_2.append({
                            'X_coord_actual': filtered_df1[x_name].values[0],
                            'Y_coord_actual': filtered_df1[y_name].values[0],
                            'Display_Label': f"Left: {selected_year1}-{selected_quarter1}-{selected_brand1}-{selected_unit1}-{selected_size1}",
                            'Point_Order': i + 1
                        })
                
                if can_plot_brand2_chart2:
                    for i, (x_name, y_name) in enumerate(coord_col_pairs_6_10):
                        chart_data_2.append({
                            'X_coord_actual': filtered_df2[x_name].values[0],
                            'Y_coord_actual': filtered_df2[y_name].values[0],
                            'Display_Label': f"Right: {selected_year2}-{selected_quarter2}-{selected_brand2}-{selected_unit2}-{selected_size2}",
                            'Point_Order': i + 1
                        })
                
                if chart_data_2:
                    st.markdown(f'<h4 style="font-size: 1.2em; margin-bottom: 0.2em; margin-top: 0.2em;">Internal Cross Section area (Supply Fan)</h4>', unsafe_allow_html=True)
                    chart_df_2 = pd.DataFrame(chart_data_2)
                    chart_df_2 = chart_df_2.sort_values(by=['Display_Label', 'Point_Order'])
                    
                    fig2 = px.line(chart_df_2,
                                x="X_coord_actual",
                                y="Y_coord_actual",
                                color="Display_Label",
                                line_group="Display_Label",
                                markers=True,
                                title=None,
                                hover_data={'X_coord_actual': True, 'Y_coord_actual': True},
                                color_discrete_map={
                                    f"Left: {selected_year1}-{selected_quarter1}-{selected_brand1}-{selected_unit1}-{selected_size1}": "green",
                                    f"Right: {selected_year2}-{selected_quarter2}-{selected_brand2}-{selected_unit2}-{selected_size2}": "blue"
                                })
                    
                    fig2.update_layout(
                        xaxis_title="Unit internal width_Supply Fan (mm)",
                        yaxis_title="Unit internal height_Supply Fan (mm)",
                        hovermode="x unified",
                        legend_title_text="Selection - Year-Quarter-Brand-Unit-Size",
                        xaxis_constrain="domain",
                        yaxis_constrain="domain",
                        showlegend=True
                    )
                    fig2.update_yaxes(scaleanchor="x", scaleratio=1)
                    st.plotly_chart(fig2, use_container_width=True)
                elif not can_plot_brand1_chart2 and not can_plot_brand2_chart2:
                    st.warning("No complete coordinate data (X6-X10, Y6-Y10) found for selected units to generate Chart 2. Please ensure data is present and valid for both selections.")
                chart2_displayed = True

            elif chart_name == "chart3" and not chart3_displayed:
                chart_data_3 = []
                required_coord_cols_chart3 = [name for pair in coord_col_pairs_11_15 for name in pair]

                can_plot_brand1_chart3 = True
                if not filtered_df1.empty:
                    all_coords_present_and_zero_or_na1 = True
                    for c in required_coord_cols_chart3:
                        if c in filtered_df1.columns and pd.notna(filtered_df1[c].values[0]) and filtered_df1[c].values[0] != 0:
                            all_coords_present_and_zero_or_na1 = False
                            break
                    
                    if all_coords_present_and_zero_or_na1 and duct_connection_diameter_col in filtered_df1.columns and pd.notna(filtered_df1[duct_connection_diameter_col].values[0]):
                        diameter1 = filtered_df1[duct_connection_diameter_col].values[0]
                        radius1 = diameter1 / 2.0
                        center_x1 = diameter1 / 2.0
                        center_y1 = diameter1 / 2.0
                        theta = np.linspace(0, 2*np.pi, 100)
                        for t in theta:
                            chart_data_3.append({
                                'X_coord_actual': center_x1 + radius1 * np.cos(t),
                                'Y_coord_actual': center_y1 + radius1 * np.sin(t),
                                'Display_Label': f"Left: {selected_year1}-{selected_quarter1}-{selected_brand1}-{selected_unit1}-{selected_size1}",
                                'Point_Order': 0
                            })
                        chart_data_3.append({
                            'X_coord_actual': center_x1 + radius1 * np.cos(0),
                            'Y_coord_actual': center_y1 + radius1 * np.sin(0),
                            'Display_Label': f"Left: {selected_year1}-{selected_quarter1}-{selected_brand1}-{selected_unit1}-{selected_size1}",
                            'Point_Order': 0
                        })
                    elif not all_coords_present_and_zero_or_na1:
                        for i, (x_name, y_name) in enumerate(coord_col_pairs_11_15):
                            if x_name in filtered_df1.columns and y_name in filtered_df1.columns and \
                               pd.notna(filtered_df1[x_name].values[0]) and pd.notna(filtered_df1[y_name].values[0]):
                                chart_data_3.append({
                                    'X_coord_actual': filtered_df1[x_name].values[0],
                                    'Y_coord_actual': filtered_df1[y_name].values[0],
                                    'Display_Label': f"Left: {selected_year1}-{selected_quarter1}-{selected_brand1}-{selected_unit1}-{selected_size1}",
                                    'Point_Order': i + 1
                                })
                            else:
                                can_plot_brand1_chart3 = False
                                st.info(f"Incomplete coordinate data (X11-X15, Y11-Y15) for 'Left: {selected_brand1} - {selected_unit1}'. Chart 3 may not include this selection.")
                                break
                    else:
                        can_plot_brand1_chart3 = False
                        st.info(f"Coordinate data (X11-X15, Y11-Y15) for 'Left: {selected_brand1} - {selected_unit1}' is all zeros/NA, but 'Duct connection Diameter' is missing or invalid. Cannot draw circle for Chart 3.")
                else:
                    can_plot_brand1_chart3 = False

                can_plot_brand2_chart3 = True
                if not filtered_df2.empty:
                    all_coords_present_and_zero_or_na2 = True
                    for c in required_coord_cols_chart3:
                        if c in filtered_df2.columns and pd.notna(filtered_df2[c].values[0]) and filtered_df2[c].values[0] != 0:
                            all_coords_present_and_zero_or_na2 = False
                            break

                    if all_coords_present_and_zero_or_na2 and duct_connection_diameter_col in filtered_df2.columns and pd.notna(filtered_df2[duct_connection_diameter_col].values[0]):
                        diameter2 = filtered_df2[duct_connection_diameter_col].values[0]
                        radius2 = diameter2 / 2.0
                        center_x2 = diameter2 / 2.0
                        center_y2 = diameter2 / 2.0
                        theta = np.linspace(0, 2*np.pi, 100)
                        for t in theta:
                            chart_data_3.append({
                                'X_coord_actual': center_x2 + radius2 * np.cos(t),
                                'Y_coord_actual': center_y2 + radius2 * np.sin(t),
                                'Display_Label': f"Right: {selected_year2}-{selected_quarter2}-{selected_brand2}-{selected_unit2}-{selected_size2}",
                                'Point_Order': 0
                            })
                        chart_data_3.append({
                            'X_coord_actual': center_x2 + radius2 * np.cos(0),
                            'Y_coord_actual': center_y2 + radius2 * np.sin(0),
                            'Display_Label': f"Right: {selected_year2}-{selected_quarter2}-{selected_brand2}-{selected_unit2}-{selected_size2}",
                            'Point_Order': 0
                        })
                    elif not all_coords_present_and_zero_or_na2:
                        for i, (x_name, y_name) in enumerate(coord_col_pairs_11_15):
                            if x_name in filtered_df2.columns and y_name in filtered_df2.columns and \
                               pd.notna(filtered_df2[x_name].values[0]) and pd.notna(filtered_df2[y_name].values[0]):
                                chart_data_3.append({
                                    'X_coord_actual': filtered_df2[x_name].values[0],
                                    'Y_coord_actual': filtered_df2[y_name].values[0],
                                    'Display_Label': f"Right: {selected_year2}-{selected_quarter2}-{selected_brand2}-{selected_unit2}-{selected_size2}",
                                    'Point_Order': i + 1
                                })
                            else:
                                can_plot_brand2_chart3 = False
                                st.info(f"Incomplete coordinate data (X11-X15, Y11-Y15) for 'Right: {selected_brand2} - {selected_unit2}'. Chart 3 may not include this selection.")
                                break
                    else:
                        can_plot_brand2_chart3 = False
                        st.info(f"Coordinate data (X11-X15, Y11-Y15) for 'Right: {selected_brand2} - {selected_unit2}' is all zeros/NA, but 'Duct connection Diameter' is missing or invalid. Cannot draw circle for Chart 3.")
                else:
                    can_plot_brand2_chart3 = False

                if chart_data_3:
                    st.markdown(f'<h4 style="font-size: 1.2em; margin-bottom: 0.2em; margin-top: 0.2em;">Supply Duct connection, mm</h4>', unsafe_allow_html=True)
                    chart_df_3 = pd.DataFrame(chart_data_3)
                    chart_df_3 = chart_df_3.sort_values(by=['Display_Label', 'Point_Order'])
                    
                    fig3 = px.line(chart_df_3,
                                x="X_coord_actual",
                                y="Y_coord_actual",
                                color="Display_Label",
                                line_group="Display_Label",
                                markers=True,
                                title=None,
                                hover_data={'X_coord_actual': True, 'Y_coord_actual': True},
                                color_discrete_map={
                                    f"Left: {selected_year1}-{selected_quarter1}-{selected_brand1}-{selected_unit1}-{selected_size1}": "green",
                                    f"Right: {selected_year2}-{selected_quarter2}-{selected_brand2}-{selected_unit2}-{selected_size2}": "blue"
                                })
                    
                    fig3.update_traces(line=dict(width=1.0))
                    fig3.update_layout(
                        xaxis_title="Supply Duct Connection Width (mm)",
                        yaxis_title="Supply Duct Connection Height (mm)",
                        hovermode="x unified",
                        legend_title_text="Selection - Year-Quarter-Brand-Unit-Size",
                        xaxis_constrain="domain",
                        yaxis_constrain="domain",
                        showlegend=True
                    )
                    fig3.update_yaxes(scaleanchor="x", scaleratio=1)
                    st.plotly_chart(fig3, use_container_width=True)
                elif not can_plot_brand1_chart3 and not can_plot_brand2_chart3:
                    st.warning("No complete coordinate data (X11-X15, Y11-Y15) or valid 'Duct connection Diameter' found for selected units to generate Chart 3. Please ensure data is present and valid for both selections.")
                chart3_displayed = True

            elif chart_name == "electrical_heater_chart" and not electrical_heater_chart_displayed:
                electrical_heater_chart_data = []

                if not filtered_df1.empty and capacity_range1_col in filtered_df1.columns and \
                   capacity_range2_col in filtered_df1.columns and capacity_range3_col in filtered_df1.columns:
                    electrical_heater_chart_data.append({
                        "Capacity Range": "Capacity range1",
                        "Value (kW)": filtered_df1[capacity_range1_col].values[0],
                        "Selection": f"Left: {selected_year1}-{selected_quarter1}-{selected_brand1}-{selected_unit1}-{selected_size1}"
                    })
                    electrical_heater_chart_data.append({
                        "Capacity Range": "Capacity range2",
                        "Value (kW)": filtered_df1[capacity_range2_col].values[0],
                        "Selection": f"Left: {selected_year1}-{selected_quarter1}-{selected_brand1}-{selected_unit1}-{selected_size1}"
                    })
                    electrical_heater_chart_data.append({
                        "Capacity Range": "Capacity range3",
                        "Value (kW)": filtered_df1[capacity_range3_col].values[0],
                        "Selection": f"Left: {selected_year1}-{selected_quarter1}-{selected_brand1}-{selected_unit1}-{selected_size1}"
                    })

                if not filtered_df2.empty and capacity_range1_col in filtered_df2.columns and \
                   capacity_range2_col in filtered_df2.columns and capacity_range3_col in filtered_df2.columns:
                    electrical_heater_chart_data.append({
                        "Capacity Range": "Capacity range1",
                        "Value (kW)": filtered_df2[capacity_range1_col].values[0],
                        "Selection": f"Right: {selected_year2}-{selected_quarter2}-{selected_brand2}-{selected_unit2}-{selected_size2}"
                    })
                    electrical_heater_chart_data.append({
                        "Capacity Range": "Capacity range2",
                        "Value (kW)": filtered_df2[capacity_range2_col].values[0],
                        "Selection": f"Right: {selected_year2}-{selected_quarter2}-{selected_brand2}-{selected_unit2}-{selected_size2}"
                    })
                    electrical_heater_chart_data.append({
                        "Capacity Range": "Capacity range3",
                        "Value (kW)": filtered_df2[capacity_range3_col].values[0],
                        "Selection": f"Right: {selected_year2}-{selected_quarter2}-{selected_brand2}-{selected_unit2}-{selected_size2}"
                    })

                if electrical_heater_chart_data:
                    st.markdown(f'<h4 style="font-size: 1.2em; margin-bottom: 0.2em; margin-top: 0.2em;">Electrical Heater Capacity (kW)</h4>', unsafe_allow_html=True)
                    electrical_heater_df = pd.DataFrame(electrical_heater_chart_data)
                    
                    fig_heater = px.bar(electrical_heater_df,
                                        x="Capacity Range",
                                        y="Value (kW)",
                                        color="Selection",
                                        barmode="group",
                                        title=None,
                                        color_discrete_map={
                                            f"Left: {selected_year1}-{selected_quarter1}-{selected_brand1}-{selected_unit1}-{selected_size1}": "green",
                                            f"Right: {selected_year2}-{selected_quarter2}-{selected_brand2}-{selected_unit2}-{selected_size2}": "blue"
                                        })
                    
                    fig_heater.update_layout(
                        hovermode="x unified",
                        legend_title_text="Selection - Year-Quarter-Brand-Unit-Size",
                        xaxis_title="Capacity Range",
                        yaxis_title="Capacity (kW)"
                    )
                    st.plotly_chart(fig_heater, use_container_width=True)
                else:
                    st.warning("No complete capacity data found for Electrical Heater to generate the chart.")
                electrical_heater_chart_displayed = True


else:
    st.warning("One of the selected combinations has no data to display for comparison. Please adjust your selections.")
