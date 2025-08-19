import streamlit as st
import pandas as pd
from PIL import Image
import plotly.express as px
import numpy as np

# Set the maximum number of units for comparison
MAX_UNITS = 10

# Load data
@st.cache_data
def load_data():
    url = "Data_2025.xlsx"
    return pd.read_excel(url, sheet_name="data", engine='openpyxl')

df = load_data()

# Initialize session state for the list of selected units
if 'selections' not in st.session_state:
    st.session_state.selections = [{}] # Start with one empty selection

# --- Helper functions for robust column access ---
def get_column_safe(df, name_options):
    for name in name_options:
        if name in df.columns:
            return name
    return None

# --- Main App Title and Layout ---
st.title("Technical Data Comparison")

# Use a two-column layout for the sidebar and main content
# The sidebar will be for filters, and the main area for the results
filter_container = st.sidebar
main_container = st.container()

# --- Dynamic Filter Generation in Sidebar ---
with filter_container:
    st.header("Select Units for Comparison")

    # Button to add a new unit
    if st.button("Add another unit for comparison"):
        if len(st.session_state.selections) < MAX_UNITS:
            st.session_state.selections.append({})
        else:
            st.warning(f"You can only compare up to {MAX_UNITS} units at once.")

    # Button to clear all selections
    if st.button("Clear all selections"):
        st.session_state.selections = [{}]

    # Loop through the list of selections to create dynamic filters
    for i, selection in enumerate(st.session_state.selections):
        with st.expander(f"Unit {i+1} Selection", expanded=True):
            # Use unique keys for each widget
            st.session_state.selections[i]['year'] = st.selectbox(
                "Year", sorted(df[get_column_safe(df, ["Year"])].dropna().unique()), key=f"year_{i}"
            )
            
            df_filtered_by_year = df[df[get_column_safe(df, ["Year"])] == st.session_state.selections[i]['year']]

            st.session_state.selections[i]['quarter'] = st.selectbox(
                "Quarter", sorted(df_filtered_by_year[get_column_safe(df, ["Quarter"])].dropna().unique()), key=f"quarter_{i}"
            )
            
            df_filtered_by_quarter = df_filtered_by_year[df_filtered_by_year[get_column_safe(df, ["Quarter"])] == st.session_state.selections[i]['quarter']]

            st.session_state.selections[i]['region'] = st.selectbox(
                "Region", sorted(df_filtered_by_quarter[get_column_safe(df, ["Region"])].dropna().unique()), key=f"region_{i}"
            )
            
            df_filtered_by_region = df_filtered_by_quarter[df_filtered_by_quarter[get_column_safe(df, ["Region"])] == st.session_state.selections[i]['region']]

            st.session_state.selections[i]['brand'] = st.selectbox(
                "Brand", sorted(df_filtered_by_region[get_column_safe(df, ["Brand name"])].dropna().unique()), key=f"brand_{i}"
            )
            
            df_filtered_by_brand = df_filtered_by_region[df_filtered_by_region[get_column_safe(df, ["Brand name"])] == st.session_state.selections[i]['brand']]

            st.session_state.selections[i]['unit_name'] = st.selectbox(
                "Unit name", sorted(df_filtered_by_brand[get_column_safe(df, ["Unit name"])].dropna().unique()), key=f"unit_name_{i}"
            )
            
            df_filtered_by_unit_name = df_filtered_by_brand[df_filtered_by_brand[get_column_safe(df, ["Unit name"])] == st.session_state.selections[i]['unit_name']]

            st.session_state.selections[i]['recovery'] = st.selectbox(
                "Recovery type", sorted(df_filtered_by_unit_name[get_column_safe(df, ["Recovery type"])].dropna().unique()), key=f"recovery_{i}"
            )
            
            df_filtered_by_recovery = df_filtered_by_unit_name[df_filtered_by_unit_name[get_column_safe(df, ["Recovery type"])] == st.session_state.selections[i]['recovery']]

            st.session_state.selections[i]['size'] = st.selectbox(
                "Unit size", sorted(df_filtered_by_recovery[get_column_safe(df, ["Unit size"])].dropna().unique()), key=f"size_{i}"
            )

            # Conditional dropdowns based on Recovery type
            current_recovery_type = st.session_state.selections[i]['recovery']
            df_filtered_for_type = df_filtered_by_recovery[df_filtered_by_recovery[get_column_safe(df, ["Unit size"])] == st.session_state.selections[i]['size']]

            type_col = get_column_safe(df, ["Type"])
            material_col = get_column_safe(df, ["Material"])

            if current_recovery_type == "RRG" and type_col:
                st.session_state.selections[i]['type'] = st.selectbox(
                    "Rotary wheel type", sorted(df_filtered_for_type[type_col].dropna().unique()), key=f"type_{i}"
                )
            else:
                st.session_state.selections[i]['type'] = None
            
            if current_recovery_type in ["HEX", "PCR"] and material_col:
                st.session_state.selections[i]['material'] = st.selectbox(
                    "PCR/HEX lamels material", sorted(df_filtered_for_type[material_col].dropna().unique()), key=f"material_{i}"
                )
            else:
                st.session_state.selections[i]['material'] = None

            # Optional: Add a button to remove this selection
            if len(st.session_state.selections) > 1 and st.button(f"Remove Unit {i+1}", key=f"remove_btn_{i}"):
                st.session_state.selections.pop(i)
                st.experimental_rerun()

# --- Main Content Area - Building the Comparison ---
with main_container:
    selected_dfs = []
    # Create a list of DataFrames based on the selections
    for selection in st.session_state.selections:
        df_temp = df.copy()
        for key, value in selection.items():
            col = get_column_safe(df, [key.capitalize().replace('_', ' ').replace('unit name', 'Unit name').replace('unit size', 'Unit size').replace('brand', 'Brand name').replace('eurovent model box', 'Eurovent Model Box').replace('recovery type', 'Recovery type').replace('type', 'Type').replace('material', 'Material')])
            if col and value:
                df_temp = df_temp[df_temp[col] == value]
        if not df_temp.empty:
            selected_dfs.append(df_temp)
        else:
            # If a selection doesn't result in a row, add an empty placeholder DataFrame
            selected_dfs.append(pd.DataFrame())

    if not selected_dfs or all(df_item.empty for df_item in selected_dfs):
        st.info("Please select at least one unit for comparison.")
    else:
        # Define columns for the main content area
        # Ratio for Parameter | Unit 1 | Unit 2 | ...
        col_ratios = [2] + [1] * len(selected_dfs)
        columns = st.columns(col_ratios)

        # --- Brand Logos ---
        with main_container:
            st.subheader("Brand Logos")
            logo_col_name = get_column_safe(df, ["Brand logo"])
            max_logo_height = 0
            
            # First pass: find max height to sync them
            loaded_images = []
            for df_item in selected_dfs:
                image = None
                if not df_item.empty and logo_col_name and logo_col_name in df_item.columns:
                    logo_path = str(df_item[logo_col_name].iloc[0])
                    if pd.notna(logo_path) and logo_path.strip():
                        try:
                            image = Image.open(f"images/{logo_path}")
                            if image:
                                max_logo_height = max(max_logo_height, image.height)
                        except FileNotFoundError:
                            st.warning(f"Brand logo image not found: images/{logo_path}")
                        except Exception as e:
                            st.warning(f"Error loading brand logo: {e}")
                loaded_images.append(image)

            for i, col in enumerate(columns[1:]):
                with col:
                    image = loaded_images[i]
                    if image:
                        width = 150
                        height = int(image.height * (width / image.width))
                        if max_logo_height > 0:
                            width = int(image.width * (max_logo_height / image.height))
                            height = max_logo_height
                        st.image(image.resize((width, height)), use_column_width=False)
                    else:
                        st.write("No logo")

        # --- Unit Photos ---
        with main_container:
            st.subheader("Unit Photo")
            photo_col_name = get_column_safe(df, ["Unit photo"])
            max_photo_height = 0
            
            loaded_photos = []
            for df_item in selected_dfs:
                photo = None
                if not df_item.empty and photo_col_name and photo_col_name in df_item.columns:
                    photo_path = str(df_item[photo_col_name].iloc[0])
                    if pd.notna(photo_path) and photo_path.strip():
                        try:
                            photo = Image.open(f"images/{photo_path}")
                            if photo:
                                max_photo_height = max(max_photo_height, photo.height)
                        except FileNotFoundError:
                            st.warning(f"Unit photo image not found: images/{photo_path}")
                        except Exception as e:
                            st.warning(f"Error loading unit photo: {e}")
                loaded_photos.append(photo)

            for i, col in enumerate(columns[1:]):
                with col:
                    photo = loaded_photos[i]
                    if photo:
                        width = 250
                        height = int(photo.height * (width / photo.width))
                        if max_photo_height > 0:
                            width = int(photo.width * (max_photo_height / photo.height))
                            height = max_photo_height
                        st.image(photo.resize((width, height)), use_column_width=False)
                    else:
                        st.write("No photo")

        # --- Dynamic Comparison Table ---
        with main_container:
            st.subheader("General data")
            
            # Columns to be excluded from the comparison table display
            excluded_cols_base = [
                get_column_safe(df, ["Brand logo", "Brand logo", "Brand Logo"]),
                get_column_safe(df, ["Unit photo", "Unit photo", "Unit Photo", "Unit Photo Name"]),
                get_column_safe(df, ["Year"]),
                get_column_safe(df, ["Quarter"]),
                get_column_safe(df, ["Region"]),
                get_column_safe(df, ["Brand name"]),
                get_column_safe(df, ["Unit name"]),
                get_column_safe(df, ["Recovery type"]),
                get_column_safe(df, ["Unit size"]),
                get_column_safe(df, ["Type"]),
                get_column_safe(df, ["Material"]),
                get_column_safe(df, ["Unit cross section area (Supply Filter) [m2]"])
            ]
            
            # Add all resolved coordinate column names (used only for charts)
            coord_col_pairs_1_5 = []
            for i in range(1, 6):
                x_col_name = get_column_safe(df, [f"x{i}", f"X{i}", f"X{i}_coord", f"x{i}_coord"])
                y_col_name = get_column_safe(df, [f"y{i}", f"Y{i}", f"Y{i}_coord", f"y{i}_coord"])
                if x_col_name and y_col_name:
                    coord_col_pairs_1_5.append((x_col_name, y_col_name))
            
            coord_col_pairs_6_10 = []
            for i in range(6, 11):
                x_col_name = get_column_safe(df, [f"x{i}", f"X{i}", f"X{i}_coord", f"x{i}_coord"])
                y_col_name = get_column_safe(df, [f"y{i}", f"Y{i}", f"Y{i}_coord", f"y{i}_coord"])
                if x_col_name and y_col_name:
                    coord_col_pairs_6_10.append((x_col_name, y_col_name))
            
            coord_col_pairs_11_15 = []
            for i in range(11, 16):
                x_col_name = get_column_safe(df, [f"x{i}", f"X{i}", f"X{i}_coord", f"x{i}_coord"])
                y_col_name = get_column_safe(df, [f"y{i}", f"Y{i}", f"Y{i}_coord", f"y{i}_coord"])
                if x_col_name and y_col_name:
                    coord_col_pairs_11_15.append((x_col_name, y_col_name))
            
            for x_name, y_name in coord_col_pairs_1_5 + coord_col_pairs_6_10 + coord_col_pairs_11_15:
                if x_name: excluded_cols_base.append(x_name)
                if y_name: excluded_cols_base.append(y_name)

            # Define header triggers
            header_triggers_map = {
                get_column_safe(df, ["Eurovent Certificate"]): "Certification data",
                get_column_safe(df, ["Supply"]): "Available configurations",
                get_column_safe(df, ["Insulation material"]): "Casing",
                get_column_safe(df, ["Minimum airflow [CMH]"]): "Airflows",
                get_column_safe(df, ["Internal Width (Supply Filter) [mm]"]): "Overall dimensions",
                get_column_safe(df, ["Type"]): "Rotary wheel",
                get_column_safe(df, ["Sens. efficiency at nominal balanced airflows_PCR/HEX [%]"]): "PCR/HEX recovery exchanger",
                get_column_safe(df, ["Motor type"]): "Fan section data",
                get_column_safe(df, ["Heating elements type"]): "Electrical heater",
                get_column_safe(df, ["Water heater_min rows"]): "Water heater",
                get_column_safe(df, ["Water cooler_min rows"]): "Water cooler",
                get_column_safe(df, ["DXH_min rows"]): "DX/DXH cooler",
                get_column_safe(df, ["Filter type_Supply"]): "Supply Filter",
                get_column_safe(df, ["Filter type_Exhaust"]): "Exhaust Filter",
                get_column_safe(df, ["Silencer casing"]): "Silencer data"
            }
            
            # Get a list of all column names that exist in the dataframe
            all_cols = df.columns.tolist()

            # First row for headers
            with columns[0]:
                st.markdown("**Parameter**")
            for i, col in enumerate(columns[1:]):
                with col:
                    df_item = selected_dfs[i]
                    if not df_item.empty:
                        brand = df_item[get_column_safe(df, ["Brand name"])].iloc[0]
                        unit = df_item[get_column_safe(df, ["Unit name"])].iloc[0]
                        size = df_item[get_column_safe(df, ["Unit size"])].iloc[0]
                        st.markdown(f'<div style="text-align: center; font-size:12px;">**{brand} - {unit} - {size}**</div>', unsafe_allow_html=True)
                    else:
                        st.markdown(f'<div style="text-align: center;">**Invalid Selection**</div>', unsafe_allow_html=True)
            
            # Loop through all columns to build the dynamic table
            displayed_headers = set()
            for col_name in all_cols:
                # Check for header triggers and insert header if needed
                if col_name in header_triggers_map:
                    header_title = header_triggers_map[col_name]
                    if header_title not in displayed_headers:
                        header_col_list = st.columns([1, len(selected_dfs)])
                        with header_col_list[0]:
                           st.markdown(f'<h4 style="text-align: left; font-size: 1.2em; margin-bottom: 0.5em; margin-top: 0.5em;">{header_title}</h4>', unsafe_allow_html=True)
                        displayed_headers.add(header_title)

                # Check if the column should be displayed in the table
                if col_name not in excluded_cols_base:
                    row_cols = st.columns(col_ratios)
                    with row_cols[0]:
                        st.write(f'<span style="font-family: sans-serif; font-size: 14px;">{col_name}</span>', unsafe_allow_html=True)
                    
                    for i, col in enumerate(row_cols[1:]):
                        with col:
                            df_item = selected_dfs[i]
                            val = df_item[col_name].iloc[0] if not df_item.empty and col_name in df_item.columns else "-"
                            st.markdown(f'<div style="text-align: center; font-family: sans-serif; font-size: 14px;">{val}</div>', unsafe_allow_html=True)
            
            # --- Dynamic Chart Generation ---
            # Helper function to get plot data and a list of labels for the legend
            def get_plot_data(df_list, coord_pairs, unit_label_fn):
                plot_data = []
                for i, df_item in enumerate(df_list):
                    if not df_item.empty:
                        label = unit_label_fn(i)
                        for x_name, y_name in coord_pairs:
                            if x_name in df_item.columns and y_name in df_item.columns:
                                x_val = df_item[x_name].iloc[0]
                                y_val = df_item[y_name].iloc[0]
                                if pd.notna(x_val) and pd.notna(y_val):
                                    plot_data.append({
                                        'X_coord_actual': x_val,
                                        'Y_coord_actual': y_val,
                                        'Display_Label': label,
                                        'Point_Order': coord_pairs.index((x_name, y_name)) + 1
                                    })
                return plot_data

            def get_chart_label(i):
                selection = st.session_state.selections[i]
                return f"Unit {i+1}: {selection['brand']} - {selection['unit_name']} - {selection['size']}"

            # Chart 1: Internal Cross Section area (Supply Filter)
            st.subheader("Internal Cross Section area (Supply Filter)")
            plot_data_1 = get_plot_data(selected_dfs, coord_col_pairs_1_5, get_chart_label)
            if plot_data_1:
                chart_df_1 = pd.DataFrame(plot_data_1)
                fig1 = px.line(chart_df_1,
                               x="X_coord_actual",
                               y="Y_coord_actual",
                               color="Display_Label",
                               line_group="Display_Label",
                               markers=True,
                               title=None)
                fig1.update_layout(xaxis_title="Unit internal width (mm)", yaxis_title="Unit internal height (mm)",
                                   hovermode="x unified", legend_title_text="Selections")
                st.plotly_chart(fig1, use_container_width=True)
            else:
                st.info("No complete data found for Chart 1.")

            # Chart 2: Internal Cross Section area (Supply Fan)
            st.subheader("Internal Cross Section area (Supply Fan)")
            plot_data_2 = get_plot_data(selected_dfs, coord_col_pairs_6_10, get_chart_label)
            if plot_data_2:
                chart_df_2 = pd.DataFrame(plot_data_2)
                fig2 = px.line(chart_df_2,
                               x="X_coord_actual",
                               y="Y_coord_actual",
                               color="Display_Label",
                               line_group="Display_Label",
                               markers=True,
                               title=None)
                fig2.update_layout(xaxis_title="Unit internal width (mm)", yaxis_title="Unit internal height (mm)",
                                   hovermode="x unified", legend_title_text="Selections")
                st.plotly_chart(fig2, use_container_width=True)
            else:
                st.info("No complete data found for Chart 2.")

            # Chart 3: Supply Duct connection, mm
            st.subheader("Supply Duct connection, mm")
            plot_data_3 = []
            duct_connection_diameter_col = get_column_safe(df, ["Duct connection Diameter [mm]"])

            for i, df_item in enumerate(selected_dfs):
                if not df_item.empty:
                    label = get_chart_label(i)
                    all_coords_present_and_zero = True
                    for c in [name for pair in coord_col_pairs_11_15 for name in pair]:
                        if c in df_item.columns and pd.notna(df_item[c].iloc[0]) and df_item[c].iloc[0] != 0:
                            all_coords_present_and_zero = False
                            break
                    
                    if all_coords_present_and_zero and duct_connection_diameter_col in df_item.columns and pd.notna(df_item[duct_connection_diameter_col].iloc[0]):
                        diameter = df_item[duct_connection_diameter_col].iloc[0]
                        radius = diameter / 2.0
                        center_x = diameter / 2.0
                        center_y = diameter / 2.0
                        theta = np.linspace(0, 2*np.pi, 100)
                        for t in theta:
                            plot_data_3.append({
                                'X_coord_actual': center_x + radius * np.cos(t),
                                'Y_coord_actual': center_y + radius * np.sin(t),
                                'Display_Label': label,
                                'Point_Order': 0
                            })
                        # Ensure the circle is closed
                        plot_data_3.append({
                            'X_coord_actual': center_x + radius * np.cos(0),
                            'Y_coord_actual': center_y + radius * np.sin(0),
                            'Display_Label': label,
                            'Point_Order': 0
                        })
                    else:
                        for x_name, y_name in coord_col_pairs_11_15:
                            if x_name in df_item.columns and y_name in df_item.columns and pd.notna(df_item[x_name].iloc[0]) and pd.notna(df_item[y_name].iloc[0]):
                                plot_data_3.append({
                                    'X_coord_actual': df_item[x_name].iloc[0],
                                    'Y_coord_actual': df_item[y_name].iloc[0],
                                    'Display_Label': label,
                                    'Point_Order': coord_col_pairs_11_15.index((x_name, y_name)) + 1
                                })
            
            if plot_data_3:
                chart_df_3 = pd.DataFrame(plot_data_3)
                fig3 = px.line(chart_df_3,
                               x="X_coord_actual",
                               y="Y_coord_actual",
                               color="Display_Label",
                               line_group="Display_Label",
                               markers=True,
                               title=None)
                fig3.update_layout(xaxis_title="Duct Connection Width (mm)", yaxis_title="Duct Connection Height (mm)",
                                   hovermode="x unified", legend_title_text="Selections")
                st.plotly_chart(fig3, use_container_width=True)
            else:
                st.info("No complete data found for Chart 3.")

            # Electrical Heater Capacity Chart
            st.subheader("Electrical Heater Capacity (kW)")
            electrical_heater_chart_data = []
            capacity_cols = [
                get_column_safe(df, ["Capacity range1 [kW]"]),
                get_column_safe(df, ["Capacity range2 [kW]"]),
                get_column_safe(df, ["Capacity range3 [kW]"])
            ]

            for i, df_item in enumerate(selected_dfs):
                if not df_item.empty:
                    label = get_chart_label(i)
                    for capacity_col in capacity_cols:
                        if capacity_col and capacity_col in df_item.columns and pd.notna(df_item[capacity_col].iloc[0]):
                            electrical_heater_chart_data.append({
                                "Capacity Range": capacity_col,
                                "Value (kW)": df_item[capacity_col].iloc[0],
                                "Selection": label
                            })
            
            if electrical_heater_chart_data:
                electrical_heater_df = pd.DataFrame(electrical_heater_chart_data)
                fig_heater = px.bar(electrical_heater_df,
                                    x="Capacity Range",
                                    y="Value (kW)",
                                    color="Selection",
                                    barmode="group",
                                    title=None)
                fig_heater.update_layout(hovermode="x unified", legend_title_text="Selections")
                st.plotly_chart(fig_heater, use_container_width=True)
            else:
                st.info("No complete capacity data found for Electrical Heater to generate the chart.")

