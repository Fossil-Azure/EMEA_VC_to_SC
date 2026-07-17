import io
import json
import os
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Excel Static Mapping Automation", page_icon="🚀")
st.title("🌍 Excel Static Mapping Automation")

# Setup project relative directories
BASE_DIR = os.path.dirname(__file__)
CONFIG_DIR = os.path.join(BASE_DIR, "country_config")
SC_TEMPLATES_DIR = os.path.join(BASE_DIR, "SC_Templates")

# Validate repository integrity
if not os.path.exists(CONFIG_DIR) or not os.path.exists(SC_TEMPLATES_DIR):
    st.error(
        "❌ Repository configuration missing. Ensure 'country_config' and 'SC_Templates' folders exist in your repo."
    )
    st.stop()

available_countries = [
    os.path.splitext(f)[0] for f in os.listdir(CONFIG_DIR) if f.endswith(".json")
]

if not available_countries:
    st.warning("⚠️ No JSON configurations found in the country_config folder.")
    st.stop()

# 1. UI Sidebar Configuration
st.sidebar.header("Step 1: Choose Profile")
target_country = st.sidebar.selectbox("Select Target Country", available_countries)

# Load Selected Metadata Matrix
config_path = os.path.join(CONFIG_DIR, f"{target_country}.json")
with open(config_path, "r", encoding="utf-8") as f:
    meta = json.load(f)

# Safely extract sheet configurations from your JSON schema
target_sheet = meta.get("sheet_name", "Template-WATCH")
sc_template_name = meta.get("sc_name_pattern")
vc_suffix = meta.get("vc_name_suffix")

st.sidebar.info(
    f"**Marketplace ID:** {meta['m_id']}\n\n"
    f"**Language Tag:** {meta['lang']}\n\n"
    f"**Target Sheet/Tab:** {target_sheet}\n\n"
    f"**Expected Base Template:** {sc_template_name}"
)

# Locate matching static template inside Git path
blank_sc_file = os.path.join(SC_TEMPLATES_DIR, sc_template_name)
if not os.path.exists(blank_sc_file):
    st.error(
        f"❌ Structural match failure: File `{sc_template_name}` was not found inside the repository's `SC_Templates` folder."
    )
    st.stop()

# 2. Document Upload Gateway
st.subheader("1. Source Document")
vc_file_upload = st.file_uploader(
    f"Upload Vendor Central Filled File (Expecting suffix matching: {vc_suffix})",
    type=["xlsx", "xlsm"],
)

# 3. Automation Processing Loop
if vc_file_upload:
    st.subheader("2. Execute Processing")

    if st.button("🚀 Run Automation Engine", type="primary"):
        output_filename = f"SC_Watch_Static_Mapped_{target_country}.xlsm"

        try:
            with st.spinner("Executing structural mapping schemas... Please wait."):

                # Dynamic mapping dictionary extraction
                base_static_map = {
                    "vendor_sku#1.value": "contribution_sku#1.value",
                    "product_type#1.value": "product_type#1.value",
                    "parentage_level#1.value": "parentage_level[marketplace_id={m_id}]#1.value",
                    "child_parent_sku_relationship#1.parent_sku": "child_parent_sku_relationship[marketplace_id={m_id}]#1.parent_sku",
                    "variation_theme#1.name": "variation_theme#1.name",
                    "item_name#1.value": "item_name[marketplace_id={m_id}][language_tag={lang}]#1.value",
                    "brand#1.value": "brand[marketplace_id={m_id}][language_tag={lang}]#1.value",
                    "external_product_id#1.type": "amzn1.volt.ca.product_id_type",
                    "external_product_id#1.value": "amzn1.volt.ca.product_id_value",
                    "model_number#1.value": "model_number[marketplace_id={m_id}]#1.value",
                    "model_name#1.value": "model_name[marketplace_id={m_id}][language_tag={lang}]#1.value",
                    "manufacturer#1.value": "manufacturer[marketplace_id={m_id}][language_tag={lang}]#1.value",
                    "rtip_product_description#1.value": "product_description[marketplace_id=A1F83G8C2ARO7P][language_tag=en_GB]#1.value",
                    "bullet_point#1.value": "bullet_point[marketplace_id={m_id}][language_tag={lang}]#1.value",
                    "bullet_point#2.value": "bullet_point[marketplace_id={m_id}][language_tag={lang}]#2.value",
                    "bullet_point#3.value": "bullet_point[marketplace_id={m_id}][language_tag={lang}]#3.value",
                    "bullet_point#4.value": "bullet_point[marketplace_id={m_id}][language_tag={lang}]#4.value",
                    "bullet_point#5.value": "bullet_point[marketplace_id={m_id}][language_tag={lang}]#5.value",
                    "generic_keyword#1.value": "generic_keyword[marketplace_id={m_id}][language_tag={lang}]#1.value",
                    "style#1.value": "style[marketplace_id={m_id}][language_tag={lang}]#1.value",
                    "department#1.value": "department[marketplace_id={m_id}][language_tag={lang}]#1.value",
                    "target_gender#1.value": "target_gender[marketplace_id={m_id}]#1.value",
                    "part_number#1.value": "part_number[marketplace_id={m_id}]#1.value",
                    "number_of_items#1.value": "number_of_items[marketplace_id={m_id}]#1.value",
                    "strap_type#1.value": "strap_type[marketplace_id={m_id}][language_tag={lang}]#1.value",
                    "water_resistance_level#1.value": "water_resistance_level[marketplace_id={m_id}]#1.value",
                    "display#1.type#1.value": "display[marketplace_id={m_id}]#1.type[language_tag={lang}]#1.value",
                    "color#1.standardized_values#1": "color[marketplace_id={m_id}][language_tag={lang}]#1.standardized_values#1",
                    "color#1.value": "color[marketplace_id={m_id}][language_tag={lang}]#1.value",
                    "clasp_type#1.value": "clasp_type[marketplace_id={m_id}][language_tag={lang}]#1.value",
                    "calendar_type#1.value": "calendar_type[marketplace_id={m_id}][language_tag={lang}]#1.value",
                    "dial_window#1.material#1.value": "dial_window[marketplace_id={m_id}]#1.material[language_tag={lang}]#1.value",
                    "case#1.diameter#1.value": "case[marketplace_id={m_id}]#1.diameter#1.value",
                    "case#1.diameter#1.unit": "case[marketplace_id={m_id}]#1.diameter#1.unit",
                    "case#1.material#1.value": "case[marketplace_id={m_id}]#1.material[language_tag={lang}]#1.value",
                    "dial#1.color#1.value": "dial[marketplace_id={m_id}]#1.color[language_tag={lang}]#1.value",
                    "warranty_type#1.value": "warranty_type[marketplace_id={m_id}][language_tag={lang}]#1.value",
                    "watch_movement_type#1.value": "watch_movement_type[marketplace_id={m_id}][language_tag={lang}]#1.value",
                    "water_resistance_depth#1.value": "water_resistance_depth[marketplace_id={m_id}]#1.value",
                    "water_resistance_depth#1.unit": "water_resistance_depth[marketplace_id={m_id}]#1.unit",
                    "item_shape#1.value": "item_shape[marketplace_id={m_id}][language_tag={lang}]#1.value",
                    "power_source_type#1.value": "power_source_type[marketplace_id={m_id}][language_tag={lang}]#1.value",
                    "included_components#1.value": "included_components[marketplace_id={m_id}][language_tag={lang}]#1.value",
                    "band#1.color#1.value": "band[marketplace_id={m_id}]#1.color[language_tag={lang}]#1.value",
                    "band#1.material#1.value": "band[marketplace_id={m_id}]#1.material[language_tag={lang}]#1.value",
                    "gdpr_risk#1.value": "gdpr_risk[marketplace_id={m_id}]#1.value",
                    "item_package_dimensions#1.length.value": "item_package_dimensions[marketplace_id={m_id}]#1.length.value",
                    "item_package_dimensions#1.length.unit": "item_package_dimensions[marketplace_id={m_id}]#1.length.unit",
                    "item_package_dimensions#1.width.value": "item_package_dimensions[marketplace_id={m_id}]#1.width.value",
                    "item_package_dimensions#1.width.unit": "item_package_dimensions[marketplace_id={m_id}]#1.width.unit",
                    "item_package_dimensions#1.height.value": "item_package_dimensions[marketplace_id={m_id}]#1.height.value",
                    "item_package_dimensions#1.height.unit": "item_package_dimensions[marketplace_id={m_id}]#1.height.unit",
                    "item_package_weight#1.value": "item_package_weight[marketplace_id={m_id}]#1.value",
                    "item_package_weight#1.unit": "item_package_weight[marketplace_id={m_id}]#1.unit",
                    "country_of_origin#1.value": "country_of_origin[marketplace_id={m_id}]#1.value",
                    "batteries_required#1.value": "batteries_required[marketplace_id={m_id}]#1.value",
                    "batteries_included#1.value": "batteries_included[marketplace_id={m_id}]#1.value",
                    "battery#1.cell_composition#1.value": "battery[marketplace_id={m_id}]#1.cell_composition#1.value",
                    "num_batteries#1.quantity": "num_batteries[marketplace_id={m_id}]#1.quantity",
                    "supplier_declared_dg_hz_regulation#1.value": "supplier_declared_dg_hz_regulation[marketplace_id={m_id}]#1.value",
                    "has_multiple_battery_powered_components#1.value": "has_multiple_battery_powered_components[marketplace_id={m_id}]#1.value",
                    "non_lithium_battery_packaging#1.value": "non_lithium_battery_packaging[marketplace_id=A1F83G8C2ARO7P]#1.value",
                    "battery_installation_device_type#1.value": "battery_installation_device_type[marketplace_id={m_id}]#1.value",
                }

                combined_map = base_static_map.copy()
                custom_mappings = meta.get("additional_mappings", {})
                if custom_mappings:
                    combined_map.update(custom_mappings)

                current_static_map = {}
                for vc_col, sc_template_str in combined_map.items():
                    current_static_map[vc_col] = sc_template_str.format(
                        m_id=meta["m_id"], lang=meta["lang"]
                    )

                current_static_values = {}
                raw_static_values = meta.get("static_values", {})
                for sc_template_str, static_val in raw_static_values.items():
                    localized_sc_col = sc_template_str.format(
                        m_id=meta["m_id"], lang=meta["lang"]
                    )
                    current_static_values[localized_sc_col] = static_val

                vc_df = pd.read_excel(vc_file_upload, sheet_name="Template-WATCH", header=3)
                vc_df.columns = [str(c).strip() for c in vc_df.columns]

                cfg_trans = meta["translations"]
                dropdown_translations = {
                    "skip_offer[marketplace_id={m_id}]#1.value".format(m_id=meta["m_id"]): cfg_trans.get("skip_offer", {}),
                    "batteries_required[marketplace_id={m_id}]#1.value".format(m_id=meta["m_id"]): cfg_trans.get("batteries_required", {}),
                    "batteries_included[marketplace_id={m_id}]#1.value".format(m_id=meta["m_id"]): cfg_trans.get("batteries_included", {}),
                    "item_package_dimensions[marketplace_id={m_id}]#1.length.unit".format(m_id=meta["m_id"]): cfg_trans.get("length_units", {}),
                    "item_package_dimensions[marketplace_id={m_id}]#1.width.unit".format(m_id=meta["m_id"]): cfg_trans.get("length_units", {}),
                    "item_package_dimensions[marketplace_id={m_id}]#1.height.unit".format(m_id=meta["m_id"]): cfg_trans.get("length_units", {}),
                    "item_package_weight[marketplace_id={m_id}]#1.unit".format(m_id=meta["m_id"]): cfg_trans.get("weight_units", {}),
                    "supplier_declared_dg_hz_regulation[marketplace_id={m_id}]#1.value".format(m_id=meta["m_id"]): cfg_trans.get("supplier_declared_dg_hz_regulation", {}),
                    "battery[marketplace_id={m_id}]#1.cell_composition#1.value".format(m_id=meta["m_id"]): cfg_trans.get("battery_cell_composition", {}),
                    "gdpr_risk[marketplace_id={m_id}]#1.value".format(m_id=meta["m_id"]): cfg_trans.get("gdpr_risk_value", {}),
                    "style[marketplace_id={m_id}][language_tag={lang}]#1.value".format(m_id=meta["m_id"], lang=meta["lang"]): cfg_trans.get("style_name", {}),
                    "department[marketplace_id={m_id}][language_tag={lang}]#1.value".format(m_id=meta["m_id"], lang=meta["lang"]): cfg_trans.get("department_name", {}),
                    "target_gender[marketplace_id={m_id}]#1.value".format(m_id=meta["m_id"]): cfg_trans.get("target_gender", {}),
                    "strap_type[marketplace_id={m_id}][language_tag={lang}]#1.value".format(m_id=meta["m_id"], lang=meta["lang"]): cfg_trans.get("strap_type", {}),
                    "water_resistance_level[marketplace_id={m_id}]#1.value".format(m_id=meta["m_id"]): cfg_trans.get("water_resistance_level", {}),
                    "display[marketplace_id={m_id}]#1.type[language_tag={lang}]#1.value".format(m_id=meta["m_id"], lang=meta["lang"]): cfg_trans.get("display_type", {}),
                    "clasp_type[marketplace_id={m_id}][language_tag={lang}]#1.value".format(m_id=meta["m_id"], lang=meta["lang"]): cfg_trans.get("clasp_type", {}),
                    "calendar_type[marketplace_id={m_id}][language_tag={lang}]#1.value".format(m_id=meta["m_id"], lang=meta["lang"]): cfg_trans.get("calendar_type", {}),
                    "warranty_type[marketplace_id={m_id}][language_tag={lang}]#1.value".format(m_id=meta["m_id"], lang=meta["lang"]): cfg_trans.get("warranty_type", {}),
                    "watch_movement_type[marketplace_id={m_id}][language_tag={lang}]#1.value".format(m_id=meta["m_id"], lang=meta["lang"]): cfg_trans.get("movement_type", {}),
                    "power_source_type[marketplace_id={m_id}][language_tag={lang}]#1.value".format(m_id=meta["m_id"], lang=meta["lang"]): cfg_trans.get("power_source_type", {}),
                    "included_components[marketplace_id={m_id}][language_tag={lang}]#1.value".format(m_id=meta["m_id"], lang=meta["lang"]): cfg_trans.get("included_components", {}),
                    "dial_window[marketplace_id={m_id}]#1.material[language_tag={lang}]#1.value".format(m_id=meta["m_id"], lang=meta["lang"]): cfg_trans.get("dial_window_material_type", {}),
                    "case[marketplace_id={m_id}]#1.material[language_tag={lang}]#1.value".format(m_id=meta["m_id"], lang=meta["lang"]): cfg_trans.get("material_type", {}),
                    "band[marketplace_id={m_id}]#1.material[language_tag={lang}]#1.value".format(m_id=meta["m_id"], lang=meta["lang"]): cfg_trans.get("material_type", {}),
                    "color[marketplace_id={m_id}][language_tag={lang}]#1.standardized_values#1".format(m_id=meta["m_id"], lang=meta["lang"]): cfg_trans.get("color_name", {}),
                    "color[marketplace_id={m_id}][language_tag={lang}]#1.value".format(m_id=meta["m_id"], lang=meta["lang"]): cfg_trans.get("color_name", {}),
                    "dial[marketplace_id={m_id}]#1.color[language_tag={lang}]#1.value".format(m_id=meta["m_id"], lang=meta["lang"]): cfg_trans.get("dial_color", {}),
                    "case[marketplace_id={m_id}]#1.diameter#1.unit".format(m_id=meta["m_id"]): cfg_trans.get("length_units", {}),
                    "water_resistance_depth[marketplace_id={m_id}]#1.unit".format(m_id=meta["m_id"]): cfg_trans.get("length_units", {}),
                    "item_shape[marketplace_id={m_id}][language_tag={lang}]#1.value".format(m_id=meta["m_id"], lang=meta["lang"]): cfg_trans.get("case_shape", {}),
                    "band[marketplace_id={m_id}]#1.color[language_tag={lang}]#1.value".format(m_id=meta["m_id"], lang=meta["lang"]): cfg_trans.get("band_color", {}),
                }

                wb = openpyxl.load_workbook(blank_sc_file, keep_vba=True)
                ws = wb[target_sheet]

                headers = [
                    str(cell.value).strip() if cell.value is not None else ""
                    for cell in ws[5]
                ]
                start_row = 7

                for idx, vc_row in vc_df.iterrows():
                    current_sku = str(vc_row.get("vendor_sku#1.value", "")).strip()
                    if isinstance(current_sku, pd.Series):
                        current_sku = str(current_sku.iloc[0]).strip()

                    if current_sku in ["ABC123", "nan", ""] or "REQUIRED" in current_sku.upper():
                        continue

                    for vc_col, sc_col in current_static_map.items():
                        if vc_col in vc_df.columns:
                            cell_value = vc_row[vc_col]
                            if isinstance(cell_value, pd.Series):
                                cell_value = cell_value.iloc[0]
                            if pd.isna(cell_value):
                                cell_value = ""

                            if isinstance(cell_value, str):
                                cell_value = cell_value.strip()

                            if sc_col in headers:
                                col_idx = headers.index(sc_col) + 1
                                target_cell = ws.cell(row=start_row, column=col_idx)

                                if sc_col in dropdown_translations and cell_value != "":
                                    lookup_dict = dropdown_translations[sc_col]
                                    val_lower = str(cell_value).lower()
                                    matched_local_val = None
                                    for eng_key, local_val in lookup_dict.items():
                                        if eng_key.lower() == val_lower:
                                            matched_local_val = local_val
                                            break
                                    if matched_local_val:
                                        cell_value = matched_local_val

                                target_cell.value = None
                                if sc_col == "amzn1.volt.ca.product_id_value":
                                    target_cell.number_format = "@"
                                    target_cell.value = str(cell_value)
                                else:
                                    target_cell.value = cell_value

                    for sc_col, static_value in current_static_values.items():
                        if sc_col in headers:
                            col_idx = headers.index(sc_col) + 1
                            target_cell = ws.cell(row=start_row, column=col_idx)
                            target_cell.value = static_value

                    start_row += 1

                MAX_ALLOWED_WIDTH = 50
                for col in ws.columns:
                    max_len = 0
                    col_letter = col[0].column_letter
                    for row_idx in range(7, ws.max_row + 1):
                        cell_value = ws.cell(row=row_idx, column=col[0].column).value
                        if cell_value:
                            max_len = max(max_len, len(str(cell_value)))
                    if max_len > 0:
                        calculated_width = max(max_len + 3, 12)
                        ws.column_dimensions[col_letter].width = min(
                            calculated_width, MAX_ALLOWED_WIDTH
                        )
                    else:
                        ws.column_dimensions[col_letter].width = 15

                wb.calculation.calcMode = "auto"

                # Cloud Optimization: Save the workbook to a byte stream instead of local disk
                excel_buffer = io.BytesIO()
                wb.save(excel_buffer)
                excel_buffer.seek(0)

            st.success("🎉 Transformation Completed Successfully!")

            # Pass the byte stream directly to the download button
            st.download_button(
                label=f"📥 Download Processed File ({target_country})",
                data=excel_buffer,
                file_name=output_filename,
                mime="application/vnd.ms-excel.sheet.macroEnabled.12",
            )

        except Exception as e:
            st.error(f"💥 Production pipeline failure: {str(e)}")
else:
    st.info("💡 Drop your completed Vendor Central mapping sheet above to run the generator.")