import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
import math
from datetime import date
import io
import os
import urllib.parse

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

st.set_page_config(page_title="Pharma Estimate & Master System", page_icon="💊", layout="wide")

# Font fallback
FONT_REGULAR = "Helvetica"
FONT_BOLD = "Helvetica-Bold"
win_arial = "C:\\Windows\\Fonts\\arial.ttf"
win_arial_bd = "C:\\Windows\\Fonts\\arialbd.ttf"
if os.path.exists(win_arial) and os.path.exists(win_arial_bd):
    try:
        pdfmetrics.registerFont(TTFont('ArialCustom', win_arial))
        pdfmetrics.registerFont(TTFont('ArialCustomBold', win_arial_bd))
        FONT_REGULAR = 'ArialCustom'
        FONT_BOLD = 'ArialCustomBold'
    except Exception:
        pass

# --- Google Sheets Connection ---
conn = st.connection("gsheets", type=GSheetsConnection)

def get_sheet_data(worksheet_name):
    try:
        df = conn.read(worksheet=worksheet_name, ttl="0s")
        return df.dropna(how="all") if df is not None else pd.DataFrame()
    except Exception:
        return pd.DataFrame()

def save_sheet_data(worksheet_name, df):
    conn.update(worksheet=worksheet_name, data=df)

# --- Master Operations ---
def get_medicines():
    df = get_sheet_data("medicines")
    if df.empty:
        df = pd.DataFrame(columns=["name", "composition", "type", "pack_size", "mrp", "ptr"])
    return df

def upsert_medicine_gsheet(name, composition, med_type, pack_size, mrp, ptr):
    df = get_medicines()
    name_clean = name.strip()
    new_row = {
        "name": name_clean,
        "composition": composition.strip(),
        "type": med_type,
        "pack_size": int(pack_size),
        "mrp": float(mrp),
        "ptr": float(ptr)
    }
    
    if not df.empty and name_clean in df["name"].values:
        df.loc[df["name"] == name_clean, ["composition", "type", "pack_size", "mrp", "ptr"]] = [
            composition.strip(), med_type, int(pack_size), float(mrp), float(ptr)
        ]
    else:
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    
    save_sheet_data("medicines", df)

def get_settings():
    df = get_sheet_data("settings")
    if df.empty:
        return {
            "store_name": "HEALTHCARE PHARMACY & CLINIC",
            "store_address": "Main Market Road",
            "store_contact": "+91 98765 43210"
        }
    return dict(zip(df["key"], df["value"]))

# --- Estimate Save / Update Logic in Google Sheets ---
def save_estimate_to_gsheet(p_name, p_phone, p_age, p_gender, d_name, items, subtotal, disc_type, disc_val, disc_amt, grand_total, existing_id=None):
    df_est = get_sheet_data("estimates")
    df_items = get_sheet_data("estimate_items")
    today_str = date.today().strftime("%d-%m-%Y")

    if existing_id:
        est_id = int(existing_id)
        # Update estimate row
        mask = df_est["id"] == est_id
        df_est.loc[mask, ["patient_name", "patient_phone", "patient_age", "patient_gender", "doctor_name", "estimate_date", "subtotal", "overall_discount_type", "overall_discount_val", "overall_discount_amt", "grand_total"]] = [
            p_name.strip(), str(p_phone).strip(), str(p_age).strip(), p_gender, d_name.strip(), today_str, subtotal, disc_type, disc_val, disc_amt, grand_total
        ]
        # Remove old line items
        if not df_items.empty:
            df_items = df_items[df_items["estimate_id"] != est_id]
    else:
        est_id = 1 if df_est.empty else int(df_est["id"].max()) + 1
        new_est = {
            "id": est_id, "patient_name": p_name.strip(), "patient_phone": str(p_phone).strip(),
            "patient_age": str(p_age).strip(), "patient_gender": p_gender, "doctor_name": d_name.strip(),
            "estimate_date": today_str, "subtotal": subtotal, "overall_discount_type": disc_type,
            "overall_discount_val": disc_val, "overall_discount_amt": disc_amt, "grand_total": grand_total
        }
        df_est = pd.concat([df_est, pd.DataFrame([new_est])], ignore_index=True)

    # Insert Line Items
    new_items = []
    item_start_id = 1 if df_items.empty else int(df_items["id"].max()) + 1
    for idx, itm in enumerate(items):
        new_items.append({
            "id": item_start_id + idx,
            "estimate_id": est_id,
            "medicine_name": itm["name"],
            "regimen": itm["regimen"],
            "timing": itm["timing"],
            "units_required": itm["units_needed"],
            "billing_mode": itm["billing_mode"],
            "billing_qty": itm["billing_qty"],
            "base_amount": itm["base_amount"],
            "discount_type": itm["disc_type"],
            "discount_val": itm["disc_val"],
            "net_amount": itm["net_amount"]
        })
    df_items = pd.concat([df_items, pd.DataFrame(new_items)], ignore_index=True)

    save_sheet_data("estimates", df_est)
    save_sheet_data("estimate_items", df_items)
    return est_id

# --- PDF Generator ---
def generate_pdf_estimate(p_name, p_phone, p_age, p_gender, d_name, items, subtotal, disc_type, disc_val, disc_amt, grand_total, est_number=None):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []
    settings = get_settings()

    store_title_style = ParagraphStyle('StoreTitle', fontName=FONT_BOLD, fontSize=16, leading=20, textColor=colors.HexColor("#1A365D"), alignment=1)
    store_sub_style = ParagraphStyle('StoreSub', fontName=FONT_REGULAR, fontSize=9, leading=12, textColor=colors.HexColor("#4A5568"), alignment=1)
    
    elements.append(Paragraph(settings.get('store_name', 'PHARMACY ESTIMATE'), store_title_style))
    elements.append(Paragraph(f"{settings.get('store_address', '')} | Ph: {settings.get('store_contact', '')}", store_sub_style))
    elements.append(Spacer(1, 10))
    
    title_style = ParagraphStyle('EstTitle', fontName=FONT_BOLD, fontSize=12, leading=15, textColor=colors.HexColor("#2B6CB0"), alignment=1)
    elements.append(Paragraph("MEDICINE ESTIMATE & PRESCRIPTION BREAKDOWN", title_style))
    elements.append(Spacer(1, 12))

    today_str = date.today().strftime("%d-%m-%Y")
    est_label = f"Estimate #{est_number}" if est_number else "Provisional Estimate"
    
    meta_p = ParagraphStyle('Meta', fontName=FONT_REGULAR, fontSize=9, leading=13)
    p_info = f"<b>Patient:</b> {p_name or 'N/A'}"
    if p_age or p_gender:
        p_info += f" ({p_age} yrs / {p_gender})"
    if p_phone:
        p_info += f" | Mob: {p_phone}"

    meta_data = [
        [Paragraph(p_info, meta_p), Paragraph(f"<b>Date:</b> {today_str}", meta_p)],
        [Paragraph(f"<b>Doctor:</b> {d_name or 'N/A'}", meta_p), Paragraph(f"<b>Ref No:</b> {est_label}", meta_p)]
    ]
    meta_table = Table(meta_data, colWidths=[360, 175])
    meta_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    elements.append(meta_table)
    elements.append(Spacer(1, 12))

    table_data = [["Sr.", "Medicine & Salt", "Dosage / Instructions", "Billing Units", "Base (₹)", "Disc", "Net (₹)"]]
    for idx, itm in enumerate(items, 1):
        med_label = f"<b>{itm['name']}</b>"
        if itm.get('composition'):
            med_label += f"<br/><font size=7 color='#64748B'>{itm['composition']}</font>"
        
        dose_label = f"{itm.get('regimen', '')}"
        if itm.get('timing') and itm.get('timing') != "None":
            dose_label += f"<br/><font size=7 color='#2563EB'>{itm['timing']}</font>"
        dose_label += f" ({itm.get('days', 1)} days)"

        disc_str = f"{itm.get('disc_val', 0):.0f}%" if itm.get('disc_type') == "%" else f"₹{itm.get('disc_val', 0):.0f}"
        if itm.get('disc_val', 0) == 0:
            disc_str = "-"

        table_data.append([
            str(idx),
            Paragraph(med_label, meta_p),
            Paragraph(dose_label, meta_p),
            itm.get('billing_qty', ''),
            f"₹{itm.get('base_amount', 0.0):.2f}",
            disc_str,
            f"₹{itm.get('net_amount', 0.0):.2f}"
        ])

    table_data.append(["", "", "", "", "", "Subtotal:", f"₹{subtotal:.2f}"])
    if disc_amt > 0:
        disc_label = f"Bill Disc ({disc_val:.0f}%):" if disc_type == "%" else "Bill Discount (₹):"
        table_data.append(["", "", "", "", "", disc_label, f"-₹{disc_amt:.2f}"])
    table_data.append(["", "", "", "", "", "Grand Total:", f"₹{grand_total:.2f}"])

    main_table = Table(table_data, colWidths=[24, 180, 120, 75, 50, 36, 50])
    main_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#2B6CB0")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('FONTNAME', (0, 0), (-1, -1), FONT_REGULAR),
        ('FONTNAME', (0, 0), (-1, 0), FONT_BOLD),
        ('FONTSIZE', (0, 0), (-1, -1), 8.5),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (1, 1), (2, -1), 'LEFT'),
        ('ALIGN', (4, 1), (-1, -1), 'RIGHT'),
        ('GRID', (0, 0), (-1, len(items)), 0.5, colors.HexColor("#CBD5E1")),
        ('FONTNAME', (-2, -3), (-1, -1), FONT_BOLD),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    elements.append(main_table)
    doc.build(elements)
    buffer.seek(0)
    return buffer

REGIMEN_MAP = {
    "1-0-1 (Twice a day / BD)": 2.0,
    "1-1-1 (Thrice a day / TDS)": 3.0,
    "1-0-0 (Morning only / OD)": 1.0,
    "0-0-1 (Night only / HS)": 1.0,
    "1-1-1-1 (Four times / QID)": 4.0,
    "0.5-0-0.5 (Half tablet twice)": 1.0,
    "Custom / SOS": 1.0
}

# --- State Management ---
if "current_estimate" not in st.session_state:
    st.session_state.current_estimate = []
if "edit_index" not in st.session_state:
    st.session_state.edit_index = None
if "loaded_estimate_id" not in st.session_state:
    st.session_state.loaded_estimate_id = None
if "patient_input" not in st.session_state:
    st.session_state.patient_input = ""
if "phone_input" not in st.session_state:
    st.session_state.phone_input = ""
if "age_input" not in st.session_state:
    st.session_state.age_input = ""
if "gender_input" not in st.session_state:
    st.session_state.gender_input = "Male"
if "doctor_input" not in st.session_state:
    st.session_state.doctor_input = ""
if "overall_disc_type" not in st.session_state:
    st.session_state.overall_disc_type = "%"
if "overall_disc_val" not in st.session_state:
    st.session_state.overall_disc_val = 0.0

tab_estimate, tab_master, tab_history = st.tabs([
    "📋 Create / Edit Estimate", 
    "💊 Medicine Master (Google Sheet)", 
    "📑 Saved Estimates"
])

# -------------------------------------------------------------
# TAB 1: CREATE / EDIT ESTIMATE
# -------------------------------------------------------------
with tab_estimate:
    if st.session_state.loaded_estimate_id:
        st.warning(f"✏️ **Editing Saved Estimate #{st.session_state.loaded_estimate_id}** from Google Sheets.")

    st.markdown("##### 👤 Patient & Doctor Details")
    p1, p2, p3, p4, p5 = st.columns([1.5, 1.2, 0.8, 0.8, 1.4])
    with p1:
        patient_name = st.text_input("Patient Name", value=st.session_state.patient_input, placeholder="e.g. Ramesh Kumar")
        st.session_state.patient_input = patient_name
    with p2:
        patient_phone = st.text_input("Mobile No (for WhatsApp)", value=st.session_state.phone_input, placeholder="e.g. 9876543210")
        st.session_state.phone_input = patient_phone
    with p3:
        patient_age = st.text_input("Age", value=st.session_state.age_input, placeholder="e.g. 45")
        st.session_state.age_input = patient_age
    with p4:
        gender_opts = ["Male", "Female", "Other"]
        g_idx = gender_opts.index(st.session_state.gender_input) if st.session_state.gender_input in gender_opts else 0
        patient_gender = st.selectbox("Gender", gender_opts, index=g_idx)
        st.session_state.gender_input = patient_gender
    with p5:
        doctor_name = st.text_input("Doctor Name", value=st.session_state.doctor_input, placeholder="e.g. Dr. S. Sharma")
        st.session_state.doctor_input = doctor_name

    st.markdown("---")
    col_input, col_view = st.columns([1.2, 1.4], gap="large")

    with col_input:
        if (st.session_state.edit_index is not None and 0 <= st.session_state.edit_index < len(st.session_state.current_estimate)):
            is_editing = True
            edit_item = st.session_state.current_estimate[st.session_state.edit_index]
        else:
            is_editing = False
            edit_item = None
            st.session_state.edit_index = None

        st.subheader("Edit Medicine" if is_editing else "Add Medicine")
        entry_method = st.radio("Selection Method", ["Search / Select Master", "Manual Entry"], horizontal=True, disabled=is_editing)
        med_df = get_medicines()

        if entry_method == "Search / Select Master" and not is_editing:
            if not med_df.empty:
                med_df["display_search"] = med_df["name"] + " | " + med_df["composition"].fillna("")
                selected_display = st.selectbox("Search by Brand or Composition/Salt", med_df["display_search"].tolist())
                med_info = med_df[med_df["display_search"] == selected_display].iloc[0]
                
                name_val = med_info["name"]
                comp_val = str(med_info["composition"]) if pd.notna(med_info["composition"]) else ""
                type_val = med_info["type"]
                pack_val = int(med_info["pack_size"])
                mrp_val = float(med_info["mrp"])
                ptr_val = float(med_info.get("ptr", 0.0))
            else:
                name_val, comp_val, type_val, pack_val, mrp_val, ptr_val = "", "", "Tablet", 10, 100.0, 0.0
        else:
            name_val = edit_item['name'] if is_editing else ""
            comp_val = edit_item.get('composition', '') if is_editing else ""
            type_val = edit_item['type'] if is_editing else "Tablet"
            pack_val = edit_item['pack_size'] if is_editing else 10
            mrp_val = edit_item['mrp'] if is_editing else 100.0
            ptr_val = edit_item.get('ptr', 0.0) if is_editing else 0.0

        i_name = st.text_input("Medicine Brand Name", value=name_val)
        i_comp = st.text_input("Composition / Salt", value=comp_val, placeholder="e.g. Amoxicillin + Clavulanic Acid")
        
        c_t1, c_t2 = st.columns(2)
        with c_t1:
            types = ["Tablet", "Capsule", "Syrup", "Injection", "Ointment", "Drops"]
            i_type = st.selectbox("Form", types, index=types.index(type_val) if type_val in types else 0)
        with c_t2:
            i_pack = st.number_input("Pack Size", min_value=1, value=pack_val, step=1)
        
        c_m1, c_m2 = st.columns(2)
        with c_m1:
            i_mrp = st.number_input("Pack MRP (₹)", min_value=0.0, value=mrp_val, step=0.5)
        with c_m2:
            i_ptr = st.number_input("PTR / Net Purchase Cost (₹)", min_value=0.0, value=ptr_val, step=0.5)

        st.markdown("**Dosage & Regimen**")
        d_c1, d_c2 = st.columns([1.6, 1])
        with d_c1:
            regimen_keys = list(REGIMEN_MAP.keys())
            saved_reg = edit_item.get('regimen', regimen_keys[0]) if is_editing else regimen_keys[0]
            sel_regimen = st.selectbox("Prescription Code", regimen_keys, index=regimen_keys.index(saved_reg) if saved_reg in regimen_keys else 0)
        with d_c2:
            if sel_regimen == "Custom / SOS":
                custom_doses = st.number_input("Doses/Day", min_value=0.5, value=float(edit_item['freq']) if is_editing else 1.0, step=0.5)
            else:
                custom_doses = REGIMEN_MAP[sel_regimen]
                st.write(f"Doses: **{custom_doses}/day**")

        t_c1, t_c2 = st.columns([1.2, 1.2])
        with t_c1:
            timings = ["None", "After Food (PC)", "Before Food (AC)", "With Food", "Empty Stomach", "At Bedtime"]
            saved_time = edit_item.get('timing', 'None') if is_editing else "None"
            sel_timing = st.selectbox("Meal Timing", timings, index=timings.index(saved_time) if saved_time in timings else 0)
        with t_c2:
            i_days = st.number_input("Course Days", min_value=1, value=int(edit_item['days']) if is_editing else 5, step=1)

        total_units = math.ceil(custom_doses * i_days)
        st.caption(f"Calculated Total Consumption: **{total_units} units**")

        b_c1, b_c2 = st.columns([1.3, 1])
        with b_c1:
            mode_opts = ["Full Pack Rounding", "Allow Cutting / Loose Units"]
            saved_mode = 0 if not is_editing or edit_item['billing_mode'] == "Full Pack Rounding" else 1
            i_mode = st.radio("Billing Mode", mode_opts, index=saved_mode, horizontal=True)
        with b_c2:
            disc_col1, disc_col2 = st.columns([1, 1.2])
            with disc_col1:
                saved_it_dt = edit_item.get('disc_type', '%') if is_editing else "%"
                item_disc_type = st.selectbox("Type", ["%", "₹"], index=0 if saved_it_dt == "%" else 1, key="item_dt")
            with disc_col2:
                saved_it_val = float(edit_item.get('disc_val', 0.0)) if is_editing else 0.0
                item_disc_val = st.number_input("Item Disc", min_value=0.0, value=saved_it_val, step=1.0)

        btn_c1, btn_c2 = st.columns(2)
        with btn_c1:
            lbl = "Update Medicine" if is_editing else "Add to Estimate"
            if st.button(lbl, type="primary", use_container_width=True):
                if not i_name.strip():
                    st.error("Please provide a medicine name.")
                else:
                    # Auto save / update in Google Sheets Master
                    upsert_medicine_gsheet(i_name, i_comp, i_type, i_pack, i_mrp, i_ptr)
                    
                    unit_rate = i_mrp / i_pack if i_pack > 0 else 0
                    if i_mode == "Full Pack Rounding":
                        packs = math.ceil(total_units / i_pack)
                        base_cost = packs * i_mrp
                        billing_label = f"{packs} Pack(s)"
                        ptr_cost = packs * i_ptr
                    else:
                        base_cost = total_units * unit_rate
                        billing_label = f"{total_units} Unit(s)"
                        ptr_cost = total_units * (i_ptr / i_pack if i_pack > 0 else 0)

                    disc_amount = (base_cost * item_disc_val / 100.0) if item_disc_type == "%" else min(item_disc_val, base_cost)
                    net_cost = max(0.0, base_cost - disc_amount)

                    payload = {
                        "name": i_name.strip(),
                        "composition": i_comp.strip(),
                        "type": i_type,
                        "pack_size": i_pack,
                        "mrp": i_mrp,
                        "ptr": i_ptr,
                        "freq": custom_doses,
                        "regimen": sel_regimen,
                        "timing": sel_timing,
                        "days": i_days,
                        "units_needed": total_units,
                        "billing_mode": i_mode,
                        "billing_qty": billing_label,
                        "base_amount": round(base_cost, 2),
                        "ptr_amount": round(ptr_cost, 2),
                        "disc_type": item_disc_type,
                        "disc_val": item_disc_val,
                        "net_amount": round(net_cost, 2)
                    }

                    if is_editing:
                        st.session_state.current_estimate[st.session_state.edit_index] = payload
                        st.session_state.edit_index = None
                    else:
                        st.session_state.current_estimate.append(payload)
                    st.rerun()

        with btn_c2:
            if is_editing and st.button("Cancel Edit", use_container_width=True):
                st.session_state.edit_index = None
                st.rerun()

    # Right Column: Estimate Summary
    with col_view:
        st.subheader("Estimate Summary & Actions")
        if not st.session_state.current_estimate:
            st.info("Prescription is empty. Add medicines from the left panel.")
        else:
            subtotal = 0.0
            total_ptr_cost = 0.0

            for idx, itm in enumerate(st.session_state.current_estimate):
                subtotal += itm['net_amount']
                total_ptr_cost += itm.get('ptr_amount', 0.0)
                
                with st.container():
                    r1, r2, r3 = st.columns([2.5, 1.2, 0.8])
                    with r1:
                        st.markdown(f"**{idx + 1}. {itm['name']}** ({itm['billing_qty']})")
                        time_str = f" • {itm['timing']}" if itm.get('timing') and itm['timing'] != "None" else ""
                        disc_str = f" • Disc: {itm['disc_val']}{itm['disc_type']}" if itm.get('disc_val', 0) > 0 else ""
                        st.caption(f"{itm['regimen']} for {itm['days']} days{time_str}{disc_str}")
                    with r2:
                        st.markdown(f"**₹{itm['net_amount']:.2f}**")
                    with r3:
                        b_edit, b_del = st.columns(2)
                        with b_edit:
                            if st.button("✏️", key=f"edit_{idx}"):
                                st.session_state.edit_index = idx
                                st.rerun()
                        with b_del:
                            if st.button("🗑️", key=f"del_{idx}"):
                                st.session_state.current_estimate.pop(idx)
                                if st.session_state.edit_index == idx:
                                    st.session_state.edit_index = None
                                elif st.session_state.edit_index is not None and st.session_state.edit_index > idx:
                                    st.session_state.edit_index -= 1
                                st.rerun()
                    st.divider()

            # Bill Discount Logic
            c_disc1, c_disc2, c_tot = st.columns([0.8, 1.1, 1.4])
            with c_disc1:
                disc_opts = ["%", "₹"]
                dt_index = 0 if st.session_state.overall_disc_type == "%" else 1
                overall_disc_type = st.selectbox("Discount Type", disc_opts, index=dt_index, key="bill_dt_select")
                st.session_state.overall_disc_type = overall_disc_type

            with c_disc2:
                step_val = 1.0 if overall_disc_type == "%" else 5.0
                overall_disc_val = st.number_input(
                    f"Bill Discount ({overall_disc_type})", 
                    min_value=0.0, 
                    value=float(st.session_state.overall_disc_val), 
                    step=step_val, 
                    key="bill_dv_input"
                )
                st.session_state.overall_disc_val = overall_disc_val

            if overall_disc_type == "₹":
                overall_disc_amt = min(float(overall_disc_val), float(subtotal))
            else:
                overall_disc_amt = (float(subtotal) * float(overall_disc_val)) / 100.0

            grand_total = max(0.0, float(subtotal) - float(overall_disc_amt))

            with c_tot:
                st.write(f"Subtotal: **₹{subtotal:,.2f}**")
                if overall_disc_amt > 0:
                    disc_label = f"{overall_disc_val:.2f}%" if overall_disc_type == "%" else f"₹{overall_disc_val:.2f}"
                    st.write(f"Discount ({disc_label}): **-₹{overall_disc_amt:,.2f}**")
                st.markdown(f"### Grand Total: ₹{grand_total:,.2f}")

            # PDF & WhatsApp Share
            pdf_buf = generate_pdf_estimate(
                patient_name, patient_phone, patient_age, patient_gender, doctor_name,
                st.session_state.current_estimate, subtotal, overall_disc_type, overall_disc_val,
                overall_disc_amt, grand_total, est_number=st.session_state.loaded_estimate_id
            )

            wa_text = f"*MEDICINE ESTIMATE*\nPatient: {patient_name or 'Valued Customer'}\n"
            for idx, itm in enumerate(st.session_state.current_estimate, 1):
                wa_text += f"{idx}. {itm['name']} ({itm['billing_qty']}) - ₹{itm['net_amount']:.2f}\n"
            wa_text += f"\n*Total Amount Payable: ₹{grand_total:.2f}*"
            encoded_text = urllib.parse.quote(wa_text)
            
            clean_phone = "".join(filter(str.isdigit, str(patient_phone)))
            if len(clean_phone) == 10:
                clean_phone = "91" + clean_phone
            wa_url = f"https://wa.me/{clean_phone}?text={encoded_text}" if clean_phone else f"https://wa.me/?text={encoded_text}"

            st.markdown("---")
            a1, a2 = st.columns(2)
            with a1:
                st.download_button(
                    label="📄 Download Estimate PDF",
                    data=pdf_buf,
                    file_name=f"Estimate_{patient_name or 'Patient'}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            with a2:
                st.link_button("📲 Share on WhatsApp", url=wa_url, use_container_width=True)

            st.write("")
            if st.session_state.loaded_estimate_id:
                s1, s2, s3 = st.columns(3)
                with s1:
                    if st.button("💾 Update in Google Sheet", type="primary", use_container_width=True):
                        with st.spinner("Updating Google Sheet..."):
                            save_estimate_to_gsheet(
                                patient_name, patient_phone, patient_age, patient_gender, doctor_name,
                                st.session_state.current_estimate, subtotal, overall_disc_type,
                                overall_disc_val, overall_disc_amt, grand_total,
                                existing_id=st.session_state.loaded_estimate_id
                            )
                        st.success("Estimate updated in Google Sheets!")
                        st.rerun()
                with s2:
                    if st.button("📄 Save as New Copy", use_container_width=True):
                        with st.spinner("Saving new record..."):
                            nid = save_estimate_to_gsheet(
                                patient_name, patient_phone, patient_age, patient_gender, doctor_name,
                                st.session_state.current_estimate, subtotal, overall_disc_type,
                                overall_disc_val, overall_disc_amt, grand_total
                            )
                        st.success(f"Saved as new Estimate #{nid}!")
                        st.session_state.loaded_estimate_id = nid
                        st.rerun()
                with s3:
                    if st.button("Reset Form", use_container_width=True):
                        st.session_state.current_estimate = []
                        st.session_state.loaded_estimate_id = None
                        st.rerun()
            else:
                s1, s2 = st.columns([1.5, 1])
                with s1:
                    if st.button("💾 Save to Google Sheet", type="primary", use_container_width=True):
                        if not patient_name.strip():
                            st.error("Please enter Patient Name before saving.")
                        else:
                            with st.spinner("Saving to Google Sheet..."):
                                nid = save_estimate_to_gsheet(
                                    patient_name, patient_phone, patient_age, patient_gender, doctor_name,
                                    st.session_state.current_estimate, subtotal, overall_disc_type,
                                    overall_disc_val, overall_disc_amt, grand_total
                                )
                            st.success(f"Estimate #{nid} saved successfully in Google Sheet!")
                            st.session_state.current_estimate = []
                            st.session_state.patient_input = ""
                            st.session_state.phone_input = ""
                            st.session_state.age_input = ""
                            st.session_state.doctor_input = ""
                            st.session_state.overall_disc_val = 0.0
                            st.rerun()
                with s2:
                    if st.button("Clear All", use_container_width=True):
                        st.session_state.current_estimate = []
                        st.session_state.edit_index = None
                        st.rerun()

# -------------------------------------------------------------
# TAB 2: MEDICINE MASTER (GOOGLE SHEET)
# -------------------------------------------------------------
with tab_master:
    st.subheader("Medicine Master Database (Live Google Sheet)")
    c1, c2 = st.columns([1, 1.4], gap="large")
    with c1:
        st.markdown("##### Add / Update Master Item")
        mm_name = st.text_input("Medicine Brand", key="mm_name")
        mm_comp = st.text_input("Composition / Salt", key="mm_comp")
        mm_types = ["Tablet", "Capsule", "Syrup", "Injection", "Ointment", "Drops"]
        mm_type = st.selectbox("Form", mm_types, key="mm_type")
        mm_pack = st.number_input("Pack Size", min_value=1, value=10, step=1, key="mm_pack")
        mm_mrp = st.number_input("Pack MRP (₹)", min_value=0.0, value=100.0, step=0.5, key="mm_mrp")
        mm_ptr = st.number_input("PTR (₹)", min_value=0.0, value=75.0, step=0.5, key="mm_ptr")

        if st.button("Save to Google Sheet", use_container_width=True):
            if not mm_name.strip():
                st.error("Medicine Name is required.")
            else:
                with st.spinner("Updating Google Sheet..."):
                    upsert_medicine_gsheet(mm_name, mm_comp, mm_type, mm_pack, mm_mrp, mm_ptr)
                st.success(f"Saved/Updated '{mm_name}' in Google Sheet.")
                st.rerun()

    with c2:
        st.markdown("##### Current Catalog in Google Sheet")
        df_meds = get_medicines()
        if not df_meds.empty:
            st.dataframe(
                df_meds[["name", "composition", "type", "pack_size", "mrp", "ptr"]],
                hide_index=True,
                use_container_width=True
            )
        else:
            st.info("No medicines found in Google Sheet.")

# -------------------------------------------------------------
# TAB 3: SAVED ESTIMATES & LOAD FOR EDITING
# -------------------------------------------------------------
with tab_history:
    st.subheader("Saved Estimates in Google Sheets")
    estimates_df = get_sheet_data("estimates")
    
    if estimates_df.empty:
        st.info("No saved estimates in Google Sheet.")
    else:
        disp_cols = [c for c in ["id", "patient_name", "patient_phone", "doctor_name", "estimate_date", "subtotal", "grand_total"] if c in estimates_df.columns]
        st.dataframe(
            estimates_df[disp_cols].sort_values(by="id", ascending=False),
            hide_index=True,
            use_container_width=True
        )
        
        sel_id = st.selectbox("Select Estimate # to inspect, load & edit", estimates_df["id"].tolist())
        if sel_id:
            row = estimates_df[estimates_df["id"] == sel_id].iloc[0]
            all_items_df = get_sheet_data("estimate_items")
            items_df = all_items_df[all_items_df["estimate_id"] == sel_id] if not all_items_df.empty else pd.DataFrame()
            
            st.write(f"**Items in Estimate #{sel_id}**")
            if not items_df.empty:
                summary_cols = [c for c in ["medicine_name", "regimen", "timing", "units_required", "billing_mode", "billing_qty", "base_amount", "discount_val", "net_amount"] if c in items_df.columns]
                st.dataframe(items_df[summary_cols], hide_index=True, use_container_width=True)
            
            c_l, c_p = st.columns(2)
            with c_l:
                if st.button("✏️ Load & Edit this Estimate", type="primary", use_container_width=True):
                    meds_master = get_medicines().set_index("name").to_dict(orient="index")
                    loaded = []
                    for _, r in items_df.iterrows():
                        m_name = r["medicine_name"]
                        m_data = meds_master.get(m_name, {"composition": "", "type": "Tablet", "pack_size": 10, "mrp": 100.0, "ptr": 0.0})
                        
                        reg_val = r.get("regimen", "") or "1-0-1 (Twice a day / BD)"
                        time_val = r.get("timing", "") or "None"
                        units_req = int(r.get("units_required", 10))
                        
                        loaded.append({
                            "name": m_name,
                            "composition": m_data.get("composition", ""),
                            "type": m_data.get("type", "Tablet"),
                            "pack_size": m_data.get("pack_size", 10),
                            "mrp": m_data.get("mrp", 100.0),
                            "ptr": m_data.get("ptr", 0.0),
                            "freq": REGIMEN_MAP.get(reg_val, 2.0),
                            "regimen": reg_val,
                            "timing": time_val,
                            "days": max(1, int(units_req / 2)),
                            "units_needed": units_req,
                            "billing_mode": r.get("billing_mode", "Full Pack Rounding"),
                            "billing_qty": r.get("billing_qty", ""),
                            "base_amount": float(r.get("base_amount", 0.0)),
                            "ptr_amount": 0.0,
                            "disc_type": r.get("discount_type", "%") or "%",
                            "disc_val": float(r.get("discount_val", 0.0)),
                            "net_amount": float(r.get("net_amount", 0.0))
                        })
                    st.session_state.current_estimate = loaded
                    st.session_state.patient_input = str(row.get("patient_name", ""))
                    st.session_state.phone_input = str(row.get("patient_phone", ""))
                    st.session_state.age_input = str(row.get("patient_age", ""))
                    st.session_state.gender_input = str(row.get("patient_gender", "Male"))
                    st.session_state.doctor_input = str(row.get("doctor_name", ""))
                    st.session_state.overall_disc_type = str(row.get("overall_discount_type", "%") or "%")
                    st.session_state.overall_disc_val = float(row.get("overall_discount_val", 0.0))
                    st.session_state.loaded_estimate_id = int(sel_id)
                    st.success(f"Estimate #{sel_id} loaded from Google Sheets! Switch to 'Create / Edit Estimate' tab.")
                    st.rerun()

            with c_p:
                hist_items = []
                for _, r in items_df.iterrows():
                    hist_items.append({
                        "name": r.get("medicine_name", ""),
                        "composition": "",
                        "regimen": r.get("regimen", "") or "",
                        "timing": r.get("timing", "") or "",
                        "days": max(1, int(r.get("units_required", 10) / 2)),
                        "units_needed": r.get("units_required", 10),
                        "billing_qty": r.get("billing_qty", ""),
                        "base_amount": float(r.get("base_amount", 0.0)),
                        "disc_type": r.get("discount_type", "%") or "%",
                        "disc_val": float(r.get("discount_val", 0.0)),
                        "net_amount": float(r.get("net_amount", 0.0))
                    })
                
                h_pdf = generate_pdf_estimate(
                    str(row.get("patient_name", "")), str(row.get("patient_phone", "")), 
                    str(row.get("patient_age", "")), str(row.get("patient_gender", "")), 
                    str(row.get("doctor_name", "")), hist_items,
                    float(row.get("subtotal", 0.0)), str(row.get("overall_discount_type", "%")), 
                    float(row.get("overall_discount_val", 0.0)), float(row.get("overall_discount_amt", 0.0)), 
                    float(row.get("grand_total", 0.0)), est_number=sel_id
                )
                
                st.download_button(
                    label=f"📄 Download PDF (#{sel_id})",
                    data=h_pdf,
                    file_name=f"Estimate_{sel_id}_{row.get('patient_name', 'Patient')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
