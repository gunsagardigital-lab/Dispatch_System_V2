import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
import openpyxl
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import webbrowser
import threading

# --- Optional Google Sheets Support ---
try:
    import gspread
    GOOGLE_SHEETS_AVAILABLE = True
except ImportError:
    GOOGLE_SHEETS_AVAILABLE = False

# --- Page Config ---
st.set_page_config(page_title="Dispatch System", page_icon="🚛", layout="wide")

# --- Google Sheet Base URL & Cloud Dynamic Paths ---
SHEET_BASE_URL = "https://docs.google.com/spreadsheets/d/1Rajn2oci_FNlzKXlnf7qo5JKH-JCznwzXUf7WlwQXl0/export?format=csv"

BASE_DIR = os.path.dirname(__file__)
REPORTS_DIR = os.path.join(BASE_DIR, "Reports")
NOTICE_TXT_FILE = os.path.join(BASE_DIR, "notice.txt")
PASSWORD_FILE = os.path.join(BASE_DIR, "password.txt")
CREDENTIALS_PATH = os.path.join(BASE_DIR, "credentials.json")

if not os.path.exists(REPORTS_DIR):
    os.makedirs(REPORTS_DIR)

# --- Register Font for PDF ---
FONT_NAME = "Helvetica"
FONT_BOLD_NAME = "Helvetica-Bold"
try:
    font_path = "C:\\Windows\\Fonts\\cambria.ttc"
    font_bold_path = "C:\\Windows\\Fonts\\cambriab.ttf"
    if os.path.exists(font_path):
        pdfmetrics.registerFont(TTFont("Cambria", font_path, subfontIndex=0))
        FONT_NAME = "Cambria"
    if os.path.exists(font_bold_path):
        pdfmetrics.registerFont(TTFont("Cambria-Bold", font_bold_path))
        FONT_BOLD_NAME = "Cambria-Bold"
except Exception:
    pass

def get_shift_date():
    now = datetime.now()
    if now.hour < 6:
        return (now - timedelta(days=1)).strftime("%d.%m.%Y")
    return now.strftime("%d.%m.%Y")

def get_notice_from_txt():
    if os.path.exists(NOTICE_TXT_FILE):
        try:
            with open(NOTICE_TXT_FILE, "r", encoding="utf-8") as f:
                content = f.read().strip()
                if content: return content
        except: pass
    return "😊 राधे - राधे 🚨 आवश्यक सूचना:- कृपया सभी गाड़ियों की लोडिंग समय पर पूरी करें!"

def save_notice_to_txt(new_notice):
    try:
        with open(NOTICE_TXT_FILE, "w", encoding="utf-8") as f:
            f.write(new_notice)
        return True
    except:
        return False

# --- Google Sheets Data Loader (Similar to Mobile Dashboard) ---
@st.cache_data(ttl=10)
def load_data_from_gsheet(sheet_name="Sheet1"):
    try:
        url = f"{SHEET_BASE_URL}&sheet={sheet_name}"
        df = pd.read_csv(url, header=2)
        return df
    except Exception:
        try:
            df = pd.read_csv(SHEET_BASE_URL, header=2)
            return df
        except Exception:
            return pd.DataFrame()

# --- Google Sheets Sync & Update Helper ---
def sync_row_to_google_sheet(sheet_name, row_data):
    if not GOOGLE_SHEETS_AVAILABLE or not os.path.exists(CREDENTIALS_PATH):
        return
    def background_sync():
        try:
            gc = gspread.service_account(filename=CREDENTIALS_PATH)
            sh = gc.open("Dispatch_Entry_Register")
            try:
                worksheet = sh.worksheet(sheet_name)
            except gspread.exceptions.WorksheetNotFound:
                worksheet = sh.add_worksheet(title=sheet_name, rows=100, cols=12)
                headers = [
                    "Sr. No.", "In Date", "In Time", "Program No.", "Vehicle No.",
                    "Transport Name", "Destinations", "Loading Plan", "ADV",
                    "Actual Loading", "Current Status", "Remarks (Shift)"
                ]
                worksheet.append_row(headers)
            
            cleaned_row = ["" if v is None else str(v) for v in row_data]
            veh_no = cleaned_row[4]
            
            cell = None
            if veh_no:
                try:
                    cell = worksheet.find(veh_no)
                except:
                    pass
            
            if cell:
                row_idx = cell.row
                for col_idx, val in enumerate(cleaned_row, 1):
                    worksheet.update_cell(row_idx, col_idx, val)
            else:
                worksheet.append_row(cleaned_row)
                all_rows = worksheet.get_all_values()
                row_idx = len(all_rows)
            
            try:
                worksheet.format(f"A{row_idx}:L{row_idx}", {
                    "textFormat": {
                        "fontFamily": "Cambria",
                        "fontSize": 14
                    }
                })
                worksheet.update_row_height(row_idx, 29)
            except Exception:
                pass
                
        except Exception as e:
            print(f"Google Sheet Sync Error: {e}")

    threading.Thread(target=background_sync, daemon=True).start()

# --- PDF Generation Functions ---
def open_pdf_safely(pdf_filename):
    if os.path.exists(pdf_filename):
        try:
            webbrowser.open(os.path.abspath(pdf_filename))
        except Exception:
            pass

def generate_loader_list_pdf(target_date):
    try:
        df = load_data_from_gsheet(target_date)
        if df.empty: return
        
        headers = ["Sr. No.", "In Time", "Vehicle No.", "Destination", "Plan (Tons)", "Prog No."]
        table_data = [headers]
        sr_counter = 1

        for _, row in df.iterrows():
            row_list = list(row)
            if len(row_list) >= 11 and pd.notna(row_list[4]):
                status = str(row_list[10]).strip().lower() if row_list[10] is not None else ""
                if status != "final":
                    raw_time = str(row_list[2]).strip() if row_list[2] is not None else ""
                    clean_time = raw_time[:5] if len(raw_time) >= 5 else raw_time
                    table_data.append([str(sr_counter), clean_time, str(row_list[4]), str(row_list[6]), str(row_list[7]), str(row_list[3])])
                    sr_counter += 1

        pdf_filename = os.path.join(REPORTS_DIR, f"Loader_List_{target_date}.pdf")
        doc = SimpleDocTemplate(pdf_filename, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle("TitleStyle", parent=styles["Heading1"], fontName=FONT_BOLD_NAME, fontSize=15, textColor=colors.HexColor("#1F4E78"), alignment=1, spaceAfter=10)
        
        t = Table(table_data, colWidths=[55, 80, 150, 210, 100, 180])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        doc.build([Paragraph(f"M/s Surya Roshni Limited - Hindupur<br/><b>LOADER DISPATCH LIST ({target_date})</b>", title_style), Spacer(1, 5), t])
        open_pdf_safely(pdf_filename)
        st.success("Loader List PDF जनरेट हो गई!")
    except Exception as e:
        st.error(f"Error: {e}")

def generate_status_pdf(selected_status, target_date):
    try:
        df = load_data_from_gsheet(target_date)
        if df.empty: return
        
        headers = ["Sr.", "In Date", "Time", "Prog No.", "Vehicle No.", "Transport Name", "Destination", "Plan", "ADV", "Actual", "Status", "Remarks"]
        table_data = [headers]
        sr_counter = 1

        for _, row in df.iterrows():
            row_list = list(row)
            if len(row_list) >= 11 and pd.notna(row_list[10]):
                if str(row_list[10]).strip().lower() == selected_status.lower():
                    row_cleaned = list(row_list[:12])
                    row_cleaned[0] = sr_counter
                    sr_counter += 1
                    table_data.append([str(cell) if pd.notna(cell) else "" for cell in row_cleaned])

        pdf_filename = os.path.join(REPORTS_DIR, f"Dispatch_{selected_status.replace(' ', '_')}_{target_date}.pdf")
        doc = SimpleDocTemplate(pdf_filename, pagesize=landscape(A4), rightMargin=15, leftMargin=15, topMargin=20, bottomMargin=20)
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle("TitleStyle", parent=styles["Heading1"], fontName=FONT_BOLD_NAME, fontSize=15, textColor=colors.HexColor("#1F4E78"), alignment=1, spaceAfter=12)
        
        t = Table(table_data, colWidths=[30, 65, 55, 65, 95, 110, 110, 50, 45, 55, 70, 55])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        doc.build([Paragraph(f"M/s Surya Roshni Limited - Hindupur ({selected_status} Report - {target_date})", title_style), Spacer(1, 5), t])
        open_pdf_safely(pdf_filename)
        st.success(f"{selected_status} Report PDF जनरेट हो गई!")
    except Exception as e:
        st.error(f"Error: {e}")

# --- Navigation Sidebar ---
page = st.sidebar.radio("📋 Navigation", ["🚛 Live Dashboard", "🔐 Admin Panel"])

# ==================== LIVE DASHBOARD ====================
if page == "🚛 Live Dashboard":
    st.markdown("""
        <style>
        @keyframes blink { 0% { opacity: 1; } 50% { opacity: 0.4; } 100% { opacity: 1; } }
        .main-title { font-size: 22px; font-weight: 900; color: #1F4E78; margin-bottom: 0px; }
        .sub-title { font-size: 13px; font-weight: 600; color: #495057; margin-bottom: 10px; }
        .card-cyan { background-color: #0dcaf0; color: #000; padding: 10px; border-radius: 6px; text-align: center; }
        .card-gray { background-color: #495057; color: white; padding: 10px; border-radius: 6px; text-align: center; }
        .card-orange { background-color: #fd7e14; color: white; padding: 10px; border-radius: 6px; text-align: center; animation: blink 1.2s infinite; }
        .card-green { background-color: #198754; color: white; padding: 10px; border-radius: 6px; text-align: center; }
        .card-title { font-size: 11px; font-weight: 800; }
        .card-value { font-size: 16px; font-weight: 900; }
        .grid-table { width: 100%; border-collapse: separate; border-spacing: 5px; margin-bottom: 10px; table-layout: fixed; }
        </style>
    """, unsafe_allow_html=True)

    current_date = get_shift_date()
    df = load_data_from_gsheet(current_date)
    
    if not df.empty:
        df.columns = df.columns.astype(str).str.strip()
        
        col_veh = next((c for c in df.columns if 'vehicle' in c.lower()), None)
        col_plan = next((c for c in df.columns if 'plan' in c.lower()), None)
        col_status = next((c for c in df.columns if 'status' in c.lower()), None)
        col_actual = next((c for c in df.columns if 'actual' in c.lower()), None)
        col_time = next((c for c in df.columns if 'time' in c.lower()), None)
        col_trans = next((c for c in df.columns if 'transport' in c.lower()), None)
        col_dest = next((c for c in df.columns if 'destination' in c.lower() or 'destinations' in c.lower()), None)
        col_prog = next((c for c in df.columns if 'prog' in c.lower() or 'pro' in c.lower()), None)
        
        if col_veh: df = df.dropna(subset=[col_veh])
        total_vehicles = len(df)
        total_plan_mt = float(df[col_plan].sum()) if col_plan and not df.empty else 0.0
        
        ul_df = df[df[col_status].astype(str).str.lower().str.contains('under loading', na=False)].copy() if col_status else pd.DataFrame()
        final_df = df[df[col_status].astype(str).str.lower().str.contains('final', na=False)].copy() if col_status else pd.DataFrame()
        
        ul_plan_sum = float(ul_df[col_plan].sum()) if col_plan and not ul_df.empty else 0.0
        final_act_sum = float(final_df[col_actual].sum()) if final_df is not None and not final_df.empty else 0.0

        st.markdown(f'<div class="main-title">🚛 M/s Surya Roshni Limited - Hindupur</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="sub-title">Smart Dispatch Management System ({current_date})</div>', unsafe_allow_html=True)
        
        st.markdown(f"""
            <table class="grid-table">
                <tr>
                    <td style="width: 50%;">
                        <div class="card-cyan"><div class="card-title">TOTAL PLAN MT</div><div class="card-value">{total_plan_mt:.2f} MT</div></div>
                    </td>
                    <td style="width: 50%;">
                        <div class="card-gray"><div class="card-title">TOTAL VEHICLES</div><div class="card-value">{total_vehicles}</div></div>
                    </td>
                </tr>
                <tr>
                    <td style="width: 50%;">
                        <div class="card-orange"><div class="card-title">⏳ UNDER LOADING ({len(ul_df)} Veh)</div><div class="card-value">{ul_plan_sum:.2f} MT</div></div>
                    </td>
                    <td style="width: 50%;">
                        <div class="card-green"><div class="card-title">FINAL VEHICLES ({len(final_df)})</div><div class="card-value">{final_act_sum:.2f} T Actual</div></div>
                    </td>
                </tr>
            </table>
        """, unsafe_allow_html=True)
        
        notice_msg = get_notice_from_txt()
        st.markdown(f"""
            <marquee style="background-color: #FFF3CD; color: #856404; padding: 6px; font-weight: bold; font-size: 13px; border: 1px solid #FFEEBA; border-radius: 4px; margin-bottom: 10px;" behavior="scroll" direction="left">
                {notice_msg}
            </marquee>
        """, unsafe_allow_html=True)
        
        st.markdown("### 🚚 Under Loading Vehicles (Detailed View)")
        if not ul_df.empty:
            view_cols = [c for c in [col_time, col_prog, col_veh, col_trans, col_dest, col_plan] if c]
            df_view = ul_df[view_cols].copy()
            df_view.insert(0, 'Sr. No.', range(1, len(df_view) + 1))
            df_view = df_view.rename(columns={
                col_time: 'Time',
                col_prog: 'Program No.',
                col_veh: 'Vehicle No.',
                col_trans: 'Transport Name',
                col_dest: 'Destination',
                col_plan: 'Plan (MT)'
            })
            st.dataframe(df_view, use_container_width=True, hide_index=True)
        else:
            st.success("फिलहाल कोई अंडर लोडिंग गाड़ी नहीं है।")

        st.markdown("### 🏁 Final Vehicles (Completed Dispatch)")
        if not final_df.empty:
            final_view_cols = [c for c in [col_time, col_prog, col_veh, col_trans, col_dest, col_plan, col_actual] if c]
            df_final_view = final_df[final_view_cols].copy()
            df_final_view.insert(0, 'Sr. No.', range(1, len(df_final_view) + 1))
            df_final_view = df_final_view.rename(columns={
                col_time: 'Time',
                col_prog: 'Program No.',
                col_veh: 'Vehicle No.',
                col_trans: 'Transport Name',
                col_dest: 'Destination',
                col_plan: 'Plan (MT)',
                col_actual: 'Actual (Tons)'
            })
            st.dataframe(df_final_view, use_container_width=True, hide_index=True)
        else:
            st.info("आज अभी तक कोई गाड़ी Final नहीं हुई है।")
    else:
        st.warning("Google Sheet से डेटा लोड नहीं हो पाया या शीट खाली है।")

# ==================== ADMIN PANEL ====================
elif page == "🔐 Admin Panel":
    if not os.path.exists(PASSWORD_FILE):
        with open(PASSWORD_FILE, "w", encoding="utf-8") as f:
            f.write("admin|1234")

    def get_credentials():
        try:
            with open(PASSWORD_FILE, "r", encoding="utf-8") as f:
                content = f.read().strip()
                if "|" in content:
                    return content.split("|")[0], content.split("|")[1]
        except:
            pass
        return "admin", "1234"

    def save_credentials(user, pwd):
        try:
            with open(PASSWORD_FILE, "w", encoding="utf-8") as f:
                f.write(f"{user}|{pwd}")
            return True
        except:
            return False

    if "logged_in" not in st.session_state: st.session_state.logged_in = False
    if "show_forgot" not in st.session_state: st.session_state.show_forgot = False

    saved_user, saved_pwd = get_credentials()

    if not st.session_state.logged_in:
        st.subheader("🔐 Admin Login")
        
        if not st.session_state.show_forgot:
            input_user = st.text_input("Username")
            input_pwd = st.text_input("Password", type="password")
            
            col_l1, col_l2 = st.columns(2)
            with col_l1:
                if st.button("Login"):
                    if input_user == saved_user and input_pwd == saved_pwd:
                        st.session_state.logged_in = True
                        st.rerun()
                    else:
                        st.error("गलत Username या Password!")
            with col_l2:
                if st.button("🔑 Forgot Password?"):
                    st.session_state.show_forgot = True
                    st.rerun()
        else:
            st.info("पासवर्ड रीसेट करने के लिए नया यूजरनेम और पासवर्ड सेट करें:")
            new_u = st.text_input("New Username", value=saved_user)
            new_p = st.text_input("New Password", type="password")
            confirm_p = st.text_input("Confirm Password", type="password")
            
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                if st.button("Save & Update"):
                    if new_p and new_p == confirm_p:
                        if save_credentials(new_u, new_p):
                            st.success("पासवर्ड सफलतापूर्वक बदल गया! अब नए पासवर्ड से लॉगिन करें।")
                            st.session_state.show_forgot = False
                            st.rerun()
                        else:
                            st.error("सेव करने में त्रुटि!")
                    else:
                        st.error("पासवर्ड मेल नहीं खा रहे हैं या खाली हैं!")
            with col_f2:
                if st.button("Back to Login"):
                    st.session_state.show_forgot = False
                    st.rerun()
        st.stop()

    st.title("🛠️ Admin Control Panel & Reports")
    
    with st.expander("⚙️ Change Username / Password"):
        with st.form("change_pwd_form"):
            cu = st.text_input("New Username", value=saved_user)
            cp = st.text_input("New Password", type="password", value=saved_pwd)
            if st.form_submit_button("Update Credentials"):
                if save_credentials(cu, cp):
                    st.success("यूजरनेम और पासवर्ड अपडेट हो गए हैं!")
                    st.rerun()
                else:
                    st.error("अपडेट करने में समस्या आई।")

    tab1, tab2, tab3 = st.tabs(["📝 एंट्री / अपडेट फॉर्म", "📊 PDF रिपोर्ट्स और WhatsApp", "📢 हिंदी नोटिस बोर्ड (Notice Editor)"])
    
    with tab1:
        current_date = get_shift_date()
        sel_sheet = st.text_input("📅 तारीख की शीट का नाम (Tab Name):", value=current_date)
        
        df_adm = load_data_from_gsheet(sel_sheet)
        
        if not df_adm.empty:
            df_adm.columns = df_adm.columns.astype(str).str.strip()
            
            a_veh = next((c for c in df_adm.columns if 'vehicle' in c.lower()), None)
            a_prog = next((c for c in df_adm.columns if 'prog' in c.lower() or 'program' in c.lower()), None)
            a_trans = next((c for c in df_adm.columns if 'transport' in c.lower()), None)
            a_dest = next((c for c in df_adm.columns if 'destination' in c.lower()), None)
            a_plan = next((c for c in df_adm.columns if 'plan' in c.lower()), None)
            a_actual = next((c for c in df_adm.columns if 'actual' in c.lower()), None)
            a_status = next((c for c in df_adm.columns if 'status' in c.lower()), None)
            a_remarks = next((c for c in df_adm.columns if 'remark' in c.lower()), None)
            
            mode = st.radio("ऑप्शन चुनें:", ["नई एंट्री करें", "मौजूदा गाड़ी अपडेट करें"], horizontal=True)
            
            if mode == "मौजूदा गाड़ी अपडेट करें" and a_veh and not df_adm.empty:
                vehicle_display_list = []
                veh_mapping = {}
                
                for _, row in df_adm.iterrows():
                    v_no = str(row[a_veh]).strip() if pd.notna(row[a_veh]) else ""
                    if v_no and v_no != "nan":
                        t_name = str(row[a_trans]) if a_trans and pd.notna(row[a_trans]) else "N/A"
                        d_name = str(row[a_dest]) if a_dest and pd.notna(row[a_dest]) else "N/A"
                        s_name = str(row[a_status]) if a_status and pd.notna(row[a_status]) else "Under Loading"
                        
                        display_text = f"{v_no} | {t_name} | {d_name} ({s_name})"
                        vehicle_display_list.append(display_text)
                        veh_mapping[display_text] = v_no
                
                selected_display = st.selectbox("गाड़ी चुनें (Vehicle | Transport | Destination):", vehicle_display_list) if vehicle_display_list else ""
                sel_v = veh_mapping.get(selected_display, "") if selected_display else ""
            else:
                sel_v = ""

            def_prog, def_trans, def_dest, def_plan, def_actual, def_status, def_remarks = "", "", "", 0.0, 0.0, "Under Loading", ""
            if mode == "मौजूदा गाड़ी अपडेट करें" and sel_v and a_veh:
                matched_row = df_adm[df_adm[a_veh].astype(str).str.strip() == sel_v]
                if not matched_row.empty:
                    r = matched_row.iloc[0]
                    def_prog = str(r[a_prog]) if a_prog and pd.notna(r[a_prog]) else ""
                    def_trans = str(r[a_trans]) if a_trans and pd.notna(r[a_trans]) else ""
                    def_dest = str(r[a_dest]) if a_dest and pd.notna(r[a_dest]) else ""
                    def_plan = float(r[a_plan]) if a_plan and pd.notna(r[a_plan]) else 0.0
                    def_actual = float(r[a_actual]) if a_actual and pd.notna(r[a_actual]) else 0.0
                    def_status = str(r[a_status]) if a_status and pd.notna(r[a_status]) else "Under Loading"
                    def_remarks = str(r[a_remarks]) if a_remarks and pd.notna(r[a_remarks]) else ""

            with st.form("entry_form"):
                col1, col2 = st.columns(2)
                with col1:
                    if mode == "नई एंट्री करें":
                        veh = st.text_input("Vehicle No.").upper()
                    else:
                        veh = st.text_input("Vehicle No.", value=sel_v, disabled=True)
                    
                    prog = st.text_input("Program No.", value=def_prog)
                    trans = st.text_input("Transport Name", value=def_trans)
                    dest = st.text_input("Destination", value=def_dest)
                with col2:
                    plan = st.number_input("Loading Plan (MT)", value=def_plan)
                    actual = st.number_input("Actual Loading (Tons)", value=def_actual)
                    
                    status_options = ["Under Loading", "Final"]
                    s_idx = status_options.index(def_status) if def_status in status_options else 0
                    status = st.selectbox("Status", status_options, index=s_idx)
                    
                    remarks = st.text_input("Remarks", value=def_remarks)
                    
                if st.form_submit_button("💾 डेटा सेव करें"):
                    if mode == "नई एंट्री करें" and not veh: 
                        st.error("गाड़ी नंबर अनिवार्य है!")
                    else:
                        t_row = len(df_adm) + 2
                        row_data = [
                            len(df_adm) + 1,
                            sel_sheet,
                            datetime.now().strftime("%H:%M"),
                            prog,
                            veh if mode == "नई एंट्री करें" else sel_v,
                            trans,
                            dest,
                            plan,
                            "",
                            actual if actual > 0 else "",
                            status,
                            remarks
                        ]
                        
                        sync_row_to_google_sheet(sel_sheet, row_data)
                        st.success("डेटा सफलतापूर्वक Google Sheet में सेव/अपडेट कर दिया गया!")
                        st.rer()

    with tab2:
        st.subheader("📊 रिपोर्ट्स और PDF जनरेटर")
        rep_date = st.text_input("Report Date (DD.MM.YYYY)", get_shift_date())
        
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            if st.button("📄 Loader List PDF"):
                generate_loader_list_pdf(rep_date)
            if st.button("⏳ Under Loading PDF"):
                generate_status_pdf("Under Loading", rep_date)
            if st.button("🏁 Final Report PDF"):
                generate_status_pdf("Final", rep_date)
        with col_b2:
            if st.button("💬 Shift A WhatsApp"):
                webbrowser.open("https://web.whatsapp.com/")
                st.success("WhatsApp ओपन हो गया है!")
            if st.button("💬 Shift B WhatsApp"):
                webbrowser.open("https://web.whatsapp.com/")
                st.success("WhatsApp ओपन हो गया है!")

    with tab3:
        st.subheader("📢 डैशबोर्ड का हिंदी/इंग्लिश नोटिस बदलें")
        current_notice = get_notice_from_txt()
        new_notice_input = st.text_area("यहाँ नया नोटिस (हिंदी या अंग्रेजी में) लिखें:", value=current_notice, height=100)
        
        if st.button("🚀 नोटिस अपडेट करें (Save Notice)"):
            if save_notice_to_txt(new_notice_input):
                st.success("✅ नोटिस सफलतापूर्वक अपडेट हो गया है! अब 'Live Dashboard' पर जाएँ, यह तुरंत बदल गया होगा।")
            else:
                st.error("❌ नोटिस सेव करने में समस्या आई।")

if st.button("🔄 Refresh Data"):
    st.rerun()
