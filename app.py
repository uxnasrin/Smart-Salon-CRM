import streamlit as st
import pandas as pd
from datetime import timedelta
from streamlit_option_menu import option_menu
import io
import re

# --- 1. PAGE SETUP & GLOBAL CSS ---
st.set_page_config(page_title="Smart Salon CRM", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
    .stApp { background-color: #0F172A; color: #F8FAFC; }
    #MainMenu {visibility: hidden;} footer {visibility: hidden;} 
    .top-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }
    .date-text { color: #94A3B8; font-size: 14px; margin-right: 15px; }
    div.stButton > button[kind="primary"] { background-color: #4F46E5 !important; color: white !important; border: none !important; border-radius: 8px !important; padding: 5px 20px !important; font-weight: 600 !important; }
    div.stButton > button[kind="secondary"] { background-color: #334155 !important; color: white !important; border: none !important; border-radius: 5px !important; font-weight: bold !important; }
    
    .metric-card { background-color: #1E293B; border: 1px solid #334155; border-radius: 12px; padding: 20px; text-align: center; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); }
    .metric-value { font-size: 28px; font-weight: bold; color: #F8FAFC; margin: 10px 0 5px 0; }
    .metric-label { font-size: 14px; color: #94A3B8; }
    
    .custom-table { width: 100%; border-collapse: collapse; background-color: #1E293B; border-radius: 12px; overflow: hidden; margin-bottom: 20px;}
    .custom-table th { text-align: left; padding: 12px; color: #94A3B8; font-size: 13px; font-weight: 600; border-bottom: 1px solid #334155; white-space: nowrap;}
    .custom-table td { padding: 12px; color: #F8FAFC; font-size: 14px; border-bottom: 1px solid #334155; }
    
    .nowrap { white-space: nowrap; }
    .wrap-text { white-space: normal; word-wrap: break-word; max-width: 250px; }
    
    .service-text { color: #60A5FA; font-weight: 600; font-size: 13px; }
    .pending-text { color: #F87171; font-weight: 600; font-size: 13px; }
    .recent-visit-row td { padding-top: 5px !important; padding-bottom: 12px !important; color: #94A3B8 !important; font-size: 12px !important; border-bottom: 1px solid #334155 !important; font-style: italic; }
    .red-dot { color: #EF4444; font-size: 16px; line-height: 0; }
    
    /* New Client Badge Style */
    .new-client-badge { color: #10B981; font-weight: 600; font-size: 13px; font-style: normal; }
</style>
""", unsafe_allow_html=True)

# --- PAGINATION HELPER FUNCTIONS ---
def paginate_dataframe(dataframe, page_size, menu_name):
    total_pages = max(1, (len(dataframe) - 1) // page_size + 1)
    
    if f'page_{menu_name}' not in st.session_state:
        st.session_state[f'page_{menu_name}'] = 1
        
    if st.session_state[f'page_{menu_name}'] > total_pages:
        st.session_state[f'page_{menu_name}'] = total_pages
        
    current_page = st.session_state[f'page_{menu_name}']
    start_idx = (current_page - 1) * page_size
    end_idx = start_idx + page_size
    return dataframe.iloc[start_idx:end_idx], current_page, total_pages

def render_pagination_controls(current_page, total_pages, total_items, menu_name):
    if total_items == 0: return
    
    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2, col3, col4, col5 = st.columns([0.4, 0.15, 0.15, 0.15, 0.15])
    
    start_item = (current_page - 1) * 10 + 1
    end_item = min(current_page * 10, total_items)
    
    with col1:
        st.markdown(f"<div style='color:#94A3B8; font-size:14px; padding-top:10px;'>Showing {start_item} to {end_item} of {total_items} customers</div>", unsafe_allow_html=True)
    
    with col3:
        if st.button("◀ Prev", key=f"prev_{menu_name}", use_container_width=True, disabled=(current_page == 1)):
            st.session_state[f'page_{menu_name}'] -= 1
            st.rerun()
    with col4:
        st.markdown(f"<div style='text-align:center; padding-top:5px; font-weight:bold; color:white; background-color:#4F46E5; border-radius:5px; padding-bottom:5px;'>{current_page} / {total_pages}</div>", unsafe_allow_html=True)
    with col5:
        if st.button("Next ▶", key=f"next_{menu_name}", use_container_width=True, disabled=(current_page == total_pages)):
            st.session_state[f'page_{menu_name}'] += 1
            st.rerun()

# --- ADD CUSTOMER LOGIC ---
def save_new_customer(name, phone, service, stylist, date):
    try:
        all_sheets = pd.read_excel("Dummy_Salon_Data.xlsx", sheet_name=None)
        sheet_name = list(all_sheets.keys())[0]
        new_row = {"Customer_Name": name, "Phone_Number": phone, "Service": service, "Stylist": stylist, "Date": date.strftime("%d-%m-%Y")}
        all_sheets[sheet_name] = pd.concat([all_sheets[sheet_name], pd.DataFrame([new_row])], ignore_index=True)
        
        with pd.ExcelWriter("Dummy_Salon_Data.xlsx") as writer:
            for s_name, s_df in all_sheets.items():
                s_df.to_excel(writer, sheet_name=s_name, index=False)
        return True
    except Exception as e:
        return False

# --- 2. ADVANCED DATA PROCESSING ---
@st.cache_data
def load_data():
    try:
        raw_df = pd.read_excel("Dummy_Salon_Data.xlsx", sheet_name=0)
        master_df = pd.read_excel("Dummy_Salon_Data.xlsx", sheet_name="Service_Master")
    except Exception as e:
        return pd.DataFrame(), pd.DataFrame(), 0, 0, True
    
    total_visits = len(raw_df)
    unique_clients = len(raw_df['Phone_Number'].dropna().unique())
    
    master_df['Clean'] = master_df.iloc[:, 0].astype(str).str.strip().str.upper().str.replace(" ", "")
    raw_df['Parsed_Date'] = pd.to_datetime(raw_df['Date'], dayfirst=True, errors='coerce')
    
    records = []
    for _, row in raw_df.iterrows():
        raw_service = str(row.get('Service', '')).upper()
        customer_name = str(row.get('Customer_Name', '')).strip().upper()
        
        female_keywords = ['A', 'I', 'E', 'EE', 'YA', 'INI', 'IKA', 'ITA', 'MI', 'VI', 'BEGUM', 'BANU', 'FATIMA', 'NASREEN', 'SHREE', 'DEVI']
        is_female = any(customer_name.endswith(k) for k in female_keywords) or any(k in customer_name for k in ['SHREE', 'DEVI'])
        
        raw_service = raw_service.replace('H/C', 'HAIRCUT').replace('B/T', 'BEARD TRIM')
        services = [s.strip() for s in re.split(r'[/,\+]', raw_service) if s.strip()]
        
        for s in services:
            s_key = s.upper().replace(" ", "")
            match = master_df[master_df['Clean'] == s_key]
            
            if not match.empty:
                if is_female and 'Ladies_Days' in match.columns:
                    duration = int(match['Ladies_Days'].values[0])
                elif not is_female and 'Gents_Days' in match.columns:
                    duration = int(match['Gents_Days'].values[0])
                else:
                    try: duration = int(match.iloc[:, 1].values[0])
                    except: duration = 30
            else:
                duration = 60 if is_female else 25 
            
            parsed_date = row['Parsed_Date']
            due_date = parsed_date + timedelta(days=duration) if pd.notnull(parsed_date) else pd.NaT
            
            new_row = row.to_dict()
            new_row['Pending_Service'] = s
            new_row['Next_Due_Date'] = due_date
            records.append(new_row)
            
    exploded_df = pd.DataFrame(records)
    valid_dates_df = exploded_df.dropna(subset=['Parsed_Date'])
    
    latest_df = valid_dates_df.sort_values('Parsed_Date').drop_duplicates(subset=['Phone_Number', 'Pending_Service'], keep='last')
    
    grouped_df = latest_df.groupby(
        ['Customer_Name', 'Phone_Number', 'Parsed_Date', 'Stylist', 'Service', 'Next_Due_Date'],
        dropna=False
    ).agg({'Pending_Service': lambda x: ', '.join(pd.Series(x).unique())}).reset_index()
    
    return raw_df, grouped_df, unique_clients, total_visits, False

raw_df, grouped_df, unique_clients, total_visits, is_error = load_data()

# --- 3. EXCEL GENERATOR (NEW!) ---
def create_excel(report_df):
    output = io.BytesIO()
    # எக்செல் ஷீட்டிற்கு தேவையான காலம்களை மட்டும் அழகாக பிரித்தெடுக்கிறோம்
    export_df = report_df[['Customer_Name', 'Phone_Number', 'Parsed_Date', 'Stylist', 'Service', 'Pending_Service', 'Next_Due_Date']].copy()
    export_df.columns = ['Customer Name', 'Mobile Number', 'Last Visit', 'Stylist', 'Service Done', 'Pending Service', 'Due Date']
    
    # தேதியை எக்செல்-க்கு ஏற்றவாறு மாற்றுதல்
    export_df['Last Visit'] = export_df['Last Visit'].dt.strftime('%d %b %Y')
    export_df['Due Date'] = export_df['Due Date'].dt.strftime('%d %b %Y')
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        export_df.to_excel(writer, index=False, sheet_name='Report')
    
    return output.getvalue()

# --- 4. SIDEBAR NAVIGATION ---
with st.sidebar:
    st.markdown("### ✂️ SMART SALON\n<span style='color:#94A3B8; font-size: 12px;'>CRM (Local Secure 🔒)</span>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    
    selected_menu = option_menu(
        menu_title=None,
        options=["Dashboard", "Due Today", "Overdue", "Upcoming", "Reports"],
        icons=["house", "calendar-day", "clock-history", "calendar-week", "file-earmark-bar-graph"],
        menu_icon="cast", default_index=0,
        styles={"container": {"padding": "0!important", "background-color": "#0F172A"}, "icon": {"color": "#94A3B8", "font-size": "18px"}, "nav-link": {"font-size": "15px", "text-align": "left", "margin":"5px", "color": "#F8FAFC"}, "nav-link-selected": {"background-color": "#4338CA"}}
    )

# --- 5. HTML TABLE GENERATOR ---
def render_custom_table(dataframe, show_recent_visit=False):
    html = "<table class='custom-table'><tr><th>Customer Name</th><th>Mobile Number</th><th>Last Visit</th><th>Stylist</th><th>Service Done</th><th>Pending Service</th><th>Due Date</th></tr>"
    for _, row in dataframe.iterrows():
        cust = row.get('Customer_Name', '-')
        phone = row.get('Phone_Number', '-')
        curr_date = row['Parsed_Date']
        last_visit = curr_date.strftime('%d %b %Y') if pd.notnull(curr_date) else '-'
        stylist = row.get('Stylist', '-')
        
        done_service = row.get('Service', '-') 
        pending = row.get('Pending_Service', '-') 
        due_date = f"{row['Next_Due_Date'].strftime('%d %b %Y')} <span class='red-dot'>•</span>" if pd.notnull(row['Next_Due_Date']) else '-'
            
        html += f"<tr><td class='nowrap'>{cust}</td><td class='nowrap' style='font-weight: bold;'>{phone}</td><td class='nowrap'>{last_visit}</td><td class='nowrap'>{stylist}</td><td class='wrap-text service-text'>{done_service}</td><td class='wrap-text pending-text'>{pending}</td><td class='nowrap'>{due_date}</td></tr>"
        
        if show_recent_visit:
            past = raw_df[(raw_df['Phone_Number'] == phone) & (raw_df['Parsed_Date'] < curr_date)].sort_values('Parsed_Date')
            if not past.empty:
                last_past = past.iloc[-1]
                p_date = last_past['Parsed_Date'].strftime('%d %b %Y')
                p_srv = last_past.get('Service', '-')
                p_sty = last_past.get('Stylist', '-')
                html += f"<tr class='recent-visit-row'><td colspan='7'>↳ Previous Visit: {p_date} - {p_srv} <b>(Stylist: {p_sty})</b></td></tr>"
            else:
                html += f"<tr class='recent-visit-row'><td colspan='7'>↳ <span class='new-client-badge'>✅ New Client</span></td></tr>"
                
    html += "</table>"
    st.markdown(html, unsafe_allow_html=True)

# --- 6. MAIN APP LOGIC ---
if not is_error and not grouped_df.empty:
    today = pd.Timestamp.now().normalize()
    today_str = today.strftime('%d %b %Y')

    if 'show_success' in st.session_state and st.session_state['show_success']:
        st.toast("✅ New Customer Added Successfully!", icon="🎉")
        st.session_state['show_success'] = False

    # --- NEW SEARCH BAR ---
    search_query = st.text_input("Search", placeholder="🔍 Search Customer Name or Mobile Number...", label_visibility="collapsed")
    
    if search_query:
        grouped_df = grouped_df[grouped_df['Customer_Name'].str.contains(search_query, case=False, na=False) | grouped_df['Phone_Number'].astype(str).str.contains(search_query, na=False)]
    # ----------------------

    due_today_df = grouped_df[grouped_df['Next_Due_Date'] == today].copy()
    
    overdue_df = grouped_df[grouped_df['Next_Due_Date'] < today].copy()
    overdue_df = overdue_df.sort_values('Parsed_Date').drop_duplicates(subset=['Phone_Number'], keep='last')
    
    upcoming_df = grouped_df[(grouped_df['Next_Due_Date'] > today) & (grouped_df['Next_Due_Date'] <= today + timedelta(days=7))].copy()
    
    col1, col2 = st.columns([0.8, 0.2])
    with col2:
        st.markdown(f"<div style='text-align: right;'><span class='date-text'>📅 {today_str}</span></div>", unsafe_allow_html=True)

    with st.expander("➕ Add New Customer", expanded=False):
        with st.form("add_customer_form", clear_on_submit=True):
            st.write("Enter Client Details:")
            c1, c2 = st.columns(2)
            new_name = c1.text_input("Customer Name")
            new_phone = c2.text_input("Phone Number")
            new_service = c1.text_input("Service Done (e.g., HAIRCUT / FACIAL)")
            new_stylist = c2.text_input("Stylist Name")
            new_date = st.date_input("Visit Date")
            
            if st.form_submit_button("Save Customer"):
                if new_name and new_phone and new_service:
                    success = save_new_customer(new_name, new_phone, new_service, new_stylist, new_date)
                    if success:
                        st.session_state['show_success'] = True
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.toast("❌ Error: ஆட் ஆகவில்லை! சர்வர் இஸ்யூ அல்லது எக்செல் ஓபனில் உள்ளது.", icon="⚠️")
                else:
                    st.warning("Please fill Name, Phone, and Service.")

    if selected_menu == "Dashboard":
        st.markdown("<h2 style='margin-top: -10px;'>Hello Admin 👋</h2>", unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns(4)
        with col1: st.markdown(f"<div class='metric-card'><div class='metric-label'>Unique Clients</div><div class='metric-value'>👥 {unique_clients:,}</div></div>", unsafe_allow_html=True)
        with col2: st.markdown(f"<div class='metric-card'><div class='metric-label'>Total Visits</div><div class='metric-value'>🔄 {total_visits:,}</div></div>", unsafe_allow_html=True)
        with col3: st.markdown(f"<div class='metric-card'><div class='metric-label'>Due Today</div><div class='metric-value'>📅 {len(due_today_df)}</div></div>", unsafe_allow_html=True)
        with col4: st.markdown(f"<div class='metric-card'><div class='metric-label'>Overdue</div><div class='metric-value'>⚠️ {len(overdue_df)}</div></div>", unsafe_allow_html=True)
        
        st.markdown("<br><h3>Recent Clients</h3>", unsafe_allow_html=True)
        recent_clients_df = grouped_df.sort_values(by='Parsed_Date', ascending=False).head(10)
        render_custom_table(recent_clients_df, show_recent_visit=True)

    elif selected_menu == "Due Today":
        st.markdown("<h2 style='margin-top: -10px;'>Due Today</h2>", unsafe_allow_html=True)
        if not due_today_df.empty: 
            paginated_df, curr_page, tot_pages = paginate_dataframe(due_today_df, 10, "due_today")
            render_custom_table(paginated_df, show_recent_visit=True)
            render_pagination_controls(curr_page, tot_pages, len(due_today_df), "due_today")
        else: st.success("No customers are due today!")

    elif selected_menu == "Overdue":
        st.markdown("<h2 style='margin-top: -10px;'>Overdue</h2>", unsafe_allow_html=True)
        if not overdue_df.empty: 
            paginated_df, curr_page, tot_pages = paginate_dataframe(overdue_df, 10, "overdue")
            render_custom_table(paginated_df, show_recent_visit=True)
            render_pagination_controls(curr_page, tot_pages, len(overdue_df), "overdue")
        else: st.success("No overdue customers!")
            
    elif selected_menu == "Upcoming":
        st.markdown("<h2 style='margin-top: -10px;'>Upcoming (Next 7 Days)</h2>", unsafe_allow_html=True)
        if not upcoming_df.empty: 
            paginated_df, curr_page, tot_pages = paginate_dataframe(upcoming_df, 10, "upcoming")
            render_custom_table(paginated_df, show_recent_visit=True)
            render_pagination_controls(curr_page, tot_pages, len(upcoming_df), "upcoming")
        else: st.info("No upcoming customers in the next 7 days.")
            
    elif selected_menu == "Reports":
        st.markdown("<h2 style='margin-top: -10px;'>Reports</h2>", unsafe_allow_html=True)
        
        col_f1, col_f2 = st.columns([0.4, 0.6])
        with col_f1:
            report_type = st.selectbox("Select Report Filter:", ["Due Today", "Overdue", "Upcoming"])
            
        if report_type == "Due Today": report_data = due_today_df
        elif report_type == "Overdue": report_data = overdue_df
        else: report_data = upcoming_df
        
        if not report_data.empty:
            # EXCEL DOWNLOAD BUTTON LOGIC
            excel_data = create_excel(report_data)
            st.download_button(
                label=f"📥 Download {report_type} Report (Excel)",
                data=excel_data,
                file_name=f"{report_type}_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.markdown("<br>", unsafe_allow_html=True)
            
            paginated_df, curr_page, tot_pages = paginate_dataframe(report_data, 10, f"reports_{report_type}")
            show_recent = False if report_type == "Overdue" else True
            render_custom_table(paginated_df, show_recent_visit=show_recent)
            render_pagination_controls(curr_page, tot_pages, len(report_data), f"reports_{report_type}")
        else:
            st.info(f"No data available for {report_type}.")
