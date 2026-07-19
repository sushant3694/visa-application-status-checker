import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
import os
import urllib3
import json
import re
import base64
import time as ptime
from datetime import datetime, time

# Suppress insecure request warnings from using verify=False
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

TARGET_URL = "https://www.ireland.ie/en/india/newdelhi/services/visas/processing-times-and-decisions/#visa-decisions"
LOCAL_FILE = "visa_decisions_latest.ods"
JSON_FILE = "visa_decisions.json"
MASTER_FILE = "Visa_Decision_Comparison_Report.xlsx"
IRELAND_FLAG_URL = "https://flagcdn.com/w80/ie.png"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
    "Connection": "keep-alive"
}

st.set_page_config(
    page_title="Ireland Visa Status Tracking Portal",
    page_icon=IRELAND_FLAG_URL,
    layout="centered"
)

# Initialize session tracking states
if "last_sync_date" not in st.session_state:
    st.session_state.last_sync_date = None
if "app_loaded" not in st.session_state:
    st.session_state.app_loaded = False

def fetch_latest_ods_url():
    try:
        response = requests.get(TARGET_URL, headers=HEADERS, timeout=15, verify=False)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        for link in soup.find_all('a', href=True):
            if link['href'].endswith('.ods'):
                url = link['href']
                return "https://www.ireland.ie" + url if url.startswith('/') else url
    except Exception:
        pass
    return None

def update_master_report(current_records, report_date, filename):
    if os.path.exists(MASTER_FILE):
        all_apps = pd.read_excel(MASTER_FILE, sheet_name="All Applications").astype(str)
        records_by_date = pd.read_excel(MASTER_FILE, sheet_name="New Records By Date").astype(str)
        daily_summary = pd.read_excel(MASTER_FILE, sheet_name="Daily Summary").astype(str)
        status_changes = pd.read_excel(MASTER_FILE, sheet_name="Status Changes").astype(str)
    else:
        all_apps = pd.DataFrame(columns=["Application Number", "Current Status", "Date First Added", "Initial Status", "Last Seen Date", "First Found File", "Last Seen File"])
        records_by_date = pd.DataFrame(columns=["Application Number", "Status", "Date Added", "First Found File"])
        daily_summary = pd.DataFrame(columns=["Report Date", "File Name", "Total Records In Report", "New Records Added", "New Approved", "New Refused", "New Other Status", "Status Changes", "Previous Report Date", "Previous Report File"])
        status_changes = pd.DataFrame(columns=["Application Number", "Previous Status", "New Status", "Status Changed Date", "Report File"])

    if report_date in daily_summary["Report Date"].values:
        return

    prev_known = dict(zip(all_apps["Application Number"], all_apps["Current Status"]))
    first_seen = dict(zip(all_apps["Application Number"], all_apps.to_dict(orient="records")))

    new_approved, new_refused, new_other, status_change_count = 0, 0, 0, 0
    latest_new_records = []

    for app_num, current_status in current_records.items():
        if app_num not in first_seen:
            first_seen[app_num] = {
                "Application Number": app_num, "Current Status": current_status, "Date First Added": report_date,
                "Initial Status": current_status, "Last Seen Date": report_date, "First Found File": filename, "Last Seen File": filename
            }
            latest_new_records.append({"Application Number": app_num, "Status": current_status, "Date Added": report_date, "First Found File": filename})
            
            if current_status == "Approved": new_approved += 1
            elif current_status == "Refused": new_refused += 1
            else: new_other += 1
        else:
            first_seen[app_num]["Current Status"] = current_status
            first_seen[app_num]["Last Seen Date"] = report_date
            first_seen[app_num]["Last Seen File"] = filename

        old_status = prev_known.get(app_num)
        if old_status and old_status != current_status:
            new_change = {"Application Number": app_num, "Previous Status": old_status, "New Status": current_status, "Status Changed Date": report_date, "Report File": filename}
            status_changes = pd.concat([status_changes, pd.DataFrame([new_change])], ignore_index=True)
            status_change_count += 1

    new_records_count = len(latest_new_records)
    if latest_new_records:
        records_by_date = pd.concat([records_by_date, pd.DataFrame(latest_new_records)], ignore_index=True)

    prev_report_date = daily_summary.iloc[-1]["Report Date"] if not daily_summary.empty else ""
    prev_report_file = daily_summary.iloc[-1]["File Name"] if not daily_summary.empty else ""
    
    new_summary = {
        "Report Date": report_date, "File Name": filename, "Total Records In Report": len(current_records),
        "New Records Added": new_records_count, "New Approved": new_approved, "New Refused": new_refused,
        "New Other Status": new_other, "Status Changes": status_change_count, "Previous Report Date": prev_report_date, "Previous Report File": prev_report_file
    }
    daily_summary = pd.concat([daily_summary, pd.DataFrame([new_summary])], ignore_index=True)
    all_apps = pd.DataFrame(list(first_seen.values()))

    with pd.ExcelWriter(MASTER_FILE, engine="openpyxl") as writer:
        pd.DataFrame(latest_new_records).to_excel(writer, sheet_name="Latest New Records", index=False)
        all_apps.to_excel(writer, sheet_name="All Applications", index=False)
        records_by_date.to_excel(writer, sheet_name="New Records By Date", index=False)
        daily_summary.to_excel(writer, sheet_name="Daily Summary", index=False)
        status_changes.to_excel(writer, sheet_name="Status Changes", index=False)

def download_and_convert_production(today_str):
    download_url = fetch_latest_ods_url()
    if download_url:
        try:
            file_response = requests.get(download_url, headers=HEADERS, timeout=30, verify=False)
            file_response.raise_for_status()
            with open(LOCAL_FILE, 'wb') as f:
                f.write(file_response.content)
            
            df = pd.read_excel(LOCAL_FILE, engine='odf', header=None).astype(str).replace(['nan', 'NaN', 'None'], '')
            cleaned_map = {}
            for row in df.values.tolist():
                cells = [str(cell).strip() for cell in row if str(cell).strip()]
                app_num = next((c for c in cells if re.match(r'^\d{8}$', c)), None)
                if app_num:
                    other_cells = [c for c in cells if c != app_num]
                    decision = "Refused" if any(x in "".join(other_cells).lower() for x in ["refus", "reject", "deni"]) else "Approved"
                    cleaned_map[app_num] = decision
            
            records = [{"application_number": k, "decision_status": v} for k, v in cleaned_map.items()]
            with open(JSON_FILE, 'w', encoding='utf-8') as f:
                json.dump(records, f, ensure_ascii=False, indent=4)

            filename_stamp = f"{today_str.replace('-', '')}_NDVO_Visa_Decisions.ods"
            update_master_report(cleaned_map, today_str, filename_stamp)
            
            load_cached_json.clear()
            return True
        except Exception:
            pass
    return False

@st.cache_data(show_spinner=False)
def load_cached_json():
    if os.path.exists(JSON_FILE):
        with open(JSON_FILE, 'r', encoding='utf-8') as f:
            return pd.DataFrame(json.load(f))
    return None

# --- FIX: FORCED RENDER LOADER PLATFORM ROUTE ---
now = datetime.now()
today_str = now.strftime("%Y-%m-%d")

if not st.session_state.app_loaded:
    # Use empty space placeholder container to block DOM caching
    placeholder = st.empty()
    with placeholder.container():
        with st.spinner("✨ Synchronizing processing ledgers and master data calculations..."):
            start_time = ptime.time()
            
            sync_needed = not os.path.exists(JSON_FILE)
            if not sync_needed and now.time() >= time(11, 5):
                if st.session_state.last_sync_date != today_str:
                    file_date = datetime.fromtimestamp(os.path.getmtime(JSON_FILE)).strftime("%Y-%m-%d")
                    if file_date != today_str:
                        sync_needed = True

            if sync_needed:
                success = download_and_convert_production(today_str)
                if success:
                    st.session_state.last_sync_date = today_str

            elapsed = ptime.time() - start_time
            remaining = 5.0 - elapsed
            if remaining > 0:
                ptime.sleep(remaining)
                
    placeholder.empty() # Explicitly wipes out the loader container block
    st.session_state.app_loaded = True
    st.rerun()

df = load_cached_json()

def get_base64_image(image_path):
    if os.path.exists(image_path):
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    return ""

img_base64 = get_base64_image("background.avif")

def get_base64_file(file_path):
    if os.path.exists(file_path):
        with open(file_path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return ""

master_base64 = get_base64_file(MASTER_FILE)

# Advanced UI style injections
st.markdown(f"""
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:ital,wght@0,300;0,400;0,600;0,700;0,800;1,400&display=swap" rel="stylesheet">

    <style>
    *, html, body, .stApp, div, span, p, h1, h2, h3, label, input, button, textarea, th, td {{
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif !important;
    }}
    
    @keyframes fadeInPage {{
        from {{ opacity: 0; transform: translateY(15px); }}
        to {{ opacity: 1; transform: translateY(0); }}
    }}
    
    .stApp {{
        background: linear-gradient(rgba(23, 23, 23, 0.92), rgba(0, 0, 0, 0.92)), 
                    url("data:image/avif;base64,{img_base64}");
        background-size: cover;
        background-position: center;
        background-attachment: fixed;
        animation: fadeInPage 0.8s ease-out forwards;
    }}
    
    .block-container {{
        padding-left: 2rem !important;
        padding-right: 2rem !important;
        position: relative;
    }}
    
    h1 {{
        color: #007A4E !important;
        font-weight: 800;
        text-align: center;
        letter-spacing: -0.5px;
        font-size: 2.5rem;
        margin-bottom: 5px !important;
    }}

    h1 span:nth-of-type(2) {{
        display: none !important;
        visibility: hidden !important;
    }}

    header, #MainMenu, .stAppHeader {{
        visibility: hidden !important;
        display: none !important;
    }}

    a[href^="https://streamlit.io/cloud"], a[href*="streamlit.io"], .stAppDeployDropdown {{
        display: none !important;
        visibility: hidden !important;
    }}

    iframe {{
        position: absolute !important;
        z-index:9999999999999999 !important;
    }}
    
    .branding-subheading {{
        text-align: center;
        margin-top: 0px;
        margin-bottom: 20px;
        font-size: 1rem;
        color: #fff;
        font-weight: 500;
    }}
    
    .branding-subheading a, .floating-download a {{
        color: #007A4E !important;
        text-decoration: none;
        font-weight: 700;
        transition: color 0.3s ease;
    }}
    
    .branding-subheading a:hover, .floating-download a:hover {{
        color: #FF883E !important;
        text-decoration: underline;
    }}
    
    .floating-download {{
        text-align:right;
        font-size: 0.95rem;
        font-weight: 700;
        margin-bottom: 10px;
    }}
    
    .quote-box {{
        text-align: center;
        font-style: italic;
        color: #fff;
        margin-bottom: 25px;
        font-size: 1.1rem;
        line-height: 1.6;
    }}
    
    div.stButton > button:first-child {{
        background: linear-gradient(135deg, #007A4E 0%, #00b371 100%) !important;
        color: white !important;
        border-radius: 6px;
        width: 100%;
        font-weight: 700;
        font-size: 1rem;
        letter-spacing: 0.5px;
        border: none;
        padding: 12px;
        box-shadow: 0 4px 6px rgba(0, 122, 78, 0.2);
        transition: all 0.4s cubic-bezier(0.165, 0.84, 0.44, 1);
        cursor: pointer;
    }}
    
    div.stButton > button:first-child:hover {{
        transform: translateY(-2px);
        background: linear-gradient(135deg, #FF883E 0%, #ffaa6b 100%) !important;
        box-shadow: 0 8px 20px rgba(255, 136, 62, 0.4);
    }}

    @media (max-width: 768px) {{
        .block-container {{
            padding-left: 1.5rem !important;
            padding-right: 1.5rem !important;
            padding-top: 4.5rem !important;
        }}
        .floating-download {{ top: 15px; right: 1.5rem; width: 100%; text-align: right; }}
        h1 {{ font-size: 1.65rem !important; line-height: 1.3 !important; }}
        .branding-subheading {{ font-size: 0.85rem !important; }}
        .quote-box {{ font-size: 0.95rem !important; }}
    }}
    </style>
""", unsafe_allow_html=True)

# --- RENDERING UI VIEWS ---
st.markdown("<h1 style='margin-bottom: 0px;'>Ireland Visa Application Tracking</h1>", unsafe_allow_html=True)
st.markdown("<div class='branding-subheading'>Built with precision by <a href='https://www.sushantthorat.com/' target='_blank'> Sushant Thorat</a></div>", unsafe_allow_html=True)
st.markdown("<div class='quote-box'>\"Céad Míle Fáilte\" - A hundred thousand welcomes. Charting your pathway to the Emerald Isle.</div>", unsafe_allow_html=True)

if os.path.exists(MASTER_FILE) and master_base64:
    st.markdown(f"""
        <div class="floating-download">
            <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{master_base64}" download="Visa_Decision_Comparison_Report.xlsx">
                Download Master Report
            </a>
        </div>
    """, unsafe_allow_html=True)

if df is not None and not df.empty:
    search_query = st.text_input("Enter Your 8-Digit Application Number:", placeholder="e.g., 12345678").strip()
    submit_button = st.button(label="Verify Application Status")

    if submit_button and search_query:
        # Spinner frame container for input entries
        with st.spinner("⏳ Analyzing data registries..."):
            ptime.sleep(1.2)
            match = df[df["application_number"] == search_query]
            
        if not match.empty:
            app_id = str(match["application_number"].values[0])
            status = str(match["decision_status"].values[0]).strip()
            
            st.metric(label="Application Number Reference", value=app_id)
            
            if "approve" in status.lower():
                st.balloons()
                st.success(f"🍏 **Status: Approved**\n\nCongratulations! Your application has been approved by the New Delhi Visa Office. Safe travels on your journey ahead!")
            else:
                st.error(f"🚨 **Status: Refused**\n\nYour application has been returned with a refusal decision. Please coordinate directly with your visa processing handler.")
        else:
            st.warning("⚠️ No current record found matching that Application Number in today's batch updates.")
            
    st.markdown("---")
    last_update_ts = datetime.fromtimestamp(os.path.getmtime(JSON_FILE)).strftime("%Y-%m-%d %I:%M %p")
    st.caption(f"System Operational Ledger Cache Sync Frame: {last_update_ts}")
else:
    st.error("Service Temporarily Unavailable: Local record registers are out of sync.")
