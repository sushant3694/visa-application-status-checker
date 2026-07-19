import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
import os
import urllib3
import json
import re
import base64
from datetime import datetime, time

# Suppress insecure request warnings from using verify=False
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

TARGET_URL = "https://www.ireland.ie/en/india/newdelhi/services/visas/processing-times-and-decisions/#visa-decisions"
LOCAL_FILE = "visa_decisions_latest.ods"
JSON_FILE = "visa_decisions.json"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
    "Connection": "keep-alive"
}

st.set_page_config(
    page_title="Ireland Visa Tracking Portal",
    page_icon="🇮🇪",
    layout="centered"
)

# Safely encode background image to Base64
def get_base64_image(image_path):
    if os.path.exists(image_path):
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    return ""

img_base64 = get_base64_image("background.avif")

# Advanced UI injection: Smooth animations, tech-glow button, and footer styling
st.markdown(f"""
    <!-- Import Plus Jakarta Sans from Google Fonts -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:ital,wght@0,300;0,400;0,600;0,700;0,800;1,400&display=swap" rel="stylesheet">

    <style>
    /* Global Font & Smooth Fade-in Page Animation */
    @keyframes fadeInPage {{
        from {{ opacity: 0; transform: translateY(15px); }}
        to {{ opacity: 1; transform: translateY(0); }}
    }}
    
    html, body, [class*="css"], .stApp {{
        font-family: 'Plus Jakarta Sans', sans-serif !important;
    }}
    
    .stApp {{
        background: linear-gradient(rgba(255, 255, 255, 0.92), rgba(0, 0, 0, 0.92)), 
                    url("data:image/avif;base64,{img_base64}");
        background-size: cover;
        background-position: center;
        background-attachment: fixed;
        animation: fadeInPage 0.8s ease-out forwards;
    }}
    
    /* Default Desktop Typography & Layout spacing */
    .block-container {{
        padding-left: 2rem !important;
        padding-right: 2rem !important;
    }}
    
    h1 {{
        color: #007A4E !important;
        font-weight: 800;
        text-align: center;
        letter-spacing: -0.5px;
        font-size: 2.5rem;
    }}
    
    .quote-box {{
        text-align: center;
        font-style: italic;
        color: #4A5568;
        margin-bottom: 25px;
        font-size: 1.1rem;
        line-height: 1.6;
        font-weight: 400;
    }}
    
    /* Catchy Modern Tech-Style Button Framework */
    div.stButton > button:first-child {{
        font-family: 'Plus Jakarta Sans', sans-serif !important;
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
    
    div.stButton > button:first-child:active {{
        transform: translateY(1px);
    }}
    
    /* Input field text adjustment */
    input {{
        font-family: 'Plus Jakarta Sans', sans-serif !important;
    }}
    
    .branding-footer {{
        text-align: center;
        margin-top: 50px;
        padding: 20px;
        font-size: 0.9rem;
        color: #718096;
        border-top: 1px solid rgba(0,0,0,0.05);
        font-weight: 400;
    }}
    
    .branding-footer a {{
        color: #007A4E !important;
        text-decoration: none;
        font-weight: 700;
        transition: color 0.3s ease;
    }}
    
    .branding-footer a:hover {{
        color: #FF883E !important;
        text-decoration: underline;
    }}

    /* Mobile & Tablet Responsive Overrides */
    @media (max-width: 768px) {{
        .block-container {{
            padding-left: 0.75rem !important;
            padding-right: 0.75rem !important;
            padding-top: 2rem !important;
        }}
        
        h1 {{
            font-size: 1.65rem !important;
            line-height: 1.3 !important;
        }}
        
        .quote-box {{
            font-size: 0.95rem !important;
            margin-bottom: 15px;
            padding: 0 5px;
        }}
        
        div.stButton > button:first-child {{
            padding: 14px !important;
            font-size: 0.95rem !important;
        }}
    }}
    </style>
""", unsafe_allow_html=True)

def fetch_latest_ods_url():
    try:
        response = requests.get(TARGET_URL, headers=HEADERS, timeout=15, verify=False)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        
        for link in soup.find_all('a', href=True):
            if link['href'].endswith('.ods'):
                url = link['href']
                if url.startswith('/'):
                    url = "https://www.ireland.ie" + url
                return url
    except Exception:
        pass
    return None

@st.cache_data(ttl=3600)
def download_and_convert_production(today_str):
    download_url = fetch_latest_ods_url()
    if download_url:
        try:
            file_response = requests.get(download_url, headers=HEADERS, timeout=30, verify=False)
            file_response.raise_for_status()
            with open(LOCAL_FILE, 'wb') as f:
                f.write(file_response.content)
            
            df = pd.read_excel(LOCAL_FILE, engine='odf', header=None)
            df = df.astype(str).replace(['nan', 'NaN', 'None'], '')
            raw_rows = df.values.tolist()
            
            cleaned_records = []
            for row in raw_rows:
                cells = [str(cell).strip() for cell in row if str(cell).strip()]
                app_num = None
                decision = "Unknown"
                
                for cell in cells:
                    if re.match(r'^\d{8}$', cell):
                        app_num = cell
                        break
                
                if app_num:
                    other_cells = [c for c in cells if c != app_num]
                    if other_cells:
                        decision = other_cells[-1]
                    
                    cleaned_records.append({
                        "application_number": app_num,
                        "decision_status": decision
                    })
            
            with open(JSON_FILE, 'w', encoding='utf-8') as f:
                json.dump(cleaned_records, f, ensure_ascii=False, indent=4)
            return True
        except Exception:
            pass
    return False

def get_production_dataset():
    now = datetime.now()
    today_str = now.strftime("%Y-%m-%d")
    json_exists = os.path.exists(JSON_FILE)
    
    should_download = False
    if not json_exists:
        should_download = True
    elif now.time() >= time(11, 5):
        last_modified_date = datetime.fromtimestamp(os.path.getmtime(JSON_FILE)).strftime("%Y-%m-%d")
        if last_modified_date != today_str:
            should_download = True
            
    if should_download:
        with st.spinner("✨ 'The future belongs to those who believe in the beauty of their dreams.' — Synchronizing current processing ledger..."):
            download_and_convert_production(today_str)
            
    if os.path.exists(JSON_FILE):
        try:
            with open(JSON_FILE, 'r', encoding='utf-8') as f:
                return pd.DataFrame(json.load(f))
        except Exception:
            return None
    return None

# --- UI Header Content ---
st.title("Ireland Visa Application Tracking System")
st.markdown("<div class='quote-box'>\"Céad Míle Fáilte\"- A hundred thousand welcomes. Charting your pathway to the Emerald Isle.</div>", unsafe_allow_html=True)

df = get_production_dataset()

if df is not None and not df.empty:
    with st.form(key="search_form"):
        search_query = st.text_input("Enter Your 8-Digit Application Number:", placeholder="e.g., 12345678").strip()
        submit_button = st.form_submit_button(label="Verify")

    if submit_button and search_query:
        match = df[df["application_number"] == search_query]
        
        if not match.empty:
            app_id = str(match["application_number"].values[0])
            status = str(match["decision_status"].values[0]).strip()
            status_lower = status.lower()
            
            st.markdown("### Application Information Record")
            st.metric(label="Application Number Reference", value=app_id)
            
            if "approve" in status_lower:
                st.balloons()
                st.success(f"🍏 **Status: {status}**\n\nCongratulations! Your application has been approved by the New Delhi Visa Office. Safe travels on your journey ahead!")
            elif "refuse" in status_lower or "deny" in status_lower:
                st.error(f"🚨 **Status: {status}**\n\nYour application has been returned with a refusal decision. Please coordinate directly with your visa processing handler for formal decision notifications and appeal parameters.")
            else:
                st.info(f"ℹ️ **Status: {status}**\n\nYour reference code was located with status flags: {status}.")
        else:
            st.warning("⚠️ No current record found matching that Application Number. The decision ledger may still be pending updates or scheduling queues.")
            
    st.markdown("---")
    last_update_ts = datetime.fromtimestamp(os.path.getmtime(JSON_FILE)).strftime("%Y-%m-%d %I:%M %p")
    st.caption(f"Data Synced : {last_update_ts}")
else:
    st.error("Service Temporarily Unavailable: Local record registers are out of sync.")

# --- Professional Marketing & Branding Signature Block ---
st.markdown("""
    <div class='branding-footer'>
        Built with precision by <a href='https://www.sushantthorat.com/' target='_blank'>Sushant Thorat</a>
    </div>
""", unsafe_allow_html=True)