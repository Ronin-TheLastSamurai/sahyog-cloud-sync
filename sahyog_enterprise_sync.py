import os
import sys
import time
import re
import json
import logging
import base64
from datetime import datetime
from urllib.parse import urljoin
import requests
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from pypdf import PdfWriter

# ==========================================
# CLOUD SECRETS & CREDENTIALS
# ==========================================
BOT_TOKEN = os.environ.get("BOT_TOKEN")
CHAT_ID = os.environ.get("CHAT_ID")
SAHYOG_USER = os.environ.get("SAHYOG_USER")
SAHYOG_PASS = os.environ.get("SAHYOG_PASS")

if not all([BOT_TOKEN, CHAT_ID, SAHYOG_USER, SAHYOG_PASS]):
    print("❌ CRITICAL: Missing Environment Variables. Ensure GitHub Secrets are set.")
    sys.exit(1)

# ==========================================
# CONFIGURATION & LOGGING
# ==========================================
CONFIG_FILE = "config.json"
LEDGER_FILE = "state_ledger.json"
OVERRIDE_FILE = "mapping_overrides.xlsx"

def load_or_create_config():
    default_config = {
        "recency_threshold_hours": 6,
        "urgent_threshold_days": 9
    }
    if not os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'w') as f:
            json.dump(default_config, f, indent=4)
        return default_config
    with open(CONFIG_FILE, 'r') as f:
        return json.load(f)

CONFIG = load_or_create_config()

log_filename = f"sahyog_sync_{datetime.now().strftime('%Y-%m-%d')}.log"
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[logging.FileHandler(log_filename, encoding='utf-8'), logging.StreamHandler(sys.stdout)]
)

# ==========================================
# HARDCODED MAPPING DATA
# ==========================================
BLOCK_DIVISION_MAP = {
    'Amnour': 'CHAPRA (EAST)', 'Dariapur': 'CHAPRA (EAST)', 'Dighwara': 'CHAPRA (EAST)', 
    'Ishupur': 'CHAPRA (EAST)', 'Maker': 'CHAPRA (EAST)', 'Marhaura': 'CHAPRA (EAST)', 
    'Mashrakh': 'CHAPRA (EAST)', 'Panapur': 'CHAPRA (EAST)', 'Parsa': 'CHAPRA (EAST)', 
    'Sonepur': 'CHAPRA (EAST)', 'Taraiya': 'CHAPRA (EAST)', 'Baniapur': 'CHAPRA (WEST)', 
    'Chapra': 'CHAPRA (WEST)', 'Ekma': 'CHAPRA (WEST)', 'Garkha': 'CHAPRA (WEST)', 
    'Jalalpur': 'CHAPRA (WEST)', 'Lahladpur': 'CHAPRA (WEST)', 'Manjhi': 'CHAPRA (WEST)', 
    'Nagra': 'CHAPRA (WEST)', 'Revelganj': 'CHAPRA (WEST)'
}

PANCHAYAT_MAP = {
    ('Revelganj', 'MUHABAT PARSA'): ('EKMA', 'TAJPUR'), ('Ekma', 'EKMA(NP)'): ('EKMA', 'EKMA_NEW'),
    ('Ekma', 'BALLIYA'): ('EKMA', 'TAJPUR'), ('Ekma', 'HUSEPUR'): ('EKMA', 'EKMA_NEW'),
    ('Ekma', 'Parsa Utari'): ('EKMA', 'EKMA_NEW'), ('Ekma', 'AMDARHI'): ('EKMA', 'TAJPUR'),
    ('Ekma', 'ASAHANI'): ('EKMA', 'TAJPUR'), ('Ekma', 'ATARSAN'): ('EKMA', 'TAJPUR'),
    ('Ekma', 'BHANPURA'): ('EKMA', 'TAJPUR'), ('Ekma', 'BHORHO PUR'): ('EKMA', 'EKMA_NEW'),
    ('Ekma', 'CHANCHAURA'): ('EKMA', 'EKMA_NEW'), ('Ekma', 'PARSA DAKHIN'): ('EKMA', 'EKMA_NEW'),
    ('Ekma', 'DEOPURA'): ('EKMA', 'TAJPUR'), ('Ekma', 'EKSAR'): ('EKMA', 'EKMA_NEW'),
    ('Ekma', 'HANS RAJ PUR'): ('EKMA', 'EKMA_NEW'), ('Ekma', 'MANE'): ('EKMA', 'EKMA_NEW'),
    ('Ekma', 'PACHUA'): ('EKMA', 'EKMA_NEW'), ('Ekma', 'PARSA PURVI'): ('EKMA', 'EKMA_NEW'),
    ('Ekma', 'PHUCHATI KALA'): ('EKMA', 'EKMA_NEW'), ('Ekma', 'RASUL PUR'): ('EKMA', 'TAJPUR'),
    ('Manjhi', 'BHAGAUNA  NACHAP'): ('EKMA', 'TAJPUR'), ('Manjhi', 'BHALUA BUJURG'): ('EKMA', 'TAJPUR'),
    ('Manjhi', 'CHEFUL'): ('EKMA', 'TAJPUR'), ('Manjhi', 'GOBARAHI'): ('EKMA', 'TAJPUR'),
    ('Manjhi', 'MAHAMADPUR'): ('EKMA', 'TAJPUR'), ('Manjhi', 'MATIYAR'): ('EKMA', 'TAJPUR'),
    ('Manjhi', 'MOBARAKPUR'): ('EKMA', 'TAJPUR'), ('Manjhi', 'SITALPUR'): ('EKMA', 'TAJPUR'),
    ('Manjhi', 'TAJPUR'): ('EKMA', 'TAJPUR'), ('Amnour', 'HUSEPUR'): ('EKMA', 'EKMA_NEW'),
    ('Ekma', 'EKMA'): ('EKMA', 'EKMA_NEW'), ('Garkha', 'KUDAR BADHA'): ('CHAPRA(RURAL)', 'BASANT'),
    ('Ekma', 'RAMPUR  BINALAL'): ('EKMA', 'DAUDPUR_NEW'), ('Manjhi', 'BALESHARA'): ('EKMA', 'DAUDPUR_NEW'),
    ('Manjhi', 'DAUDPUR'): ('EKMA', 'DAUDPUR_NEW'), ('Manjhi', 'ENAITPUR'): ('EKMA', 'DAUDPUR_NEW'),
    ('Manjhi', 'LEGUAR'): ('EKMA', 'DAUDPUR_NEW'), ('Manjhi', 'NASIRA'): ('EKMA', 'DAUDPUR_NEW'),
    ('Manjhi', 'JAITPUR'): ('EKMA', 'MANJHI_NEW'), ('Manjhi', 'BAREJA'): ('EKMA', 'MANJHI_NEW'),
    ('Manjhi', 'BANGARA'): ('EKMA', 'MANJHI_NEW'), ('Manjhi', 'GHORHAT'): ('EKMA', 'MANJHI_NEW'),
    ('Manjhi', 'KAWARU DHAWARU'): ('EKMA', 'MANJHI_NEW'), ('Manjhi', 'MADAN SATH'): ('EKMA', 'MANJHI_NEW'),
    ('Manjhi', 'MANJHI PASCHIMI'): ('EKMA', 'MANJHI_NEW'), ('Manjhi', 'MANJHI PURVI'): ('EKMA', 'MANJHI_NEW'),
    ('Manjhi', 'MARAHAN'): ('EKMA', 'MANJHI_NEW'), ('Manjhi', 'SONBARSA'): ('EKMA', 'MANJHI_NEW'),
    ('Manjhi', 'DUMARI'): ('BANIYAPUR', 'NAGRA'), ('Nagra', 'DUMARI'): ('BANIYAPUR', 'NAGRA'),
    ('Taraiya', 'DUMARI'): ('BANIYAPUR', 'NAGRA'), ('Chapra', 'DUMARI'): ('CHAPRA(RURAL)', 'DORIGANJ'),
    ('Ekma', 'NAWADA'): ('EKMA', 'EKMA_NEW'), ('Jalalpur', 'NAWADA'): ('EKMA', 'EKMA_NEW'),
    ('Mashrakh', 'NAWADA'): ('EKMA', 'EKMA_NEW'), ('Chapra', 'KARINGA'): ('CHAPRA(RURAL)', 'CHAPRA(S)'),
    ('Chapra', 'LOHARI'): ('CHAPRA(RURAL)', 'CHAPRA(S)'), ('Chapra', 'MAUNA'): ('CHAPRA(RURAL)', 'CHAPRA(S)'),
    ('Chapra', 'BADALU TOLA'): ('CHAPRA(RURAL)', 'CHAPRA(S)'), ('Chapra', 'NAINI'): ('CHAPRA(RURAL)', 'CHAPRA(S)'),
    ('Chapra', 'PHAKULI'): ('CHAPRA(RURAL)', 'CHAPRA(S)'), ('Chapra', 'PURBARI TELPA'): ('CHAPRA(RURAL)', 'CHAPRA(S)'),
    ('Chapra', 'SARHA'): ('CHAPRA(RURAL)', 'CHAPRA(S)'), ('Chapra', 'TENUA'): ('CHAPRA(RURAL)', 'CHAPRA(S)'),
    ('Revelganj', 'DAKNIWARI CHAKKI'): ('CHAPRA(RURAL)', 'RIVILGANJ'), ('Revelganj', 'DILIYA RAHIM PUR'): ('CHAPRA(RURAL)', 'RIVILGANJ'),
    ('Revelganj', 'INAI'): ('CHAPRA(RURAL)', 'RIVILGANJ'), ('Revelganj', 'KACHANAR'): ('CHAPRA(RURAL)', 'RIVILGANJ'),
    ('Revelganj', 'KHAIRWAR'): ('CHAPRA(RURAL)', 'RIVILGANJ'), ('Revelganj', 'MUKRERA'): ('CHAPRA(RURAL)', 'RIVILGANJ'),
    ('Revelganj', 'SITAB DIYARA'): ('CHAPRA(RURAL)', 'RIVILGANJ'), ('Revelganj', 'TEKNIWAS'): ('CHAPRA(RURAL)', 'RIVILGANJ'),
    ('Garkha', 'BAJIT PUR'): ('CHAPRA(RURAL)', 'GARKHA'), ('Garkha', 'FERUSA'): ('CHAPRA(RURAL)', 'GARKHA'),
    ('Garkha', 'GARKHA'): ('CHAPRA(RURAL)', 'GARKHA'), ('Garkha', 'HASAN PURA'): ('CHAPRA(RURAL)', 'GARKHA'),
    ('Garkha', 'KOTHIA'): ('CHAPRA(RURAL)', 'GARKHA'), ('Garkha', 'MAHAMDA'): ('CHAPRA(RURAL)', 'GARKHA'),
    ('Garkha', 'MAKIM PUR'): ('CHAPRA(RURAL)', 'GARKHA'), ('Garkha', 'MIRPUR JUARA'): ('CHAPRA(RURAL)', 'GARKHA'),
    ('Garkha', 'MITHE PUR'): ('CHAPRA(RURAL)', 'GARKHA'), ('Garkha', 'MOTIRAJ PUR'): ('CHAPRA(RURAL)', 'GARKHA'),
    ('Garkha', 'PIRAUNA'): ('CHAPRA(RURAL)', 'GARKHA'), ('Garkha', 'SADH PUR'): ('CHAPRA(RURAL)', 'GARKHA'),
    ('Garkha', 'SARGATTI'): ('CHAPRA(RURAL)', 'GARKHA'), ('Garkha', 'ITWA'): ('CHAPRA(RURAL)', 'BASANT'),
    ('Garkha', 'JALAL BASANT'): ('CHAPRA(RURAL)', 'BASANT'), ('Garkha', 'MAHAMAD PUR'): ('CHAPRA(RURAL)', 'BASANT'),
    ('Garkha', 'MIRZA PUR'): ('CHAPRA(RURAL)', 'BASANT'), ('Garkha', 'MOAZAM PUR'): ('CHAPRA(RURAL)', 'BASANT'),
    ('Garkha', 'NARAO'): ('CHAPRA(RURAL)', 'BASANT'), ('Garkha', 'PACHPATIA'): ('CHAPRA(RURAL)', 'BASANT'),
    ('Garkha', 'RAM PUR'): ('CHAPRA(RURAL)', 'BASANT'), ('Garkha', 'SRIPAL BASANT'): ('CHAPRA(RURAL)', 'BASANT'),
    ('Chapra', 'BISHUNPURA'): ('CHAPRA(RURAL)', 'DORIGANJ'), ('Chapra', 'CHIRAND'): ('CHAPRA(RURAL)', 'DORIGANJ'),
    ('Chapra', 'JALALPUR'): ('CHAPRA(RURAL)', 'DORIGANJ'), ('Chapra', 'KHALPURA BALA'): ('CHAPRA(RURAL)', 'DORIGANJ'),
    ('Chapra', 'KOTAWA PATTI RAMPUR'): ('CHAPRA(RURAL)', 'DORIGANJ'), ('Chapra', 'MAHARAJGANJ'): ('CHAPRA(RURAL)', 'DORIGANJ'),
    ('Chapra', 'BARHARA MAHAJI'): ('CHAPRA(RURAL)', 'DORIGANJ'), ('Chapra', 'BHAIROPUR NIZAMAT'): ('CHAPRA(RURAL)', 'DORIGANJ'),
    ('Chapra', 'MUSEPUR'): ('CHAPRA(RURAL)', 'DORIGANJ'), ('Chapra', 'RAIPUR BINGAWAN'): ('CHAPRA(RURAL)', 'DORIGANJ'),
    ('Chapra', 'SHERPUR'): ('CHAPRA(RURAL)', 'DORIGANJ'), ('Nagra', 'NAGRA'): ('BANIYAPUR', 'NAGRA'),
    ('Nagra', 'APHAUR'): ('BANIYAPUR', 'NAGRA'), ('Nagra', 'DHUPNAGAR   DHOBWAL'): ('BANIYAPUR', 'NAGRA'),
    ('Nagra', 'JAGADISHPUR'): ('BANIYAPUR', 'NAGRA'), ('Nagra', 'KADIPUR'): ('BANIYAPUR', 'NAGRA'),
    ('Nagra', 'KHAIRA'): ('BANIYAPUR', 'NAGRA'), ('Nagra', 'KOREA'): ('BANIYAPUR', 'NAGRA'),
    ('Nagra', 'TAKIA'): ('BANIYAPUR', 'NAGRA'), ('Nagra', 'TUJAR PUR'): ('BANIYAPUR', 'NAGRA'),
    ('Baniapur', 'BANIAPUR'): ('BANIYAPUR', 'BANIYAPUR'), ('Baniapur', 'KAMTA'): ('BANIYAPUR', 'BANIYAPUR'),
    ('Baniapur', 'BEDAULI'): ('BANIYAPUR', 'BANIYAPUR'), ('Baniapur', 'BHUSHAW'): ('BANIYAPUR', 'BANIYAPUR'),
    ('Baniapur', 'HARPUR'): ('BANIYAPUR', 'BANIYAPUR'), ('Baniapur', 'KANHAULI MANOHAR'): ('BANIYAPUR', 'BANIYAPUR'),
    ('Baniapur', 'KARHI'): ('BANIYAPUR', 'BANIYAPUR'), ('Baniapur', 'KRAH'): ('BANIYAPUR', 'BANIYAPUR'),
    ('Baniapur', 'PAIGAMBAR PUR'): ('BANIYAPUR', 'BANIYAPUR'), ('Baniapur', 'PIRAUTA KHAS'): ('BANIYAPUR', 'BANIYAPUR'),
    ('Baniapur', 'PITHAURI'): ('BANIYAPUR', 'BANIYAPUR'), ('Baniapur', 'SARAYA'): ('BANIYAPUR', 'BANIYAPUR'),
    ('Lahladpur', 'BANPURA'): ('BANIYAPUR', 'LAHLADPUR'), ('Lahladpur', 'BASAHI'): ('BANIYAPUR', 'LAHLADPUR'),
    ('Lahladpur', 'DANDAS PUR'): ('BANIYAPUR', 'LAHLADPUR'), ('Lahladpur', 'DAYAL PUR'): ('BANIYAPUR', 'LAHLADPUR'),
    ('Lahladpur', 'KATAIYA'): ('BANIYAPUR', 'LAHLADPUR'), ('Lahladpur', 'KISUN PURLAVWAR'): ('BANIYAPUR', 'LAHLADPUR'),
    ('Lahladpur', 'MIRZAPUR'): ('BANIYAPUR', 'LAHLADPUR'), ('Lahladpur', 'PURUSHOTIMPUR'): ('BANIYAPUR', 'LAHLADPUR'),
    ('Marhaura', 'MIRZAPUR'): ('BANIYAPUR', 'LAHLADPUR'), ('Baniapur', 'LAUA KALA'): ('BANIYAPUR', 'LAHLADPUR'),
    ('Baniapur', 'Manikpura'): ('BANIYAPUR', 'LAHLADPUR'), ('Jalalpur', 'ANAWAL'): ('BANIYAPUR', 'JALALPUR'),
    ('Jalalpur', 'ASHOK NAGAR'): ('BANIYAPUR', 'JALALPUR'), ('Jalalpur', 'BHAT KESAR'): ('BANIYAPUR', 'JALALPUR'),
    ('Jalalpur', 'BISHUN PURI'): ('BANIYAPUR', 'JALALPUR'), ('Jalalpur', 'DEURIA'): ('BANIYAPUR', 'JALALPUR'),
    ('Jalalpur', 'KISHUNPUR'): ('BANIYAPUR', 'JALALPUR'), ('Jalalpur', 'KOPA'): ('BANIYAPUR', 'JALALPUR'),
    ('Jalalpur', 'KUMNA'): ('BANIYAPUR', 'JALALPUR'), ('Jalalpur', 'MADHO PUR'): ('BANIYAPUR', 'JALALPUR'),
    ('Jalalpur', 'RAMPUR NUR NAGAR'): ('BANIYAPUR', 'JALALPUR'), ('Jalalpur', 'REWARI'): ('BANIYAPUR', 'JALALPUR'),
    ('Jalalpur', 'SAKARDIH'): ('BANIYAPUR', 'JALALPUR'), ('Jalalpur', 'SAMAHUTA'): ('BANIYAPUR', 'JALALPUR'),
    ('Jalalpur', 'SANWARI'): ('BANIYAPUR', 'JALALPUR'), ('Baniapur', 'BHITHI SHAHABUDDIN'): ('BANIYAPUR', 'BANIYAPUR_II'),
    ('Baniapur', 'DHANGAHRA'): ('BANIYAPUR', 'BANIYAPUR_II'), ('Baniapur', 'SATUA'): ('BANIYAPUR', 'BANIYAPUR_II'),
    ('Baniapur', 'DHAWRI'): ('BANIYAPUR', 'BANIYAPUR_II'), ('Baniapur', 'GOAPIPARPATI'): ('BANIYAPUR', 'BANIYAPUR_II'),
    ('Baniapur', 'MANOPALI'): ('BANIYAPUR', 'BANIYAPUR_II'), ('Baniapur', 'MARICHA'): ('BANIYAPUR', 'BANIYAPUR_II'),
    ('Baniapur', 'RAMDHNAW'): ('BANIYAPUR', 'BANIYAPUR_II'), ('Baniapur', 'SAHAJITPUR'): ('BANIYAPUR', 'BANIYAPUR_II'),
    ('Baniapur', 'SISAI'): ('BANIYAPUR', 'BANIYAPUR_II'), ('Baniapur', 'SURAUDHA'): ('BANIYAPUR', 'BANIYAPUR_II'),
    ('Chapra', 'Chapra (Nagar Parishad)'): ('CHAPRA(URBAN)', 'CHAPRA(URBAN)')
}

# ==========================================
# TELEGRAM COMMUNICATIONS
# ==========================================
def send_telegram_message(message):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
    payload = {"chat_id": CHAT_ID, "text": message}
    try: requests.post(url, data=payload)
    except: pass

def send_telegram_photo(photo_path, caption=""):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendPhoto"
    try:
        with open(photo_path, 'rb') as photo:
            payload = {'chat_id': CHAT_ID, 'caption': caption}
            files = {'photo': photo}
            requests.post(url, data=payload, files=files)
    except Exception as e: logging.error(f"Failed to send photo: {e}")

def send_telegram_document(file_path):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendDocument"
    if not os.path.exists(file_path): return
    try:
        with open(file_path, 'rb') as file_data:
            payload = {'chat_id': CHAT_ID}
            files = {'document': file_data}
            requests.post(url, data=payload, files=files)
            logging.info(f"📤 Uploaded {os.path.basename(file_path)} to Telegram.")
    except Exception as e: logging.error(f"Failed to send file: {e}")

def wait_for_telegram_reply(prompt_message):
    send_telegram_message(prompt_message)
    logging.info("⏳ Waiting for Telegram reply...")
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/getUpdates"
    
    response = requests.get(url).json()
    last_update_id = response["result"][-1]["update_id"] if response.get("result") else 0

    while True:
        poll_url = f"{url}?offset={last_update_id + 1}&timeout=5"
        try:
            resp = requests.get(poll_url).json()
            if resp.get("result"):
                for update in resp["result"]:
                    last_update_id = update["update_id"]
                    if "message" in update and "text" in update["message"]:
                        if str(update["message"]["chat"]["id"]) == str(CHAT_ID):
                            return update["message"]["text"].strip()
        except: pass
        time.sleep(2)

# ==========================================
# LEDGER & STATE MANAGEMENT
# ==========================================
def save_ledger(data_list):
    with open(LEDGER_FILE, 'w', encoding='utf-8') as f:
        json.dump(data_list, f, ensure_ascii=False, indent=4)

def check_for_resume():
    if os.path.exists(LEDGER_FILE):
        logging.warning("🚨 INCOMPLETE RUN DETECTED")
        choice = wait_for_telegram_reply("🚨 INCOMPLETE RUN DETECTED.\nA previous session was interrupted.\nDo you want to RESUME exactly where you left off? (y/n):").lower()
        if choice in ['y', 'yes']:
            with open(LEDGER_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            send_telegram_message(f"✅ Resuming... Loaded {len(data)} records from memory ledger.")
            return data, True
        else:
            os.remove(LEDGER_FILE)
            send_telegram_message("🗑️ Old ledger deleted. Starting a fresh run.")
    return [], False

def load_overrides():
    overrides = {}
    if os.path.exists(OVERRIDE_FILE):
        try:
            df = pd.read_excel(OVERRIDE_FILE).fillna("")
            for _, row in df.iterrows():
                ref = str(row.get('Registration No.', '')).strip()
                if ref:
                    overrides[ref] = {
                        'Subdivision': str(row.get('Subdivision', '')).strip(),
                        'Section': str(row.get('Section', '')).strip()
                    }
        except Exception as e: logging.error(f"Failed to read overrides: {e}")
    return overrides

# ==========================================
# CLICK HELPERS & TIMEOUT RECOVERY
# ==========================================
def safe_click(driver, element):
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", element)
        time.sleep(0.5) 
        driver.execute_script("arguments[0].click();", element)
    except: pass

def verify_page_state(driver):
    for attempt in range(3):
        try:
            page_src = driver.page_source.lower()
            if "server error" in page_src or "500 - internal server error" in page_src:
                wait_time = 5 if attempt == 0 else 300
                logging.error(f"Server Error! Retry {attempt+1}/3. Waiting {wait_time}s...")
                time.sleep(wait_time)
                driver.refresh()
                time.sleep(3)
            else: return True
        except: pass
    
    # Pause via Telegram instead of CMD input
    wait_for_telegram_reply("❌ CRITICAL: Server Error persists. Manual intervention required.\nReply 'continue' once you believe the portal is stable.")
    return True

def check_session(driver):
    try:
        driver.find_element(By.ID, "ctl00_lblwelcome")
        return True
    except:
        logging.warning("🚨 SESSION EXPIRED! Need Re-login.")
        send_telegram_message("🚨 [STATUS] SESSION EXPIRED! Triggering re-login flow...")
        perform_login(driver)
        return True

# ==========================================
# AUTHENTICATION ENGINE
# ==========================================
def perform_login(driver):
    while True:
        try:
            driver.get("https://sahyog.bihar.gov.in/Sahyog/LoginAdm.aspx")
            time.sleep(4)
            
            captcha_element = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_imgCaptcha")
            captcha_path = "captcha_screenshot.png"
            captcha_element.screenshot(captcha_path)
            
            send_telegram_photo(captcha_path, "🚨 SERVER WAKING UP: Please reply with this Captcha code.")
            captcha_text = wait_for_telegram_reply("Enter Captcha:")
            
            driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtUserName").send_keys(SAHYOG_USER)
            driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtPassword").send_keys(SAHYOG_PASS)
            driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtCode").send_keys(captcha_text)
            
            driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnLogin").click()
            time.sleep(5)
            
            if len(driver.find_elements(By.XPATH, "//a[contains(text(), 'Grievance Action')]")) > 0:
                send_telegram_message("✅ Login Successful! Resuming data capture...")
                break
            else:
                send_telegram_message("❌ Login Failed (Incorrect Captcha/Credentials). Retrying...")
        except Exception as e:
            logging.error(f"Login error: {e}")
            time.sleep(5)

# ==========================================
# POST-PROCESSING ENGINE
# ==========================================
def clean_for_excel(val):
    if not isinstance(val, str): return val
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', val)

def atomic_save_excel(df, final_path):
    temp_path = final_path.replace(".xlsx", "_tmp.xlsx") 
    try:
        df.to_excel(temp_path, index=False, engine='openpyxl')
        os.replace(temp_path, final_path)
        return True
    except Exception:
        if os.path.exists(temp_path): os.remove(temp_path)
        return False

def run_post_processing(master_logged_rows, target_output_dir, timestamp):
    send_telegram_message("⚙️ Commencing Data Split & Master Excel Generation...")
    
    combined_data, west_data, east_data, other_data, urgent_data = [], [], [], [], []
    overrides = load_overrides()

    for row in master_logged_rows:
        raw_date = str(row.get("Registration Date", ""))
        clean_date = raw_date.split(" ")[0] if raw_date else ""
        
        block = str(row.get("Block Name", "")).replace("nan", "").strip()
        panchayat = str(row.get("Panchayat Name", "")).replace("nan", "").strip()
        ref_no = str(row.get("Registration No.", "")).strip()
        
        if ref_no in overrides:
            subdiv, section = overrides[ref_no]['Subdivision'], overrides[ref_no]['Section']
        else:
            subdiv, section = PANCHAYAT_MAP.get((block, panchayat), ("", ""))
            if not subdiv and block == "Chapra": subdiv, section = "CHAPRA(RURAL)", "UNKNOWN_SECTION"

        formatted_row = {
            "Type": row.get("Type", ""), "Category": row.get("Category", ""), "Status": row.get("Status", ""),
            "Subdivision": subdiv, "Section": section, "Registration No.": ref_no,
            "Applicant Name": row.get("Applicant Name", ""), "Registration Date": clean_date,
            "Mobile Number": row.get("Mobile Number", ""), "Block Name": block,
            "Panchayat Name": panchayat, "Grievance Type": row.get("Grievance Type", ""),
            "Grievance Description": row.get("Grievance Description", ""), "Delegated To": row.get("Delegated To", ""),
            "Action Status": row.get("Action Status", ""), "Action Date": row.get("Action Date", ""),
            "Officer Remarks": row.get("Officer Remarks", ""), "Pending Days": row.get("Pending Days", 0),
            "Last Updated Time": row.get("Last Updated Time", "")
        }

        for key in formatted_row: formatted_row[key] = clean_for_excel(formatted_row[key])

        combined_data.append(formatted_row)
        try: days_pending = int(float(row.get("Pending Days", 0)))
        except: days_pending = 0
            
        if days_pending >= CONFIG["urgent_threshold_days"]: urgent_data.append(formatted_row)
        division = BLOCK_DIVISION_MAP.get(block, "OTHER")
        
        if division == "CHAPRA (WEST)": west_data.append(formatted_row)
        elif division == "CHAPRA (EAST)": east_data.append(formatted_row)
        else: other_data.append(formatted_row)

    final_cols = list(formatted_row.keys())
    generated_files = []

    if combined_data:
        path = os.path.join(target_output_dir, f"Sahyog_Shivir_{timestamp}_Combined.xlsx")
        if atomic_save_excel(pd.DataFrame(combined_data)[final_cols], path): generated_files.append(path)
        
    if west_data:
        path = os.path.join(target_output_dir, f"Chapra West_{timestamp}.xlsx")
        if atomic_save_excel(pd.DataFrame(west_data)[final_cols], path): generated_files.append(path)
        
    if east_data:
        path = os.path.join(target_output_dir, f"Chapra East_{timestamp}.xlsx")
        if atomic_save_excel(pd.DataFrame(east_data)[final_cols], path): generated_files.append(path)
        
    if urgent_data:
        path = os.path.join(target_output_dir, f"URGENT_COMPLIANCE_{timestamp}.xlsx")
        if atomic_save_excel(pd.DataFrame(urgent_data)[final_cols], path): generated_files.append(path)

    # Deliver final files via Telegram
    send_telegram_message("📁 Uploading Final Reports to Telegram...")
    for file_path in generated_files:
        send_telegram_document(file_path)

# ==========================================
# AUDIT & EXTRACTION SYSTEMS
# ==========================================
def get_dropdown_options(driver, element_id):
    try:
        select_elem = Select(driver.find_element(By.ID, element_id))
        options_list = []
        for opt in select_elem.options:
            val = opt.get_attribute("value")
            txt = opt.text.strip()
            if val and val != "0" and "--" not in txt:
                options_list.append({"value": val, "text": txt})
        return options_list
    except: return []

def perform_scraping_cycle(driver, main_tab, target_output_dir, skip_list, master_logged_rows, audit_log):
    WebDriverWait(driver, 30).until(lambda d: len(d.find_elements(By.ID, "ctl00_ContentPlaceHolder1_ddlFilterType")) > 0)
    type_options = get_dropdown_options(driver, "ctl00_ContentPlaceHolder1_ddlFilterType")
    complaint_options = get_dropdown_options(driver, "ctl00_ContentPlaceHolder1_ddlFilterComplaint")
    status_options = get_dropdown_options(driver, "ctl00_ContentPlaceHolder1_ddlFilterStatus")

    for t_opt in type_options:
        for c_opt in complaint_options:
            for s_opt in status_options:
                try:
                    driver.switch_to.window(main_tab)
                    combo_name = f"{t_opt['text']} > {c_opt['text']} > {s_opt['text']}"
                    
                    Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlFilterType")).select_by_value(t_opt['value'])
                    Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlFilterComplaint")).select_by_value(c_opt['value'])
                    Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlFilterStatus")).select_by_value(s_opt['value'])
                    time.sleep(2) 

                    delegated_labels = driver.find_elements(By.CSS_SELECTOR, ".delegated-panel .delegated-radio label")
                    
                    if delegated_labels:
                        for i in range(len(delegated_labels)):
                            fresh_labels = driver.find_elements(By.CSS_SELECTOR, ".delegated-panel .delegated-radio label")
                            if i < len(fresh_labels):
                                safe_click(driver, fresh_labels[i])
                                time.sleep(2) 
                                try: Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlTopFilter")).select_by_value("0")
                                except: pass
                                safe_click(driver, driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnViewNormal"))
                                time.sleep(4)
                                verify_page_state(driver)
                                run_extraction_sequence(driver, main_tab, target_output_dir, skip_list, master_logged_rows, t_opt, c_opt, s_opt)
                    else:
                        try: Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlTopFilter")).select_by_value("0")
                        except: pass
                        safe_click(driver, driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnViewNormal"))
                        time.sleep(4)
                        verify_page_state(driver)
                        run_extraction_sequence(driver, main_tab, target_output_dir, skip_list, master_logged_rows, t_opt, c_opt, s_opt)

                except Exception as e: continue

def run_extraction_sequence(driver, main_tab, target_output_dir, skip_list, master_logged_rows, t_opt, c_opt, s_opt):
    complaint_cards = driver.find_elements(By.CSS_SELECTOR, ".list-container .complaint-card")
    total_cards = len(complaint_cards)
    if total_cards == 0: return 

    for card_idx in range(total_cards):
        check_session(driver)
        verify_page_state(driver)
        try:
            driver.switch_to.window(main_tab)
            WebDriverWait(driver, 15).until(lambda d: len(d.find_elements(By.CSS_SELECTOR, ".list-container .complaint-card")) > card_idx)
            card = driver.find_elements(By.CSS_SELECTOR, ".list-container .complaint-card")[card_idx]
            ref_match = re.search(r'REF\d+', card.text)
            expected_ref = ref_match.group(0) if ref_match else ""
            
            is_old = False
            for row in master_logged_rows:
                if row.get("Registration No.") == expected_ref:
                    try:
                        last_update = pd.to_datetime(row.get("Last Updated Time", "2000-01-01"))
                        if (datetime.now() - last_update).total_seconds() / 3600 < CONFIG["recency_threshold_hours"]:
                            is_old = True
                            break
                    except: pass
                    is_old = True
                    break
            
            if is_old and (datetime.now() - last_update).total_seconds() / 3600 < CONFIG["recency_threshold_hours"]: continue

            try:
                delay_elem = card.find_element(By.CSS_SELECTOR, ".delay-box span").text
                pending_days = int(re.search(r'\d+', delay_elem).group()) if re.search(r'\d+', delay_elem) else 0
            except: pending_days = 0
            
            safe_click(driver, card.find_element(By.CSS_SELECTOR, ".stretched-link"))
            WebDriverWait(driver, 10).until(lambda d: d.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblAck").text.strip() == expected_ref)

            data = {
                "Type": t_opt['text'], "Category": c_opt['text'], "Status": s_opt['text'], "Registration No.": expected_ref, 
                "Applicant Name": driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblApplicant").text.strip(),
                "Mobile Number": driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblMobile").text.strip(),
                "Block Name": driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblAppBlock").text.strip(),
                "Panchayat Name": driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblAppPanch").text.strip(),
                "Grievance Description": driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblDesc").text.strip() if len(driver.find_elements(By.ID, "ctl00_ContentPlaceHolder1_lblDesc")) > 0 else "N/A",
                "Delegated To": driver.find_element(By.XPATH, "//table[@id='ctl00_ContentPlaceHolder1_grdDelegate']//tr[2]/td[2]/div").text.strip() if len(driver.find_elements(By.ID, "ctl00_ContentPlaceHolder1_grdDelegate")) > 0 else "Not Delegated",
                "Action Status": driver.find_element(By.XPATH, "//table[@id='ctl00_ContentPlaceHolder1_grdDelegate']//tr[2]/td[3]").text.strip() if len(driver.find_elements(By.ID, "ctl00_ContentPlaceHolder1_grdDelegate")) > 0 else "N/A",
                "Action Date": driver.find_element(By.XPATH, "//table[@id='ctl00_ContentPlaceHolder1_grdDelegate']//tr[2]/td[4]").text.strip() if len(driver.find_elements(By.ID, "ctl00_ContentPlaceHolder1_grdDelegate")) > 0 else "N/A",
                "Officer Remarks": driver.find_element(By.XPATH, "//table[@id='ctl00_ContentPlaceHolder1_grdDelegate']//tr[2]/td[5]/div").text.strip() if len(driver.find_elements(By.ID, "ctl00_ContentPlaceHolder1_grdDelegate")) > 0 else "N/A",
                "Pending Days": pending_days, "Last Updated Time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }

            if is_old:
                for row in master_logged_rows:
                    if row.get("Registration No.") == expected_ref:
                        row.update(data)
                        break
            else:
                master_logged_rows.append(data)
                skip_list.add(expected_ref)

            save_ledger(master_logged_rows)
            driver.switch_to.window(main_tab)
            safe_click(driver, driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnBack"))
            time.sleep(2)
        except Exception: 
            try:
                driver.switch_to.window(main_tab)
                safe_click(driver, driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnBack"))
            except: pass
            continue

# ==========================================
# RUN MAIN ASSEMBLY
# ==========================================
def main():
    send_telegram_message("🚀 Sahyog Enterprise Multi-Login Automation (Cloud Version) Starting...")
    
    master_logged_rows, _ = check_for_resume()
    skip_list = set(row.get("Registration No.") for row in master_logged_rows if row.get("Registration No."))
    audit_log = []
    
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    chrome_options.add_argument("--window-size=1920,1080")
    
    driver = webdriver.Chrome(options=chrome_options)
    driver.set_page_load_timeout(45)

    timestamp = datetime.now().strftime('%d-%m-%Y_%H%M')
    target_output_dir = os.path.join(os.getcwd(), f"Sahyog_Shivir_{timestamp}")
    os.makedirs(target_output_dir, exist_ok=True)

    while True:
        perform_login(driver)
        
        try: safe_click(driver, driver.find_element(By.XPATH, "//a[contains(text(), 'Grievance Action')]"))
        except: driver.get("https://sahyog.bihar.gov.in/Sahyog/IGRS_InnerPage/Common/ComplaintAction.aspx")
        
        verify_page_state(driver)
        perform_scraping_cycle(driver, driver.current_window_handle, target_output_dir, skip_list, master_logged_rows, audit_log)
        
        cont = wait_for_telegram_reply("✅ Data collection for this ID is complete.\nDo you wish to loop another ID? (y/n):").lower()
        if cont in ['y', 'yes']:
            try: driver.find_element(By.ID, "ctl00_LoginViewSTDPTDPADM_LoginStatusSTDPTDPADM").click()
            except: pass
        else: break

    run_post_processing(master_logged_rows, target_output_dir, timestamp)
    
    if os.path.exists(LEDGER_FILE): os.remove(LEDGER_FILE)
    send_telegram_message("🎉 Automation Sequence Completed Successfully. Ledger Cleared. Server shutting down.")
    driver.quit()

if __name__ == "__main__":
    main()