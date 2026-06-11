import os
import sys
import time
import re
import json
import logging
import base64
import zipfile
import urllib3
import traceback
from datetime import datetime
from urllib.parse import urljoin
import requests
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from pypdf import PdfWriter
from selenium.common.exceptions import TimeoutException, WebDriverException, InvalidSessionIdException

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==========================================
# CLOUD SECRETS & CREDENTIALS
# ==========================================
BOT_TOKEN = os.environ.get("BOT_TOKEN")
CHAT_ID = os.environ.get("CHAT_ID")
SAHYOG_USER = os.environ.get("SAHYOG_USER")
SAHYOG_PASS = os.environ.get("SAHYOG_PASS")
WEBHOOK_URL = os.environ.get("WEBHOOK_URL")

if not all([BOT_TOKEN, CHAT_ID, SAHYOG_USER, SAHYOG_PASS]):
    print("❌ CRITICAL: Missing Environment Variables. Ensure GitHub Secrets are set.")
    sys.exit(1)

# ==========================================
# 1. CONFIGURATION & LOGGING
# ==========================================
timestamp_str = datetime.now().strftime('%Y-%m-%d_%H%M%S')
log_filename = f"SAHYOG_DETAILED_LOG_{timestamp_str}.txt"

file_handler = logging.FileHandler(log_filename, encoding='utf-8')
file_handler.setLevel(logging.DEBUG) 
file_formatter = logging.Formatter('%(asctime)s [%(levelname)s] [Line: %(lineno)d] %(message)s')
file_handler.setFormatter(file_formatter)

console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.INFO) 
console_formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
console_handler.setFormatter(console_formatter)

logging.basicConfig(level=logging.DEBUG, handlers=[file_handler, console_handler])
logging.getLogger("urllib3").setLevel(logging.WARNING)
logging.getLogger("selenium").setLevel(logging.WARNING)

CONFIG = {
    "excel_save_batch_size": 5, 
    "target_department": "Energy",
    "target_district": "SARAN",
    "urgent_threshold_days": 9
}

LEDGER_FILE = "state_ledger.json"

# ==========================================
# 2. PANCHAYAT MAPPING & EXCEL COLUMNS
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

FINAL_EXCEL_COLUMNS = [
    "S.No.", "Subdivision", "Section", "Registration No.", "Applicant Name", "Registration Date", 
    "Application Type", "Father's Name", "Mobile No", "Email",
    "Full Address", "Applicant District", "Applicant Block", "Applicant Panchayat", "Applicant Police Station", "Pincode",
    "Department Name", "Complaint Status", "Grievance Division", "Grievance District", "Grievance Sub Division", 
    "Grievance Block", "Grievance Panchayat", "Grievance Police Station", "Grievance Type", 
    "Designation", "Designation Level", "Delegated Status", "Pending Duration (Level 1)", 
    "Delegation Duration (Days)", "Grievance Description", 
    "Delegated To Officer", "Delegated Action Status", "Delegated Action Date", "Delegated Remarks",
    "History Officer Details", "History Action Date", "History Feedback", "History Remarks",
    "Last Updated Time"
]

# ==========================================
# WEBHOOK TOGGLE ENGINE
# ==========================================
def toggle_webhook(enable=True):
    if enable and WEBHOOK_URL:
        url = f"https://api.telegram.org/bot{BOT_TOKEN}/setWebhook?url={WEBHOOK_URL}"
        try: requests.get(url, verify=False)
        except: pass
        logging.info("🔗 Webhook re-enabled. Google is back in control.")
    else:
        url = f"https://api.telegram.org/bot{BOT_TOKEN}/deleteWebhook"
        try: requests.get(url, verify=False)
        except: pass
        logging.info("🔌 Webhook disabled. Python is now listening via polling.")

# ==========================================
# TELEGRAM COMMUNICATIONS (FLUSH QUEUE)
# ==========================================
def send_telegram_message(message):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
    payload = {"chat_id": CHAT_ID, "text": message}
    try: requests.post(url, data=payload, verify=False)
    except: pass

def send_telegram_photo(photo_path, caption=""):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendPhoto"
    try:
        with open(photo_path, 'rb') as photo:
            payload = {'chat_id': CHAT_ID, 'caption': caption}
            files = {'photo': photo}
            requests.post(url, data=payload, files=files, verify=False)
    except Exception as e: logging.error(f"Failed to send photo: {e}")

def send_telegram_document(file_path):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendDocument"
    if not os.path.exists(file_path): return
    try:
        with open(file_path, 'rb') as file_data:
            payload = {'chat_id': CHAT_ID}
            files = {'document': file_data}
            requests.post(url, data=payload, files=files, timeout=120, verify=False) 
            logging.info(f"📤 Uploaded {os.path.basename(file_path)} to Telegram.")
    except Exception as e: logging.error(f"Failed to send file: {e}")

def get_telegram_file(file_id):
    file_info = requests.get(f"https://api.telegram.org/bot{BOT_TOKEN}/getFile?file_id={file_id}", verify=False).json()
    file_path = file_info['result']['file_path']
    download_url = f"https://api.telegram.org/file/bot{BOT_TOKEN}/{file_path}"
    return requests.get(download_url, verify=False).content

def wait_for_telegram_reply_or_file(prompt_message):
    send_telegram_message(prompt_message)
    logging.info("⏳ Waiting for Telegram text or file...")
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/getUpdates"
    
    try:
        resp = requests.get(url, verify=False).json()
        if resp.get("result"):
            last_id = resp["result"][-1]["update_id"]
            requests.get(f"{url}?offset={last_id + 1}", verify=False)
            last_update_id = last_id
        else:
            last_update_id = 0
    except:
        last_update_id = 0

    while True:
        poll_url = f"{url}?offset={last_update_id + 1}&timeout=5"
        try:
            resp = requests.get(poll_url, verify=False).json()
            if resp.get("result"):
                for update in resp["result"]:
                    last_update_id = update["update_id"]
                    if "message" in update:
                        if str(update["message"]["chat"]["id"]) == str(CHAT_ID):
                            if "document" in update["message"]:
                                doc = update["message"]["document"]
                                file_id, file_name = doc["file_id"], doc["file_name"]
                                send_telegram_message(f"📥 Received file: {file_name}. Processing...")
                                requests.get(f"{url}?offset={last_update_id + 1}", verify=False) 
                                return {"type": "file", "content": get_telegram_file(file_id), "name": file_name}
                            elif "text" in update["message"]:
                                text = update["message"]["text"].strip()
                                send_telegram_message(f"✅ Received: {text}")
                                requests.get(f"{url}?offset={last_update_id + 1}", verify=False) 
                                return {"type": "text", "text": text}
        except: pass
        time.sleep(2)

# ==========================================
# 3. HELPER FUNCTIONS & SYNC LOCKS
# ==========================================
def clean_for_excel(val):
    if not isinstance(val, str) and not isinstance(val, float): return val
    if pd.isna(val): return ""
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', str(val)).strip()

def sanitize_filename(name):
    if not name: return "Unknown"
    return re.sub(r'[\\/*?:"<>|]', "", str(name)).strip().replace(" ", "_")

def safe_click(driver, element):
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", element)
        time.sleep(0.3)
        element.click() 
    except Exception as e:
        logging.debug(f"Native Click failed, attempting JS click: {e}")
        driver.execute_script("arguments[0].click();", element)

def wait_for_table(driver, table_id, timeout=25):
    try:
        WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, table_id)))
        return True
    except TimeoutException:
        return False

def wait_for_ajax(driver, timeout=30):
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return typeof Sys !== 'undefined' ? !Sys.WebForms.PageRequestManager.getInstance().get_isInAsyncPostBack() : true;")
        )
        time.sleep(0.5) 
    except Exception as e:
        logging.debug(f"AJAX wait timeout: {e}")

def close_extra_tabs(driver, keep_handles):
    current = driver.current_window_handle
    for handle in list(driver.window_handles):
        if handle not in keep_handles:
            try:
                driver.switch_to.window(handle)
                driver.close()
            except: pass
    try:
        if current in keep_handles: driver.switch_to.window(current)
        else: driver.switch_to.window(keep_handles[-1])
    except: pass

def parse_cell_data(text):
    data = {}
    for line in text.split('\n'):
        if ':-' in line:
            parts = line.split(':-', 1)
            data[parts[0].strip()] = parts[1].strip()
    return data

def is_valid_pdf(filepath):
    if not os.path.exists(filepath): return False
    try:
        with open(filepath, 'rb') as f: return f.read(4) == b'%PDF'
    except: return False

def get_subdivision_section(block, panchayat):
    block = str(block).strip()
    panchayat = str(panchayat).strip()
    subdiv, section = PANCHAYAT_MAP.get((block, panchayat), ("", ""))
    if not subdiv and block == "Chapra":
        subdiv, section = "CHAPRA(RURAL)", "UNKNOWN_SECTION"
    return subdiv, section

def verify_page_state(driver):
    for attempt in range(3):
        try:
            page_src = driver.page_source.lower()
            if "server error" in page_src or "500 - internal server error" in page_src:
                wait_time = 5 if attempt == 0 else 120
                logging.error(f"Server Error! Retry {attempt+1}/3. Waiting {wait_time}s...")
                time.sleep(wait_time)
                driver.refresh()
                time.sleep(3)
            else: return True
        except: pass
    msg = "❌ CRITICAL: Server Error persists. Manual intervention required.\nReply 'continue' once stable."
    wait_for_telegram_reply_or_file(msg)
    return True

def check_session(driver, current_user, current_pass):
    try:
        # PATCHED: Using flexible XPATH to bypass the missing 'ctl00_' UI update
        driver.find_element(By.XPATH, "//*[contains(@id, 'lblwelcome')]")
        return True
    except:
        msg = "🚨 SESSION EXPIRED! Triggering re-login flow..."
        logging.warning(msg)
        send_telegram_message(msg)
        perform_login(driver, current_user, current_pass)
        return True

# ==========================================
# AUTH & DYNAMIC LOGIN ENGINE
# ==========================================
def perform_login(driver, login_user, login_pass):
    max_retries = 3 
    for attempt in range(max_retries):
        try:
            driver.get("https://sahyog.bihar.gov.in/Sahyog/LoginAdm.aspx")
            time.sleep(4)
            
            # PATCHED: Updated IDs for the new website deployment (removed ctl00_)
            captcha_element = driver.find_element(By.ID, "ContentPlaceHolder1_imgCaptcha")
            captcha_path = "captcha_screenshot.png"
            captcha_element.screenshot(captcha_path)
            send_telegram_photo(captcha_path, f"🚨 SERVER WAKING UP (User: {login_user}): Please reply with this Captcha code. (Attempt {attempt+1}/{max_retries})")
            
            resp = wait_for_telegram_reply_or_file("Enter Captcha:")
            captcha_text = resp['text'] if resp['type'] == 'text' else ""
            
            driver.find_element(By.ID, "ContentPlaceHolder1_txtUserName").send_keys(login_user)
            driver.find_element(By.ID, "ContentPlaceHolder1_txtPassword").send_keys(login_pass)
            driver.find_element(By.ID, "ContentPlaceHolder1_txtCode").send_keys(captcha_text)
            driver.find_element(By.ID, "ContentPlaceHolder1_btnLogin").click()
            
            WebDriverWait(driver, 15).until(lambda d: len(d.find_elements(By.XPATH, "//a[contains(text(), 'Grievance Action')]")) > 0)
            logging.info(f"✅ Login Successful for {login_user}!")
            return True
            
        except Exception as e: 
            logging.warning(f"Login failed on attempt {attempt+1}: {e}")
            time.sleep(5)
            
    msg = "❌ CRITICAL: Login failed 3 times. Server is down or Captcha loop detected. Terminating to protect resources."
    send_telegram_message(msg)
    sys.exit(1)

# ==========================================
# 4. EXECUTIVE BLUE PDF ENGINE
# ==========================================
def generate_replica_pdf(driver, print_tab_handle, main_tab_handle, r_data, output_path):
    html_content = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 11px; color: #334155; padding: 15px; margin: 0; background-color: #ffffff; }}
            .header-main {{ text-align: center; font-weight: 800; font-size: 16px; margin-bottom: 5px; color: #1C4D8D; text-transform: uppercase; letter-spacing: 0.5px; }}
            .header-sub {{ text-align: center; font-size: 11px; margin-bottom: 12px; color: #64748b; }}
            table {{ width: 100%; border-collapse: collapse; margin-bottom: 10px; border: 1px solid #cbd5e1; }}
            th, td {{ border: 1px solid #cbd5e1; padding: 5px 8px; text-align: left; vertical-align: middle; line-height: 1.3; }}
            .section-title {{ font-weight: bold; background-color: #1C4D8D; color: #ffffff; text-align: left; padding: 6px 10px; font-size: 12px; letter-spacing: 0.5px; }}
            .label-cell {{ font-weight: 600; width: 22%; background-color: #f1f5f9; color: #0f172a; }}
            .val-cell {{ width: 28%; color: #334155; }}
            .val-cell-wide {{ width: 78%; color: #334155; }}
            .highlight-text {{ color: #b45309; font-weight: 600; }}
        </style>
    </head>
    <body>
        <div class="header-main">Sahyog Portal - RTMS | Govt. of Bihar<br>सहयोग पोर्टलः त्वरित अनुश्रवण प्रणाली</div>
        <div class="header-sub">The Sahyog Portal - RTMS, Government of Bihar</div>
        
        <table>
            <tr>
                <td class="label-cell">पंजीकरण संख्या (Reg.no.)</td><td class="val-cell highlight-text">{r_data.get('Registration No.', 'N/A')}</td>
                <td class="label-cell">पंजीकरण तिथि (Reg.Date)</td><td class="val-cell">{r_data.get('Registration Date', 'N/A')}</td>
            </tr>
        </table>

        <table>
            <tr><td colspan="4" class="section-title">आवेदक का विवरण (Applicant Details)</td></tr>
            <tr>
                <td class="label-cell">आवेदक (Applicant) :-</td><td class="val-cell">{r_data.get('Applicant Name', 'N/A')}</td>
                <td class="label-cell">पिता का नाम (Father's Name) :-</td><td class="val-cell">{r_data.get("Father's Name", 'N/A')}</td>
            </tr>
            <tr>
                <td class="label-cell">मोबाइल (Mob.No.) :-</td><td class="val-cell">{r_data.get('Mobile No', 'N/A')}</td>
                <td class="label-cell">ईमेल (Email) :-</td><td class="val-cell">{r_data.get('Email', '')}</td>
            </tr>
            <tr>
                <td class="label-cell">पता (Address) :-</td>
                <td colspan="3" class="val-cell-wide">
                    {r_data.get('Full Address', '')}, Dist:- {r_data.get('Applicant District', '')}, 
                    Block:- {r_data.get('Applicant Block', '')}, Panchayat:- {r_data.get('Applicant Panchayat', '')}, 
                    Pin:- {r_data.get('Pincode', '')}
                </td>
            </tr>
        </table>
        
        <table>
            <tr><td colspan="4" class="section-title">शिकायत का विवरण (Grievance Details)</td></tr>
            <tr>
                <td class="label-cell">प्रकार (Type) :-</td><td class="val-cell">{r_data.get('Grievance Type', 'N/A')}</td>
                <td class="label-cell">विभाग (Department) :-</td><td class="val-cell">{r_data.get('Department Name', 'N/A')}</td>
            </tr>
            <tr>
                <td class="label-cell">विवरण (Description) :-</td><td colspan="3" class="val-cell-wide">{r_data.get('Grievance Description', 'N/A')}</td>
            </tr>
            <tr>
                <td class="label-cell">डिवीजन (Division) :-</td><td class="val-cell">{r_data.get('Grievance Division', 'N/A')}</td>
                <td class="label-cell">उप-डिवीजन (Sub-Div) :-</td><td class="val-cell">{r_data.get('Grievance Sub Division', 'N/A')}</td>
            </tr>
            <tr>
                <td class="label-cell">जिला (District) :-</td><td class="val-cell">{r_data.get('Grievance District', 'N/A')}</td>
                <td class="label-cell">प्रखंड (Block) :-</td><td class="val-cell">{r_data.get('Grievance Block', 'N/A')}</td>
            </tr>
        </table>

        <table>
            <tr><td colspan="4" class="section-title">उक्त आवेदन आवश्यक कार्यवाही हेतु भेजा जाता है (Current Delegation)</td></tr>
            <tr>
                <td class="label-cell">अधिकारी (Delegated To) :-</td><td class="val-cell">{r_data.get('Delegated To Officer', 'N/A')}</td>
                <td class="label-cell">स्थिति (Status) :-</td><td class="val-cell">{r_data.get('Delegated Action Status', 'N/A')}</td>
            </tr>
            <tr>
                <td class="label-cell">टिप्पणी (Remarks) :-</td><td colspan="3" class="val-cell-wide">{r_data.get('Delegated Remarks', 'N/A')}</td>
            </tr>
        </table>

        <table>
            <tr><td colspan="4" class="section-title">अधिकारियों द्वारा की गई कार्रवाई (Officer Action History)</td></tr>
            <tr>
                <td class="label-cell">अधिकारी (Details) :-</td><td class="val-cell">{r_data.get('History Officer Details', 'N/A')}</td>
                <td class="label-cell">तिथि (Action Date) :-</td><td class="val-cell">{r_data.get('History Action Date', 'N/A')}</td>
            </tr>
            <tr>
                <td class="label-cell">प्रतिक्रिया (Feedback) :-</td><td class="val-cell">{r_data.get('History Feedback', 'N/A')}</td>
                <td class="label-cell">टिप्पणी (Remarks) :-</td><td class="val-cell">{r_data.get('History Remarks', 'N/A')}</td>
            </tr>
        </table>
    </body>
    </html>
    """
    temp_html_path = os.path.abspath("temp_drilldown_replica.html")
    with open(temp_html_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    
    driver.switch_to.window(print_tab_handle)
    driver.get(f"file:///{temp_html_path.replace(chr(92), '/')}")
    
    try:
        WebDriverWait(driver, 3).until(lambda d: "Sahyog Portal" in d.page_source)
        result = driver.execute_cdp_cmd("Page.printToPDF", {'landscape': False, 'printBackground': True, 'marginTop': 0.2, 'marginBottom': 0.2})
        if result and 'data' in result:
            with open(output_path, "wb") as f:
                f.write(base64.b64decode(result['data']))
    except Exception as e:
        logging.warning(f"Local Replica PDF generation failed: {e}")
        
    driver.get("about:blank")
    driver.switch_to.window(main_tab_handle)
    if os.path.exists(temp_html_path):
        try: os.remove(temp_html_path)
        except: pass

# ==========================================
# DYNAMIC BATCH ZIPPER
# ==========================================
def zip_and_send_pdfs(pdf_list, target_output_dir, timestamp):
    if not pdf_list: return
    MAX_SIZE_BYTES = 45 * 1024 * 1024  
    current_size = 0
    batch_pdfs = []
    part_num = 1
    
    send_telegram_message(f"🗜️ Compressing {len(pdf_list)} PDFs for delivery...")

    def create_and_send_zip(files, part):
        zip_name = os.path.join(target_output_dir, f"Sahyog_PDFs_{timestamp}_Part{part}.zip")
        with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file in files:
                zipf.write(file, os.path.basename(file))
        send_telegram_document(zip_name)
        logging.info(f"Sent ZIP Part {part} ({len(files)} files)")

    for pdf in pdf_list:
        if not os.path.exists(pdf): continue
        file_size = os.path.getsize(pdf)
        if current_size + file_size > MAX_SIZE_BYTES:
            create_and_send_zip(batch_pdfs, part_num)
            part_num += 1
            batch_pdfs = []
            current_size = 0
        batch_pdfs.append(pdf)
        current_size += file_size
        
    if batch_pdfs:
        create_and_send_zip(batch_pdfs, part_num)

# ==========================================
# 5. POST-PROCESSING ENGINE
# ==========================================
def save_safe_df(data_list, filename, target_output_dir):
    if not data_list: return None
    df = pd.DataFrame(data_list)
    for col in FINAL_EXCEL_COLUMNS:
        if col not in df.columns: df[col] = ""
    filepath = os.path.join(target_output_dir, filename)
    df[FINAL_EXCEL_COLUMNS].to_excel(filepath, index=False)
    return filepath

def run_post_processing(master_data, target_output_dir, timestamp, generated_pdfs):
    try:
        msg = "⚙️ Commencing Data Split & Final Group Uploads..."
        logging.info(msg)
        send_telegram_message(msg)
        
        west_data, east_data, other_data, urgent_data = [], [], [], []

        for row in master_data:
            block = str(row.get("Grievance Block", "")).replace("nan", "").strip()
            try: days_pending = int(float(row.get("Pending Duration (Level 1)", 0)))
            except: days_pending = 0
                
            if days_pending >= CONFIG["urgent_threshold_days"]:
                urgent_data.append(row)

            division = BLOCK_DIVISION_MAP.get(block, "OTHER")
            if division == "CHAPRA (WEST)": west_data.append(row)
            elif division == "CHAPRA (EAST)": east_data.append(row)
            else: other_data.append(row)

        excel_files = []
        if master_data: excel_files.append(save_safe_df(master_data, f"Sahyog_Data_Combined_{timestamp}.xlsx", target_output_dir))
        if west_data: excel_files.append(save_safe_df(west_data, f"Chapra West_{timestamp}.xlsx", target_output_dir))
        if east_data: excel_files.append(save_safe_df(east_data, f"Chapra East_{timestamp}.xlsx", target_output_dir))
        if other_data: excel_files.append(save_safe_df(other_data, f"Other Area_{timestamp}.xlsx", target_output_dir))
        if urgent_data: excel_files.append(save_safe_df(urgent_data, f"URGENT_COMPLIANCE_{timestamp}.xlsx", target_output_dir))
        
        excel_files = [f for f in excel_files if f is not None]
        
        send_telegram_message("📁 Uploading Final Excels...")
        for f in excel_files: send_telegram_document(f)
        
        zip_and_send_pdfs(generated_pdfs, target_output_dir, timestamp)
        logging.info("   [POST-PROCESSING] Files Generated and Sent Successfully.")
    except Exception as e:
        logging.error(f" [CRITICAL ERROR] Post-Processing Failed: {e}")
        logging.error(traceback.format_exc())

# ==========================================
# 6. SINGLE-PASS PRODUCTION ENGINE
# ==========================================
def main():
    toggle_webhook(False)
    
    try:
        logging.info("\n" + "="*80)
        logging.info("   SAHYOG V4.13 - GROUP CHAT PRODUCTION & LOOP PROTECTION ACTIVE")
        logging.info(f"   Detailed Logs: {log_filename}")
        logging.info("="*80)
        send_telegram_message("🚀 Sahyog Cloud Engine Starting (Drilldown Group Mode)...")

        extracted_ack_set = set()
        previous_df = pd.DataFrame()
        
        # HISTORICAL SOFT SYNC (TELEGRAM GROUP CHAT FEATURE)
        msg = "Do you want to load a historical Excel file for Soft Sync? Send the .xlsx file now, or reply 'skip'."
        resp = wait_for_telegram_reply_or_file(msg)
        if resp['type'] == 'file' and resp['name'].endswith(('.xls', '.xlsx')):
            with open("temp_history.xlsx", "wb") as f: f.write(resp['content'])
            try:
                previous_df = pd.read_excel("temp_history.xlsx").fillna("")
                if "Registration No." in previous_df.columns:
                    extracted_ack_set = set(previous_df["Registration No."].dropna().astype(str).str.strip())
                    send_telegram_message(f"✅ Successfully loaded {len(extracted_ack_set)} reference numbers to Skip List.")
                elif "Ack No" in previous_df.columns:
                    extracted_ack_set = set(previous_df["Ack No"].dropna().astype(str).str.strip())
                    send_telegram_message(f"✅ Successfully loaded {len(extracted_ack_set)} legacy reference numbers to Skip List.")
                else:
                    send_telegram_message("⚠️ Reference column not found in Excel. Skip list empty.")
            except Exception as e:
                send_telegram_message(f"❌ Failed to parse Excel: {e}")

        timestamp = datetime.now().strftime('%d-%m-%Y_%H%M')
        target_output_dir = os.path.join(os.getcwd(), f"Sahyog_Drilldown_{timestamp}")
        os.makedirs(target_output_dir, exist_ok=True)

        master_data = []
        generated_pdfs = []
        block_status = {} 

        # HEADLESS CHROME ARCHITECTURE WITH SSL BYPASS
        chrome_options = Options()
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-popup-blocking") 
        chrome_options.add_argument("--ignore-certificate-errors")
        chrome_options.set_capability("acceptInsecureCerts", True)
        chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        chrome_options.add_argument("--window-size=1920,1080")
        
        driver = webdriver.Chrome(options=chrome_options)
        driver.set_page_load_timeout(45)

        main_tab = driver.current_window_handle
        
        # Execute Login Once (No dynamic prompts)
        perform_login(driver, SAHYOG_USER, SAHYOG_PASS)

        driver.execute_script("window.open('about:blank', 'print_tab');")
        WebDriverWait(driver, 5).until(lambda d: len(d.window_handles) > 1)
        print_tab = [h for h in driver.window_handles if h != main_tab][0]

        # LOOP PROTECTION SAFETY ENGINE
        recovery_attempts = 0
        MAX_RECOVERY_ATTEMPTS = 5

        # ================= THE SOFT RECOVERY LOOP =================
        while True: 
            logging.info("\n 🔄 Initiating Drilldown Routine...")
            drilldown_url = "https://sahyog.bihar.gov.in/Sahyog/IGRS_InnerPage/Reports/DepartmentWiseConsolidatedReport.aspx"
            driver.get(drilldown_url)
            wait_for_ajax(driver)
            time.sleep(3) 

            # 1. Department Selection
            dept_clicked = False
            for trial in range(3):
                if not wait_for_table(driver, "ctl00_ContentPlaceHolder1_gvDepartmentReport", timeout=10): 
                    break
                
                dept_table = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_gvDepartmentReport")
                found = False
                for row in dept_table.find_elements(By.TAG_NAME, "tr")[1:]:
                    cols = row.find_elements(By.TAG_NAME, "td")
                    if len(cols) > 1 and CONFIG["target_department"].lower() in cols[1].text.lower():
                        safe_click(driver, cols[1].find_element(By.TAG_NAME, "a"))
                        time.sleep(1) 
                        wait_for_ajax(driver)
                        found = True
                        break
                
                if found:
                    if wait_for_table(driver, "ctl00_ContentPlaceHolder1_gvDistrictWise", timeout=5):
                        dept_clicked = True
                        break
                    else:
                        logging.warning(f"⚠️ 'Energy' click failed to load next page (Trial {trial + 1}/3). Retrying click...")
                else:
                    logging.warning("⚠️ 'Energy' department row not found. May have 0 pending complaints.")
                    break

            if not dept_clicked:
                logging.warning("⚠️ Failed to verify Department click. Reloading Soft Loop...")
                continue 

            # 2. District Selection
            dist_clicked = False
            for trial in range(3):
                if not wait_for_table(driver, "ctl00_ContentPlaceHolder1_gvDistrictWise", timeout=10): 
                    break
                
                dist_table = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_gvDistrictWise")
                found = False
                for row in dist_table.find_elements(By.TAG_NAME, "tr")[1:-1]:
                    cols = row.find_elements(By.TAG_NAME, "td")
                    if len(cols) > 1 and CONFIG["target_district"].lower() in cols[1].text.lower():
                        safe_click(driver, cols[1].find_element(By.TAG_NAME, "a"))
                        time.sleep(1) 
                        wait_for_ajax(driver)
                        found = True
                        break
                        
                if found:
                    if wait_for_table(driver, "ctl00_ContentPlaceHolder1_gvBlock", timeout=5):
                        dist_clicked = True
                        break
                    else:
                        logging.warning(f"⚠️ District click failed to load next page (Trial {trial + 1}/3). Retrying click...")
                else:
                    logging.warning("⚠️ Target District row not found.")
                    break
            
            if not dist_clicked:
                logging.warning("⚠️ Failed to verify District click. Reloading Soft Loop...")
                continue

            # Block Count
            if not wait_for_table(driver, "ctl00_ContentPlaceHolder1_gvBlock"): continue
            try:
                block_rows_count = len(driver.find_elements(By.XPATH, "//table[@id='ctl00_ContentPlaceHolder1_gvBlock']//tr")) - 2
            except Exception:
                logging.warning("⚠️ Failed to count block rows. Retrying Soft Loop...")
                continue

            block_loop_crashed = False

            # Block Loop
            for i in range(1, block_rows_count + 1):
                driver.switch_to.window(main_tab)
                
                try:
                    row = driver.find_elements(By.XPATH, "//table[@id='ctl00_ContentPlaceHolder1_gvBlock']//tr")[i]
                    cols = row.find_elements(By.TAG_NAME, "td")
                    block_name = cols[1].text.strip()
                    pending_count = int(cols[3].text.strip()) if cols[3].text.strip().isdigit() else 0
                except IndexError:
                    logging.warning(f"⚠️ Server sent a broken Block List page (Missing row index {i}). Triggering Soft Recovery...")
                    block_loop_crashed = True
                    break
                except Exception as e:
                    logging.warning(f"⚠️ Unexpected error parsing Block List row {i}: {e}. Triggering Soft Recovery...")
                    block_loop_crashed = True
                    break
                
                if block_name not in block_status:
                    block_status[block_name] = {'expected': pending_count, 'success': False, 'handled': 0}
                else:
                    block_status[block_name]['expected'] = pending_count 
                
                if pending_count == 0 or block_status[block_name]['success']:
                    logging.debug(f"Skipping {block_name} - Already Complete or 0 Pending.")
                    continue
                    
                block_start_msg = f"📍 [BLOCK START] Entering {block_name} (Expected: {pending_count})"
                logging.info(block_start_msg)
                send_telegram_message(block_start_msg)
                
                safe_click(driver, cols[3].find_element(By.TAG_NAME, "a"))
                time.sleep(1) 
                wait_for_ajax(driver)
                
                if not wait_for_table(driver, "ctl00_ContentPlaceHolder1_gvDetails"):
                    block_loop_crashed = True
                    break

                detail_rows_count = len(driver.find_elements(By.XPATH, "//table[@id='ctl00_ContentPlaceHolder1_gvDetails']//tr[position()>1]"))
                handled_this_block = 0
                
                # Detail Grid Loop
                for row_idx in range(detail_rows_count):
                    driver.switch_to.window(main_tab)
                    check_session(driver, SAHYOG_USER, SAHYOG_PASS)
                    verify_page_state(driver)
                    try:
                        d_row = driver.find_elements(By.XPATH, "//table[@id='ctl00_ContentPlaceHolder1_gvDetails']//tr[position()>1]")[row_idx]
                        d_cols = d_row.find_elements(By.TAG_NAME, "td")
                        if len(d_cols) < 23: continue
                        
                        ack_no = d_cols[1].text.strip()
                        
                        g_block = d_cols[14].text.strip()
                        g_panch = d_cols[15].text.strip()
                        subdiv, section = get_subdivision_section(g_block, g_panch)

                        row_data = {
                            "S.No.": d_cols[0].text, "Registration No.": ack_no, "Application Type": d_cols[2].text,
                            "Applicant Name": d_cols[3].text, "Mobile No": d_cols[4].text, "Department Name": d_cols[5].text,
                            "Complaint Status": d_cols[6].text, "Applicant District": d_cols[7].text, "Applicant Block": d_cols[8].text,
                            "Applicant Panchayat": d_cols[9].text, "Applicant Police Station": d_cols[10].text, 
                            "Grievance Division": d_cols[11].text, "Grievance District": d_cols[12].text, 
                            "Grievance Sub Division": d_cols[13].text, "Grievance Block": g_block, 
                            "Grievance Panchayat": g_panch, "Grievance Police Station": d_cols[16].text, 
                            "Grievance Type": d_cols[17].text, "Designation": d_cols[18].text, "Designation Level": d_cols[19].text,
                            "Delegated Status": d_cols[20].text, "Pending Duration (Level 1)": d_cols[21].text, 
                            "Delegation Duration (Days)": d_cols[22].text,
                            "Subdivision": subdiv, "Section": section
                        }
                        
                        # IN-FLIGHT SMART MERGE
                        if ack_no in extracted_ack_set:
                            old_row = None
                            if not previous_df.empty:
                                ref_col = "Registration No." if "Registration No." in previous_df.columns else "Ack No"
                                match = previous_df[previous_df[ref_col].astype(str).str.strip() == ack_no]
                                if not match.empty:
                                    old_row = match.iloc[0].to_dict()
                            
                            if old_row:
                                merged_data = row_data.copy()
                                deep_keys = ["Registration Date", "Father's Name", "Email", "Full Address", "Pincode", 
                                             "Grievance Description", "Delegated To Officer", "Delegated Action Status",
                                             "Delegated Action Date", "Delegated Remarks", "History Officer Details",
                                             "History Action Date", "History Feedback", "History Remarks"]
                                for key in deep_keys:
                                    merged_data[key] = clean_for_excel(old_row.get(key, "N/A"))
                                
                                merged_data["Last Updated Time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                master_data.append(merged_data)
                                handled_this_block += 1
                                continue 
                        
                        ack_url = d_cols[1].find_element(By.TAG_NAME, "a").get_attribute("href")
                        driver.execute_script(f"window.open('{ack_url}', '_blank');")
                        WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) >= 3)
                        detail_tab = [h for h in driver.window_handles if h not in [main_tab, print_tab]][0]
                        driver.switch_to.window(detail_tab)
                        
                        desc, del_off, del_stat, del_date, del_rem = "N/A", "N/A", "N/A", "N/A", "N/A"
                        hist_off, hist_date, hist_feed, hist_rem = "N/A", "N/A", "N/A", "N/A"
                        app_date, father_name, email, full_address, pincode = "N/A", "N/A", "N/A", "N/A", "N/A"
                        has_sec_pdf = False
                        
                        temp_pdf1 = os.path.join(target_output_dir, f"temp1_{ack_no}.pdf")
                        temp_pdf2 = os.path.join(target_output_dir, f"temp2_{ack_no}.pdf")

                        try:
                            wait_for_table(driver, "ctl00_ContentPlaceHolder1_gvpreview", 5)
                            preview_table = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_gvpreview")
                            data_row = preview_table.find_element(By.XPATH, ".//tr[last()]")
                            tds = data_row.find_elements(By.TAG_NAME, "td")
                            
                            applicant_dict = parse_cell_data(tds[1].text)
                            father_name = applicant_dict.get("Father/Husband Name", "N/A")
                            
                            address_dict = parse_cell_data(tds[2].text)
                            full_address = address_dict.get("Address", "N/A")
                            pincode = address_dict.get("PinCode", "N/A")
                            
                            app_dict = parse_cell_data(tds[3].text)
                            app_date = app_dict.get("Date", "N/A")
                            
                            desc = tds[5].text.strip()
                        except Exception as e: logging.debug(f"{ack_no}: Failed to parse preview: {e}")

                        try:
                            del_table = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_grdDelegate")
                            del_off = del_table.find_element(By.XPATH, ".//tr[2]/td[2]").text.strip()
                            del_stat = del_table.find_element(By.XPATH, ".//tr[2]/td[3]").text.strip()
                            del_date = del_table.find_element(By.XPATH, ".//tr[2]/td[4]").text.strip()
                            del_rem = del_table.find_element(By.XPATH, ".//tr[2]/td[5]").text.strip()
                        except: pass

                        try:
                            hist_table = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_Gvforwarding")
                            hist_row = hist_table.find_element(By.XPATH, ".//tbody/tr[1]")
                            hist_off = hist_row.find_element(By.XPATH, "./td[2]").text.strip()
                            hist_date = hist_row.find_element(By.XPATH, "./td[3]").text.strip()
                            hist_feed = hist_row.find_element(By.XPATH, "./td[4]").text.strip()
                            hist_rem = hist_row.find_element(By.XPATH, "./td[5]").text.strip()
                        except: pass

                        row_data.update({
                            "Father's Name": father_name, "Email": email, "Full Address": full_address, "Pincode": pincode,
                            "Registration Date": app_date, "Grievance Description": desc, 
                            "Delegated To Officer": del_off, "Delegated Action Status": del_stat,
                            "Delegated Action Date": del_date, "Delegated Remarks": del_rem, 
                            "History Officer Details": hist_off, "History Action Date": hist_date, 
                            "History Feedback": hist_feed, "History Remarks": hist_rem,
                            "Last Updated Time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        })
                        for k in row_data: row_data[k] = clean_for_excel(row_data[k])

                        logging.info(f" 🖨️  Generating Executive Blue PDF for {ack_no}...")
                        generate_replica_pdf(driver, print_tab, detail_tab, row_data, temp_pdf1)

                        logging.info(f" 📎 Fetching Attached Document for {ack_no}...")
                        try:
                            match = re.search(r'(JVBER[A-Za-z0-9+/=\s]{100,})', driver.page_source)
                            
                            if not match:
                                pdf_links = driver.find_elements(By.XPATH, "//a[contains(@href, 'IDoc.aspx')]")
                                if pdf_links:
                                    safe_click(driver, pdf_links[0])
                                    WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 3)
                                    attach_tab = [h for h in driver.window_handles if h not in [main_tab, print_tab, detail_tab]][0]
                                    driver.switch_to.window(attach_tab)
                                    WebDriverWait(driver, 15).until(lambda d: "JVBER" in d.page_source)
                                    match = re.search(r'(JVBER[A-Za-z0-9+/=\s]{100,})', driver.page_source)
                                    
                            if match:
                                pdf_bytes = base64.b64decode(match.group(1).replace('\n', '').replace('\r', '').replace(' ', ''))
                                if pdf_bytes.startswith(b'%PDF'):
                                    with open(temp_pdf2, "wb") as f: f.write(pdf_bytes)
                                    has_sec_pdf = True
                                    
                        except Exception as e: 
                            logging.error(f"Failed to extract attached PDF for {ack_no}: {e}")
                        finally:
                            close_extra_tabs(driver, [main_tab, print_tab, detail_tab])
                            driver.switch_to.window(detail_tab)

                        base_filename = f"{sanitize_filename(ack_no)}_{sanitize_filename(row_data['Applicant Name'])}_{sanitize_filename(row_data['Applicant Block'])}"
                        final_pdf = os.path.join(target_output_dir, f"{base_filename}.pdf")

                        merger = PdfWriter()
                        if os.path.exists(temp_pdf1) and is_valid_pdf(temp_pdf1): merger.append(temp_pdf1)
                        if has_sec_pdf and is_valid_pdf(temp_pdf2): merger.append(temp_pdf2)
                        if os.path.exists(temp_pdf1) or has_sec_pdf: merger.write(final_pdf)
                        merger.close()

                        if os.path.exists(temp_pdf1): os.remove(temp_pdf1)
                        if os.path.exists(temp_pdf2): os.remove(temp_pdf2)

                        master_data.append(row_data)
                        extracted_ack_set.add(ack_no)
                        handled_this_block += 1
                        generated_pdfs.append(final_pdf)

                        if len(master_data) % CONFIG["excel_save_batch_size"] == 0:
                            save_safe_df(master_data, f"Sahyog_Data_TMP_{timestamp}.xlsx", target_output_dir)

                    except Exception as row_e:
                        logging.error(f"Failed to process row {ack_no}: {row_e}")
                    finally:
                        close_extra_tabs(driver, [main_tab, print_tab])
                        driver.switch_to.window(main_tab)
                
                block_status[block_name]['handled'] = handled_this_block
                if handled_this_block >= pending_count:
                    block_status[block_name]['success'] = True
                    block_end_msg = f" ✅ Block {block_name} 100% Complete."
                    logging.info(block_end_msg)
                    send_telegram_message(block_end_msg)
                else:
                    logging.warning(f" ⚠️ Block {block_name} incomplete. Scheduled for Soft Recovery.")

                safe_click(driver, driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnClose"))
                time.sleep(1) 
                wait_for_ajax(driver)
                if not wait_for_table(driver, "ctl00_ContentPlaceHolder1_gvBlock", timeout=15):
                    logging.warning("⚠️ Failed to load Block List after backtracking. Triggering Soft Recovery...")
                    block_loop_crashed = True
                    break

            if block_loop_crashed:
                pass

            # --- SOFT LOOP AUDIT & PROTECTION LOCK ---
            all_success = True
            total_expected, total_handled = 0, 0
            
            for b, st in block_status.items():
                total_expected += st['expected']
                total_handled += st.get('handled', 0)
                if st['expected'] > 0 and not st['success']: all_success = False

            audit_str = "\n".join([f"{block.ljust(15)} | Exp: {stats['expected']} | Ext: {stats.get('handled', 0)} | {'✅' if stats.get('success') else '❌'}" for block, stats in block_status.items() if stats['expected'] > 0])
            
            logging.info("\n" + "="*80)
            logging.info("                      CYCLE AUDIT REPORT")
            logging.info("="*80)
            logging.info("\n" + audit_str)
            
            if all_success:
                success_msg = f"✅ 100% INTEGRITY REACHED ({total_handled}/{total_expected}). BREAKING SOFT LOOP."
                logging.info(success_msg)
                send_telegram_message(success_msg)
                break 
            else:
                recovery_attempts += 1
                missing = total_expected - total_handled
                
                # Hard Cap protection to stop infinite loop crashes
                if recovery_attempts >= MAX_RECOVERY_ATTEMPTS:
                    fail_msg = f"🚨 INFINITE LOOP TRIGGERED: Recovery failed {MAX_RECOVERY_ATTEMPTS} times due to portal lag. Force breaking to protect server. Handing over saved partial dataset ({total_handled}/{total_expected})."
                    logging.critical(fail_msg)
                    send_telegram_message(fail_msg)
                    break
                    
                fail_msg = f"⚠️ AUDIT FAILED (Attempt {recovery_attempts}/{MAX_RECOVERY_ATTEMPTS}): Expected {total_expected}, but only handled {total_handled}. Missing {missing}.\n🔄 Executing SMART SOFT-RECOVERY..."
                logging.warning(fail_msg)
                send_telegram_message(fail_msg)
                time.sleep(3)
                continue 
        
        # End loop directly - Clean Single-pass run execution
        break

    # --- FINAL DATA AUDIT BROADCAST ---
    final_audit_msg = (
        f"📊 *Final Sahyog Sync Audit Report*\n\n"
        f"📅 Date: {datetime.now().strftime('%d-%m-%Y')}\n"
        f"🎯 Target District: {CONFIG['target_district']}\n\n"
        f"📥 *Records Summary:*\n"
        f"  • Total Expected: {total_expected}\n"
        f"  • Total Downloaded: {total_handled}\n"
    )
    if total_expected == total_handled:
        final_audit_msg += "\n✅ *Status:* 100% Synced successfully."
    else:
        final_audit_msg += f"\n⚠️ *Status:* Partial Run. {total_expected - total_handled} records missed due to server lag."
    
    send_telegram_message(final_audit_msg)

    # ==========================================
    # 7. EXECUTIVE POST-PROCESSING
    # ==========================================
    if master_data:
        run_post_processing(master_data, target_output_dir, timestamp, generated_pdfs)
        
    logging.info(f"Process Complete. Detailed logs saved to: {log_filename}")
    send_telegram_document(log_filename)

# 🌍 GLOBAL GROUP EXCEPTION INTERCEPTOR 🌍
except Exception as e:
    error_msg = f"💥 CRITICAL BOT FAILURE IN PRODUCTION:\n{str(e)}\n\nTraceback:\n{traceback.format_exc()[-1000:]}"
    logging.critical(error_msg)
    try: send_telegram_message(error_msg)
    except: pass

finally:
    try: driver.quit()
    except: pass
    toggle_webhook(True)

if __name__ == "__main__":
    main()
