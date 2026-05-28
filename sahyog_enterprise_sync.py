import os
import sys
import time
import re
import json
import logging
import base64
import zipfile
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
# WEBHOOK TOGGLE ENGINE
# ==========================================
def toggle_webhook(enable=True):
    if enable and WEBHOOK_URL:
        url = f"https://api.telegram.org/bot{BOT_TOKEN}/setWebhook?url={WEBHOOK_URL}"
        try: requests.get(url)
        except: pass
        logging.info("🔗 Webhook re-enabled. Google is back in control.")
    else:
        url = f"https://api.telegram.org/bot{BOT_TOKEN}/deleteWebhook"
        try: requests.get(url)
        except: pass
        logging.info("🔌 Webhook disabled. Python is now listening via polling.")

# ==========================================
# TELEGRAM COMMUNICATIONS & SOFT SYNC
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
            requests.post(url, data=payload, files=files, timeout=120) 
            logging.info(f"📤 Uploaded {os.path.basename(file_path)} to Telegram.")
    except Exception as e: logging.error(f"Failed to send file: {e}")

def get_telegram_file(file_id):
    file_info = requests.get(f"https://api.telegram.org/bot{BOT_TOKEN}/getFile?file_id={file_id}").json()
    file_path = file_info['result']['file_path']
    download_url = f"https://api.telegram.org/file/bot{BOT_TOKEN}/{file_path}"
    return requests.get(download_url).content

def wait_for_telegram_reply_or_file(prompt_message):
    send_telegram_message(prompt_message)
    logging.info("⏳ Waiting for Telegram text or file...")
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/getUpdates"
    requests.get(url) # Flush old messages
    time.sleep(1)
    response = requests.get(url).json()
    last_update_id = response["result"][-1]["update_id"] if response.get("result") else 0

    while True:
        poll_url = f"{url}?offset={last_update_id + 1}&timeout=5"
        try:
            resp = requests.get(poll_url).json()
            if resp.get("result"):
                for update in resp["result"]:
                    last_update_id = update["update_id"]
                    if "message" in update:
                        if str(update["message"]["chat"]["id"]) == str(CHAT_ID):
                            if "document" in update["message"]:
                                doc = update["message"]["document"]
                                file_id, file_name = doc["file_id"], doc["file_name"]
                                send_telegram_message(f"📥 Received file: {file_name}. Processing...")
                                return {"type": "file", "content": get_telegram_file(file_id), "name": file_name}
                            elif "text" in update["message"]:
                                text = update["message"]["text"].strip()
                                send_telegram_message(f"✅ Received: {text}")
                                return {"type": "text", "text": text}
        except: pass
        time.sleep(2)

# ==========================================
# UTILITY & PDF ENGINE
# ==========================================
def sanitize_filename(name):
    if not name: return "Unknown"
    return re.sub(r'[\\/*?:"<>|]', "", str(name)).strip().replace(" ", "_")

def save_page_as_pdf(driver, output_path):
    print_options = {'landscape': False, 'displayHeaderFooter': False, 'printBackground': True, 'preferCSSPageSize': True}
    result = driver.execute_cdp_cmd("Page.printToPDF", print_options)
    with open(output_path, "wb") as f:
        f.write(base64.b64decode(result['data']))

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

def check_session(driver):
    try:
        driver.find_element(By.ID, "ctl00_lblwelcome")
        return True
    except:
        msg = "🚨 SESSION EXPIRED! Triggering re-login flow..."
        logging.warning(msg)
        send_telegram_message(msg)
        perform_login(driver)
        return True

def load_overrides():
    overrides = {}
    if os.path.exists(OVERRIDE_FILE):
        try:
            df = pd.read_excel(OVERRIDE_FILE).fillna("")
            if all(col in df.columns for col in ['Registration No.', 'Subdivision', 'Section']):
                for _, row in df.iterrows():
                    ref = str(row['Registration No.']).strip()
                    if ref:
                        overrides[ref] = {'Subdivision': str(row['Subdivision']).strip(), 'Section': str(row['Section']).strip()}
        except Exception as e: logging.error(f"Failed to read overrides: {e}")
    return overrides

# ==========================================
# DYNAMIC BATCH ZIPPER (45 MB THRESHOLD)
# ==========================================
def zip_and_send_pdfs(pdf_list, target_output_dir, timestamp):
    if not pdf_list: return
    
    MAX_SIZE_BYTES = 45 * 1024 * 1024  # 45 MB absolute limit
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

def run_post_processing(master_logged_rows, target_output_dir, timestamp, generated_pdfs):
    msg = "⚙️ Commencing Data Split & Final Uploads..."
    logging.info(msg)
    send_telegram_message(msg)
    
    combined_data, west_data, east_data, other_data, urgent_data = [], [], [], [], []
    overrides = load_overrides()
    
    for row in master_logged_rows:
        raw_date = str(row.get("Registration Date", ""))
        clean_date = raw_date.split(" ")[0] if raw_date else ""
        block = str(row.get("Block Name", "")).replace("nan", "").strip()
        panchayat = str(row.get("Panchayat Name", "")).replace("nan", "").strip()
        ref_no = str(row.get("Registration No.", "")).strip()
        
        # APPLY MANUAL OVERRIDES ENGINE
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
        else: other_data.append(formatted_row) # "OTHER AREA" RESTORED

    final_cols = list(formatted_row.keys())
    excel_files = []

    if combined_data:
        path = os.path.join(target_output_dir, f"Sahyog_Shivir_{timestamp}_Combined.xlsx")
        if atomic_save_excel(pd.DataFrame(combined_data)[final_cols], path): excel_files.append(path)
    if west_data:
        path = os.path.join(target_output_dir, f"Chapra West_{timestamp}.xlsx")
        if atomic_save_excel(pd.DataFrame(west_data)[final_cols], path): excel_files.append(path)
    if east_data:
        path = os.path.join(target_output_dir, f"Chapra East_{timestamp}.xlsx")
        if atomic_save_excel(pd.DataFrame(east_data)[final_cols], path): excel_files.append(path)
    if other_data:
        path = os.path.join(target_output_dir, f"Other Area_{timestamp}.xlsx")
        if atomic_save_excel(pd.DataFrame(other_data)[final_cols], path): excel_files.append(path)
    if urgent_data:
        path = os.path.join(target_output_dir, f"URGENT_COMPLIANCE_{timestamp}.xlsx")
        if atomic_save_excel(pd.DataFrame(urgent_data)[final_cols], path): excel_files.append(path)

    send_telegram_message("📁 Uploading Final Excels...")
    for f in excel_files: send_telegram_document(f)
    
    zip_and_send_pdfs(generated_pdfs, target_output_dir, timestamp)

# ==========================================
# AUTH & SCRAPING ENGINE
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
            
            resp = wait_for_telegram_reply_or_file("Enter Captcha:")
            captcha_text = resp['text'] if resp['type'] == 'text' else ""
            
            driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtUserName").send_keys(SAHYOG_USER)
            driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtPassword").send_keys(SAHYOG_PASS)
            driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtCode").send_keys(captcha_text)
            driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnLogin").click()
            time.sleep(5)
            
            if len(driver.find_elements(By.XPATH, "//a[contains(text(), 'Grievance Action')]")) > 0:
                logging.info("✅ Login Successful!")
                break
        except Exception as e: time.sleep(5)

def get_dropdown_options(driver, element_id):
    try:
        select_elem = Select(driver.find_element(By.ID, element_id))
        return [{"value": opt.get_attribute("value"), "text": opt.text.strip()} 
                for opt in select_elem.options if opt.get_attribute("value") and opt.get_attribute("value") != "0" and "--" not in opt.text]
    except: return []

# --- EXTRACTED DROPDOWN PROCESSOR WITH AUTO-RETRY ---
def process_dropdown_combo(driver, main_tab, target_output_dir, master_logged_rows, audit_log, generated_pdfs, t_opt, c_opt, s_opt, radio_index=None, radio_name=""):
    combo_name = f"{t_opt['text']} > {c_opt['text']} > {s_opt['text']}"
    if radio_name: combo_name += f" > {radio_name}"
    
    max_retries = 3
    for attempt in range(max_retries):
        try:
            driver.switch_to.window(main_tab)
            logging.info(f"🔄 Scanning Filter: {combo_name} (Attempt {attempt+1}/{max_retries})")

            # EXPLICIT UI BUFFERS AND STALENESS CHECKS
            old_status = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlFilterStatus")
            Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlFilterType")).select_by_value(t_opt['value'])
            time.sleep(1) 
            try: WebDriverWait(driver, 10).until(EC.staleness_of(old_status))
            except: pass
            time.sleep(1) 

            old_status = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlFilterStatus")
            Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlFilterComplaint")).select_by_value(c_opt['value'])
            time.sleep(1)
            try: WebDriverWait(driver, 10).until(EC.staleness_of(old_status))
            except: pass
            time.sleep(1)
            
            Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlFilterStatus")).select_by_value(s_opt['value'])
            time.sleep(2) 
            
            # DELEGATED RADIO BUTTON ENGINE RESTORED
            if radio_index is not None:
                radios = driver.find_elements(By.CSS_SELECTOR, ".delegated-panel .delegated-radio label")
                if radio_index < len(radios):
                    safe_click(driver, radios[radio_index])
                    time.sleep(2)
            
            # FORCE TOP RECORDS TO --ALL-- RESTORED
            try: 
                Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlTopFilter")).select_by_value("0")
            except: pass
            time.sleep(3) 

            metrics = {'combo': combo_name, 'total': 0, 'skipped': 0, 'extracted': 0, 'errors': 0}
            safe_click(driver, driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnViewNormal"))
            time.sleep(4)
            verify_page_state(driver)
            
            # Run the extraction and check if math balances
            success = run_extraction_sequence(driver, main_tab, target_output_dir, master_logged_rows, t_opt, c_opt, s_opt, metrics, generated_pdfs)
            
            if success:
                audit_log.append(metrics)
                break 
            else:
                if attempt < max_retries - 1:
                    backoff_time = 5 * (2 ** attempt) 
                    msg = f"⚠️ WARNING: Data Mismatch in {combo_name}. Missing records! Redoing filter in {backoff_time}s..."
                    logging.warning(msg)
                    send_telegram_message(msg)
                    time.sleep(backoff_time) 
                    driver.refresh()
                    time.sleep(5)
                else:
                    logging.error(f"❌ Failed to reconcile {combo_name} after {max_retries} attempts. Moving on to protect server.")
                    audit_log.append(metrics)
                    
        except Exception as e: 
            logging.error(f"[ERROR] Failed loop: {e}")
            driver.refresh()
            time.sleep(5)


def perform_scraping_cycle(driver, main_tab, target_output_dir, master_logged_rows, audit_log, generated_pdfs):
    WebDriverWait(driver, 30).until(lambda d: len(d.find_elements(By.ID, "ctl00_ContentPlaceHolder1_ddlFilterType")) > 0)
    type_options = get_dropdown_options(driver, "ctl00_ContentPlaceHolder1_ddlFilterType")
    complaint_options = get_dropdown_options(driver, "ctl00_ContentPlaceHolder1_ddlFilterComplaint")
    status_options = get_dropdown_options(driver, "ctl00_ContentPlaceHolder1_ddlFilterStatus")

    for t_opt in type_options:
        for c_opt in complaint_options:
            for s_opt in status_options:
                # Dry run to check for Delegated Radios
                driver.switch_to.window(main_tab)
                try: Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlFilterType")).select_by_value(t_opt['value'])
                except: pass
                time.sleep(1)
                try: Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlFilterComplaint")).select_by_value(c_opt['value'])
                except: pass
                time.sleep(1)
                try: Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlFilterStatus")).select_by_value(s_opt['value'])
                except: pass
                time.sleep(2)
                
                radios = driver.find_elements(By.CSS_SELECTOR, ".delegated-panel .delegated-radio label")
                radio_names = [r.text.strip() for r in radios] if radios else []
                
                if radio_names:
                    for i, r_name in enumerate(radio_names):
                        process_dropdown_combo(driver, main_tab, target_output_dir, master_logged_rows, audit_log, generated_pdfs, t_opt, c_opt, s_opt, radio_index=i, radio_name=r_name)
                else:
                    process_dropdown_combo(driver, main_tab, target_output_dir, master_logged_rows, audit_log, generated_pdfs, t_opt, c_opt, s_opt)

def run_extraction_sequence(driver, main_tab, target_output_dir, master_logged_rows, t_opt, c_opt, s_opt, metrics, generated_pdfs):
    # 1. Establish the True Expected Total from the Blue Badge
    expected_total = 0
    try:
        total_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Total :') or contains(text(), 'Total:')]")
        for el in total_elements:
            match = re.search(r'Total\s*:\s*(\d+)', el.text, re.IGNORECASE)
            if match:
                expected_total = int(match.group(1))
                break
    except: pass

    # Fallback if badge fails
    if expected_total == 0:
        expected_total = len(driver.find_elements(By.CSS_SELECTOR, ".list-container .complaint-card"))

    metrics['total'] = expected_total
    
    if expected_total == 0: 
        logging.info(f"🔍 [FOUND] 0 records in '{metrics['combo']}'. Skipping.")
        return True # Nothing to process, success!

    # VERBOSE TELEMETRY: SEARCH STARTED
    start_msg = f"🔍 [FOUND] {expected_total} records in '{metrics['combo']}'. Starting extraction..."
    logging.info(start_msg)
    send_telegram_message(start_msg)

    # 2. Force Lazy-Loading by scrolling until all cards appear
    for _ in range(5): 
        cards = driver.find_elements(By.CSS_SELECTOR, ".list-container .complaint-card")
        if len(cards) >= expected_total: break
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)

    # Re-evaluate visible cards
    complaint_cards = driver.find_elements(By.CSS_SELECTOR, ".list-container .complaint-card")
    visible_cards = len(complaint_cards)
    
    # 3. Extraction Loop
    for card_idx in range(visible_cards):
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
                            metrics['skipped'] += 1
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
                    if row.get("Registration No.") == expected_ref: row.update(data); break
                metrics['skipped'] += 1
            else:
                temp_pdf1 = os.path.join(target_output_dir, f"temp_print_{card_idx}.pdf")
                temp_pdf2 = os.path.join(target_output_dir, f"temp_view_{card_idx}.pdf")
                
                hidden_pdf = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_hfPdfComp").get_attribute("value")
                safe_click(driver, driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_printing"))
                WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)
                driver.switch_to.window(driver.window_handles[-1])
                
                try:
                    WebDriverWait(driver, 10).until(lambda d: expected_ref in d.page_source)
                    reg_date = driver.find_element(By.XPATH, "//td[contains(text(), 'Reg.Date')]/following-sibling::td").text.strip()
                    g_type = driver.find_element(By.XPATH, "//td[contains(text(), 'Grievance Type')]/following-sibling::td").text.strip()
                except: reg_date, g_type = "N/A", "N/A"

                driver.execute_script("document.querySelectorAll('.btn, footer').forEach(e => e.style.display='none')")
                save_page_as_pdf(driver, temp_pdf1)
                driver.close()
                driver.switch_to.window(main_tab)

                has_sec = False
                if hidden_pdf and hidden_pdf.strip():
                    try:
                        resp = requests.get(urljoin("https://sahyog.bihar.gov.in/", hidden_pdf.strip()), cookies={c['name']: c['value'] for c in driver.get_cookies()}, timeout=15)
                        if resp.status_code == 200 and b"%PDF" in resp.content[:1024]:
                            with open(temp_pdf2, "wb") as f: f.write(resp.content)
                            has_sec = True
                    except: pass

                base_filename = f"{sanitize_filename(expected_ref)}_{sanitize_filename(data['Applicant Name'])}_{sanitize_filename(data['Block Name'])}_{sanitize_filename(data['Panchayat Name'])}"
                final_pdf = os.path.join(target_output_dir, f"{base_filename}.pdf")

                merger = PdfWriter()
                if os.path.exists(temp_pdf1): merger.append(temp_pdf1)
                if has_sec and os.path.exists(temp_pdf2): merger.append(temp_pdf2)
                merger.write(final_pdf)
                merger.close()

                if os.path.exists(temp_pdf1): os.remove(temp_pdf1)
                if os.path.exists(temp_pdf2): os.remove(temp_pdf2)

                data["Registration Date"] = reg_date
                data["Grievance Type"] = g_type
                master_logged_rows.append(data)
                
                # Save ledger immediately so if it crashes, data is perfectly safe
                with open(LEDGER_FILE, 'w', encoding='utf-8') as f: json.dump(master_logged_rows, f, ensure_ascii=False, indent=4)
                
                generated_pdfs.append(final_pdf)
                metrics['extracted'] += 1

            driver.switch_to.window(main_tab)
            safe_click(driver, driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnBack"))
            time.sleep(2)
        except Exception as e: 
            metrics['errors'] += 1
            try:
                driver.switch_to.window(main_tab)
                safe_click(driver, driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnBack"))
            except: pass

    # 4. The Accounting Gate: Do the numbers match?
    total_processed = metrics['extracted'] + metrics['skipped'] + metrics['errors']
    
    if total_processed < expected_total:
        return False # This tells the main loop to REDO this entire filter
        
    # VERBOSE TELEMETRY: CYCLE COMPLETED SUCCESSFULLY
    summary_msg = (
        f"✅ [COMPLETED] '{metrics['combo']}'\n"
        f"Total Expected: {expected_total}\n"
        f"📥 Extracted: {metrics['extracted']}\n"
        f"⏭️ Skipped: {metrics['skipped']}\n"
        f"❌ Errors: {metrics['errors']}"
    )
    logging.info(summary_msg)
    send_telegram_message(summary_msg)
    return True # Math matches perfectly, move to next!

def print_audit_report(audit_log):
    logging.info("\n" + "="*85)
    logging.info("FINAL EXTRACTION AUDIT REPORT")
    logging.info(f"{'FILTER COMBINATION'.ljust(40)} | {'TOTAL'.ljust(6)} | {'SKIPPED'.ljust(8)} | {'EXTRACTED'.ljust(10)} | {'ERRORS'.ljust(6)}")
    total_found = total_skipped = total_extracted = total_errors = 0
    for row in audit_log:
        logging.info(f"{row['combo'][:38].ljust(40)} | {str(row['total']).ljust(6)} | {str(row['skipped']).ljust(8)} | {str(row['extracted']).ljust(10)} | {str(row['errors']).ljust(6)}")
        total_found += row['total']; total_skipped += row['skipped']; total_extracted += row['extracted']; total_errors += row['errors']
    logging.info("="*85)
    logging.info(f"GRAND TOTALS: {total_found} Found | {total_skipped} Skipped | {total_extracted} New | {total_errors} Errors\n")

# ==========================================
# RUN MAIN ASSEMBLY
# ==========================================
def main():
    toggle_webhook(False)
    
    try:
        send_telegram_message("🚀 Sahyog Cloud Engine Starting...")
        
        master_logged_rows = []
        is_resuming = False
        
        if os.path.exists(LEDGER_FILE):
            msg = "🚨 INCOMPLETE RUN DETECTED. Reply 'yes' to resume, or 'no' for fresh run."
            resp = wait_for_telegram_reply_or_file(msg)
            if resp['type'] == 'text' and resp['text'].lower() in ['y', 'yes']:
                with open(LEDGER_FILE, 'r', encoding='utf-8') as f: master_logged_rows = json.load(f)
                is_resuming = True
                send_telegram_message(f"✅ Resuming... Loaded {len(master_logged_rows)} records from memory.")
            else:
                os.remove(LEDGER_FILE)
                send_telegram_message("🗑️ Old ledger deleted. Starting a fresh run.")
        
        # HISTORICAL SOFT SYNC (TELEGRAM FEATURE RESTORED)
        if not is_resuming:
            msg = "Do you want to load a historical Excel file for Soft Sync? Send the .xlsx file now, or reply 'skip'."
            resp = wait_for_telegram_reply_or_file(msg)
            if resp['type'] == 'file' and resp['name'].endswith(('.xls', '.xlsx')):
                with open("temp_history.xlsx", "wb") as f: f.write(resp['content'])
                try:
                    df = pd.read_excel("temp_history.xlsx").fillna("")
                    master_logged_rows = df.to_dict('records')
                    # Save to ledger instantly so history is preserved
                    with open(LEDGER_FILE, 'w', encoding='utf-8') as f: json.dump(master_logged_rows, f, ensure_ascii=False, indent=4)
                    skip_list = set(r.get("Registration No.") for r in master_logged_rows if r.get("Registration No."))
                    send_telegram_message(f"✅ Database loaded. Syncing {len(skip_list)} historical IDs.")
                except Exception as e:
                    send_telegram_message(f"❌ Failed to parse Excel: {e}")
        
        audit_log = []
        generated_pdfs = []
        
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
            
            perform_scraping_cycle(driver, driver.current_window_handle, target_output_dir, master_logged_rows, audit_log, generated_pdfs)
            
            resp = wait_for_telegram_reply_or_file("✅ Data collection complete. Loop another ID? (y/n):")
            cont = resp['text'].lower() if resp['type'] == 'text' else 'n'
            if cont not in ['y', 'yes']: break

        run_post_processing(master_logged_rows, target_output_dir, timestamp, generated_pdfs)
        print_audit_report(audit_log)
        
        if os.path.exists(LEDGER_FILE): os.remove(LEDGER_FILE)
        send_telegram_message("🎉 Automation Completed Successfully. Server shutting down.")
        driver.quit()

    finally:
        toggle_webhook(True)

if __name__ == "__main__":
    main()
