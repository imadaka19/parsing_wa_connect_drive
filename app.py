# app.py
import os
import io
import re
import json
import zipfile
import pickle
from difflib import get_close_matches
from datetime import datetime

import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

# ------------------------------------------------------------
# KONFIGURASI
# ------------------------------------------------------------
SCOPES = ["https://www.googleapis.com/auth/drive.file"]
DEFAULT_OUTPUT_FILENAME = "Stock_Opname_KNO_2025.xlsx"
FOLDER_ID = "1xBFx4yA7jV5OvevS1nG3HKON5glx1YUa"

st.set_page_config(page_title="Parser Stock Opname (WhatsApp) → Drive", layout="wide")
st.title("📦 Parser Stock Opname (WhatsApp) → Google Drive")

st.markdown(
    """
    **Petunjuk:**
    1️⃣ Upload file ZIP hasil export chat WhatsApp (format 24 jam).  
    2️⃣ Pastikan format chat seperti:
    ```
    LOC: K7
    BIN: RCM1
    PN: AR0006500
    SN: 12345678
    QTY EMRO: 46
    QTY ACTUAL: 47
    REMARK(S): SURPLUS 1
    ```
    3️⃣ Aplikasi akan parse data, upload foto ke Drive, dan buat file Excel.
    """
)

# ------------------------------------------------------------
# AUTENTIKASI GOOGLE DRIVE (pakai Streamlit Secrets)
# ------------------------------------------------------------
def get_credentials(uploaded_token):
    """Bangun kredensial dari secrets + token"""
    client_secret = {"installed": dict(st.secrets["google_oauth_client.installed"])}
    os.makedirs("temp", exist_ok=True)
    creds_path = "temp/client_secret.json"
    with open(creds_path, "w") as f:
        json.dump(client_secret, f)

    creds = None
    if uploaded_token is not None:
        with open("temp/token_uploaded.pickle", "wb") as f:
            f.write(uploaded_token.getbuffer())
        with open("temp/token_uploaded.pickle", "rb") as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            st.error(
                "Token tidak valid. Jalankan autentikasi di lokal untuk membuat `token.pickle`, "
                "lalu upload di sidebar."
            )
            st.stop()
    return creds


def build_drive_service(creds):
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def upload_to_drive(service, folder_id, img_path, entry):
    """Upload gambar ke Google Drive"""
    try:
        loc = (entry.get("LOC", "") or "").replace("/", "-").strip()
        bin_ = (entry.get("BIN", "") or "").replace("/", "-").strip()
        pn = (entry.get("PN", "") or "").replace("/", "-").strip()
        sn = (entry.get("SN", "") or "").replace("/", "-").strip() or "NO_SN"
        tanggal = (entry.get("Tanggal", "") or "").split(" ")[0].replace("/", "-").strip()

        parts = [p for p in [loc, bin_, pn, sn, tanggal] if p]
        filename = "-".join(parts) + os.path.splitext(img_path)[1]

        file_metadata = {"name": filename, "parents": [folder_id]}
        media = MediaFileUpload(img_path, mimetype="image/jpeg")

        uploaded = service.files().create(
            body=file_metadata, media_body=media, fields="id"
        ).execute()

        service.permissions().create(
            fileId=uploaded["id"], body={"type": "anyone", "role": "reader"}
        ).execute()

        return f"https://drive.google.com/uc?id={uploaded['id']}"
    except Exception as e:
        st.error(f"Upload error for {os.path.basename(img_path)}: {e}")
        return ""


# ------------------------------------------------------------
# FUNGSI PARSING CHAT
# ------------------------------------------------------------
def extract_zip_to_dir(zip_bytes_io, extract_dir):
    with zipfile.ZipFile(zip_bytes_io) as zip_ref:
        zip_ref.extractall(extract_dir)


def find_text_chat_file(extract_dir):
    for root, _, files in os.walk(extract_dir):
        for f in files:
            if f.lower().endswith(".txt"):
                return os.path.join(root, f)
    return None


def index_images(extract_dir):
    image_index = {}
    for root, _, files in os.walk(extract_dir):
        for f in files:
            if f.lower().endswith((".jpg", ".jpeg", ".png")):
                image_index[f.lower()] = os.path.join(root, f)
    return image_index


def find_image_fuzzy(name, image_index):
    if not name:
        return None
    name_low = name.lower()
    if name_low in image_index:
        return image_index[name_low]
    candidates = get_close_matches(name_low, image_index.keys(), n=1, cutoff=0.5)
    if candidates:
        return image_index[candidates[0]]
    return None


def parse_chat_file(chat_file_path):
    entries = []
    current = {}
    image_file = None
    current_sender = ""
    current_date = ""

    def save_current():
        if current.get("LOC"):
            entries.append({
                "Tanggal": current_date,
                "Pengirim": current_sender,
                "LOC": current.get("LOC", ""),
                "BIN": current.get("BIN", ""),
                "PN": current.get("PN", ""),
                "SN": current.get("SN", ""),
                "Qty EMRO": current.get("QTY EMRO", ""),
                "Qty ACTUAL": current.get("QTY ACTUAL", ""),
                "REMARK": current.get("REMARK", ""),
                "PHOTO FILE": image_file or "",
                "PHOTO LINK": ""
            })

    pattern_msg = re.compile(r"^(\d{1,2}/\d{1,2}/\d{2,4}), (\d{1,2}:\d{2}) - ([^:]+):")

    with open(chat_file_path, encoding="utf-8", errors="ignore") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            m = pattern_msg.match(line)
            if m:
                save_current()
                current = {}
                current_date = f"{m.group(1)} {m.group(2)}"
                current_sender = m.group(3).strip()
                match_img = re.search(
                    r'IMG-\d{8}-WA\d{4}.*?\.(?:jpg|jpeg|png)', line, re.IGNORECASE
                )
                image_file = match_img.group(0).strip() if match_img else None
                continue

            match_img = re.search(
                r'IMG-\d{8}-WA\d{4}.*?\.(?:jpg|jpeg|png)', line, re.IGNORECASE
            )
            if match_img:
                image_file = match_img.group(0).strip()
                continue

            if "<media omitted>" in line.lower():
                image_file = None
                continue

            if ":" in line:
                key, val = line.split(":", 1)
                key = key.strip().upper()
                val = val.strip()
                key_map = {
                    "QTY": "QTY ACTUAL",
                    "QTY ACT": "QTY ACTUAL",
                    "QTY ACTUAL": "QTY ACTUAL",
                    "QTY EMRO": "QTY EMRO",
                    "REMARK": "REMARK",
                    "REMARKS": "REMARK",
                    "REMARK(S)": "REMARK",
                }
                key = key_map.get(key, key)
                current[key] = val

    save_current()
    return entries


def create_excel_bytes(entries):
    wb = Workbook()
    ws = wb.active
    ws.title = "Stock Opname"

    headers = [
        "Tanggal", "Pengirim", "LOC", "BIN", "PN", "SN",
        "Qty EMRO", "Qty ACTUAL", "REMARK", "PHOTO FILE", "PHOTO LINK"
    ]
    ws.append(headers)

    col_widths = {
        "A": 18, "B": 25, "C": 10, "D": 10, "E": 20, "F": 20,
        "G": 12, "H": 12, "I": 25, "J": 35, "K": 40
    }
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    for idx, item in enumerate(entries, start=2):
        for col_idx, key in enumerate(headers, start=1):
            ws.cell(row=idx, column=col_idx, value=item.get(key, ""))

    header_font = Font(bold=True)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "A2"

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


# ------------------------------------------------------------
# UI: Upload ZIP dan Token
# ------------------------------------------------------------
uploaded_zip = st.file_uploader("📁 Upload file ZIP WhatsApp", type=["zip"])
uploaded_token = st.file_uploader("🔑 Upload token.pickle", type=["pickle", "pkl", "dat"])

if st.button("🚀 Proses Data"):
    if not uploaded_zip:
        st.warning("Upload file ZIP terlebih dahulu.")
        st.stop()

    creds = get_credentials(uploaded_token)
    service = build_drive_service(creds)

    extract_dir = "temp_extracted"
    if os.path.exists(extract_dir):
        import shutil
        shutil.rmtree(extract_dir)
    os.makedirs(extract_dir, exist_ok=True)

    extract_zip_to_dir(io.BytesIO(uploaded_zip.read()), extract_dir)
    chat_file = find_text_chat_file(extract_dir)
    if not chat_file:
        st.error("❌ File chat .txt tidak ditemukan dalam ZIP.")
        st.stop()

    image_index = index_images(extract_dir)
    entries = parse_chat_file(chat_file)

    st.info(f"📄 Chat ditemukan: {chat_file}")
    st.info(f"🖼️ Total gambar terindeks: {len(image_index)}")
    st.info(f"📊 Total entri ditemukan: {len(entries)}")

    progress = st.progress(0)
    for i, entry in enumerate(entries, start=1):
        photo_name = entry.get("PHOTO FILE", "")
        if photo_name:
            img_path = find_image_fuzzy(photo_name, image_index)
            if img_path and os.path.exists(img_path):
                link = upload_to_drive(service, FOLDER_ID, img_path, entry)
                entry["PHOTO LINK"] = link
            else:
                entry["PHOTO LINK"] = "NOT FOUND"
        else:
            entry["PHOTO LINK"] = "-"
        progress.progress(int(100 * i / len(entries)))

    excel_bytes = create_excel_bytes(entries)
    st.download_button(
        label="📥 Unduh Excel",
        data=excel_bytes,
        file_name=DEFAULT_OUTPUT_FILENAME,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.success("✅ Selesai! Semua foto telah diupload ke Drive dan hasil bisa diunduh.")
