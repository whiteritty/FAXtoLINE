import os
import time
import json
import requests
import dropbox
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from datetime import datetime, timedelta, timezone
import re
import pandas as pd
from dropbox.oauth import DropboxOAuth2FlowNoRedirect

# 設定
CONFIG_FILE = "config.json"
TOKEN_EXPIRES_IN = 14400  # トークンの有効期限（4時間）
UNSENT_FILE = "未送信.txt"
WARNING_THRESHOLD = 20  # トークン有効期限の警告秒数

# 設定ファイルの読み込みと初期化
def create_config():
    config = {
        "dropbox_app_key": "YOUR_DROPBOX_APP_KEY",
        "dropbox_app_secret": "YOUR_DROPBOX_APP_SECRET",
        "dropbox_access_token": "",
        "dropbox_refresh_token": "",
        "dropbox_base_folder": "/path/to/dropbox/folder",
        "line_channel_access_token": "YOUR_LINE_CHANNEL_ACCESS_TOKEN",
        "monitor_folder": "/path/to/monitor/folder",
        "excel_file_path": "/path/to/excel.xls"
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)
    print("Config file created")
    return config

def load_config():
    if not os.path.exists(CONFIG_FILE):
        return create_config()
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        config = json.load(f)
    print("Config.jsonを読み込みました")
    return config

def update_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)
    print("Config file updated")

def get_oauth_tokens(config):
    auth_flow = DropboxOAuth2FlowNoRedirect(config["dropbox_app_key"], config["dropbox_app_secret"])
    authorize_url = auth_flow.start()
    print(f"1. Go to: {authorize_url}")
    print("2. Click 'Allow'")
    print("3. Copy the authorization code.")
    auth_code = input("Enter the authorization code here: ").strip()
    print("\n認証コードを受け取りました。トークン取得を開始します...\n")
    try:
        oauth_result = auth_flow.finish(auth_code)
        config["dropbox_access_token"] = oauth_result.access_token
        config["token_obtained_at"] = time.time()
        update_config(config)
        print("アクセストークンの取得に成功しました。\n")
        return oauth_result.access_token
    except Exception as e:
        print(f"OAuth2トークン交換中のエラー: {e}\n")
        return None

def check_and_refresh_token(config):
    remaining_time = get_token_remaining_time(config)
    if remaining_time <= WARNING_THRESHOLD:
        print(f"アクセストークンの有効期限が近づいています。残り: {remaining_time}秒")
        return get_oauth_tokens(config)
    return config["dropbox_access_token"]

def get_token_remaining_time(config):
    current_time = time.time()
    token_obtained_at = config.get("token_obtained_at", 0)
    if token_obtained_at:
        remaining_time = (token_obtained_at + TOKEN_EXPIRES_IN) - current_time
        return max(0, remaining_time)
    return 0

# 未送信ファイルの管理
def read_unsent_files():
    if not os.path.exists(UNSENT_FILE):
        return []
    with open(UNSENT_FILE, "r", encoding="utf-8") as f:
        return [line.strip() for line in f.readlines()]

def write_unsent_file(file_path):
    with open(UNSENT_FILE, "a", encoding="utf-8") as f:
        f.write(file_path + "\n")

def clear_unsent_file():
    if os.path.exists(UNSENT_FILE):
        open(UNSENT_FILE, "w").close()

def process_unsent_files(handler):
    unsent_files = read_unsent_files()
    for file_path in unsent_files:
        file_name = os.path.basename(file_path)
        handler.upload_to_dropbox(file_path, file_name)
    clear_unsent_file()

# Dropbox初期化
def initialize_dropbox(config):
    try:
        token = check_and_refresh_token(config)
        return dropbox.Dropbox(token)
    except Exception as e:
        print(f"Dropbox初期化エラー: {e}")
        return None

config = load_config()
dbx = initialize_dropbox(config)
DROPBOX_BASE_FOLDER = config["dropbox_base_folder"]
LINE_CHANNEL_ACCESS_TOKEN = config["line_channel_access_token"]
MONITOR_FOLDER = config["monitor_folder"]
EXCEL_FILE_PATH = config["excel_file_path"]

class PDFEventHandler(FileSystemEventHandler):
    def __init__(self):
        self.notified_files = set()

    def on_created(self, event):
        if not event.is_directory and event.src_path.endswith(".pdf"):
            file_path = event.src_path
            file_name = os.path.basename(file_path)
            if file_name not in self.notified_files:
                self.notified_files.add(file_name)
                print(f"PDFファイルが検出されました: \n{file_name}")
                time.sleep(3)
                self.upload_to_dropbox(file_path, file_name)

    def upload_to_dropbox(self, file_path, file_name):
        try:
            if not dbx:
                raise Exception("Dropbox client is not initialized.")
            dropbox_path = f"{DROPBOX_BASE_FOLDER}/{file_name}"
            with open(file_path, "rb") as f:
                dbx.files_upload(f.read(), dropbox_path, mode=dropbox.files.WriteMode.overwrite)
            shared_link_metadata = dbx.sharing_create_shared_link_with_settings(dropbox_path)
            shared_link = shared_link_metadata.url.replace("?dl=0", "?dl=1")
            print(f"Dropboxリンクを生成しました:\n {shared_link}")
            name = self.extract_name(file_name) or file_name
            self.notify_line(name, shared_link)
        except Exception as e:
            print(f"アップロード中のエラー: {e}")
            write_unsent_file(file_path)

    def extract_name(self, file_name):
        try:
            base_name = file_name[:-22]  # 右から22文字を削除
            normalized_name = re.sub(r'[\s-]', '', base_name)  # ハイフォンとスペースを削除
            print(f"正規化された名前: {normalized_name}")  # 正規化後の名前をPrint
            if not normalized_name:
                return None
            name = search_fax_destination(normalized_name)
            print(f"送信者名: {name}")  # 取得した送信者名をPrint
            return name
        except Exception as e:
            print(f"Name抽出中のエラー: {e}")
            return None

    def notify_line(self, name, shared_link=None):
        try:
            message = {
                "type": "text",
                "text": f"新しいFAXが届きました:\n{name}\n{shared_link}\n閲覧可能期間は7日間" if shared_link else "エラーが発生しました。"
            }
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {LINE_CHANNEL_ACCESS_TOKEN}"
            }
            response = requests.post(
                "https://api.line.me/v2/bot/message/broadcast",
                headers=headers,
                json={"messages": [message]}
            )
            response.raise_for_status()
            print("LINE通知を送信しました")
        except requests.exceptions.RequestException as e:
            print(f"LINE通知の送信中にエラーが発生しました: {e}")

def search_fax_destination(phone_number):
    try:
        normalized_number = f"[{phone_number}]"
        df = pd.read_excel(EXCEL_FILE_PATH, engine='openpyxl')
        for _, row in df.iterrows():
            if normalized_number == row.get('FAX'):
                return row.get('Name')
        return None
    except Exception as e:
        print(f"Excel処理中のエラー: {e}")
        return None

def start_observer():
    event_handler = PDFEventHandler()
    observer = Observer()
    observer.schedule(event_handler, MONITOR_FOLDER, recursive=False)
    observer.start()
    print(f"{MONITOR_FOLDER} の監視を開始しました")
    process_unsent_files(event_handler)
    try:
        while True:
            time.sleep(3600)
            delete_old_files()
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

def delete_old_files():
    try:
        now = datetime.now(timezone.utc)
        cutoff = now - timedelta(days=7)
        result = dbx.files_list_folder(DROPBOX_BASE_FOLDER)
        for entry in result.entries:
            if isinstance(entry, dropbox.files.FileMetadata) and entry.server_modified < cutoff:
                dbx.files_delete(entry.path_lower)
                print(f"削除しました: {entry.name}")
    except Exception as e:
        print(f"ファイル削除中のエラー: {e}")

if __name__ == "__main__":
    start_observer()
