import os
import tkinter as tk
from tkinter import filedialog, messagebox
from cryptography.fernet import Fernet
import json
import time
import requests
import logging
import mimetypes
import re
import pandas as pd
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import msal
import webbrowser

# --- 設定 ---
SECRET_DIR = "secrets"
SECRET_KEY_FILE = os.path.join(SECRET_DIR, "secret.key")
SECRET_DATA_FILE = os.path.join(SECRET_DIR, "secret_data.enc")
UNSENT_FILE = "未送信.txt"
LOG_FILE = "fax_to_line.log"
TIMEZONE_OFFSET = 9
DELETE_THRESHOLD_DAYS = 7

logging.basicConfig(filename=LOG_FILE, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def ensure_secret_dir():
    if not os.path.exists(SECRET_DIR):
        os.makedirs(SECRET_DIR)

def generate_key(key_path):
    key = Fernet.generate_key()
    with open(key_path, "wb") as key_file:
        key_file.write(key)

def load_key(key_path):
    with open(key_path, "rb") as key_file:
        return key_file.read()

def save_secret(data_dict, key_path, secret_path):
    key = load_key(key_path)
    f = Fernet(key)
    data = json.dumps(data_dict).encode()
    encrypted = f.encrypt(data)
    with open(secret_path, "wb") as secret_file:
        secret_file.write(encrypted)

def load_secret(key_path, secret_path):
    key = load_key(key_path)
    f = Fernet(key)
    with open(secret_path, "rb") as secret_file:
        encrypted = secret_file.read()
    data = f.decrypt(encrypted).decode()
    return json.loads(data)

class SecretInputGUIAll:
    def __init__(self, root, error_fields=None, error_messages=None):
        self.root = root
        self.root.title("API情報入力")
        self.data = {}
        self.entries = {}
        self.error_labels = {}
        self.fields = [
            {"label": "監視するフォルダのパス", "type": "folder", "key": "monitor_folder"},
            {"label": "FAX番号と名前の対応表Excelファイルのパス（不要なら空欄）", "type": "file", "key": "excel_file_path", "optional": True},
            {"label": "OneDriveアプリのClient ID", "type": "text", "key": "onedrive_client_id"},
            {"label": "OneDriveのアップロード先フォルダ名", "type": "text", "key": "onedrive_folder"},
            {"label": "LINEチャネルアクセストークン", "type": "text", "key": "line_token"},
        ]
        for i, field in enumerate(self.fields):
            label = tk.Label(root, text=field["label"])
            label.grid(row=i, column=0, sticky="w", padx=5, pady=2)
            entry = tk.Entry(root, width=50, show="*" if "secret" in field["key"] else "")
            entry.grid(row=i, column=1, padx=5, pady=2)
            self.entries[field["key"]] = entry
            error_label = tk.Label(root, text="", fg="red")
            error_label.grid(row=i, column=2, sticky="w")
            self.error_labels[field["key"]] = error_label
            if field["type"] in ("folder", "file"):
                browse_btn = tk.Button(root, text="参照", command=lambda k=field["key"]: self.browse(k))
                browse_btn.grid(row=i, column=3, padx=2)
            # エラーがある場合は表示
            if error_fields and field["key"] in error_fields:
                msg = error_messages.get(field["key"], "")
                self.error_labels[field["key"]].config(text=msg)
        self.save_btn = tk.Button(root, text="保存", command=self.save)
        self.save_btn.grid(row=len(self.fields), column=0, columnspan=4, pady=10)

    def browse(self, key):
        field = next(f for f in self.fields if f["key"] == key)
        if field["type"] == "folder":
            path = filedialog.askdirectory()
        else:
            path = filedialog.askopenfilename()
        if path:
            self.entries[key].delete(0, tk.END)
            self.entries[key].insert(0, path)

    def save(self):
        valid = True
        data = {}
        for field in self.fields:
            key = field["key"]
            value = self.entries[key].get().strip()
            self.error_labels[key].config(text="")
            if field["type"] == "folder":
                if not os.path.isdir(value):
                    self.error_labels[key].config(text="フォルダが存在しません。")
                    valid = False
            elif field["type"] == "file":
                if value == "" and field.get("optional"):
                    value = None
                elif value and not os.path.isfile(value):
                    self.error_labels[key].config(text="ファイルが存在しません。")
                    valid = False
            else:
                if not value:
                    self.error_labels[key].config(text="入力してください。")
                    valid = False
            data[key] = value
        if valid:
            ensure_secret_dir()
            if not os.path.exists(SECRET_KEY_FILE):
                generate_key(SECRET_KEY_FILE)
            save_secret(data, SECRET_KEY_FILE, SECRET_DATA_FILE)
            messagebox.showinfo("完了", "設定が保存されました。")
            self.root.destroy()

def get_secret_data():
    ensure_secret_dir()
    # 設定ファイルがあれば値を読み込む
    data = {}
    if os.path.exists(SECRET_KEY_FILE) and os.path.exists(SECRET_DATA_FILE):
        data = load_secret(SECRET_KEY_FILE, SECRET_DATA_FILE)
    # 入力フォームを常に表示し、値を編集可能にする
    root = tk.Tk()
    gui = SecretInputGUIAll(root)
    # 既存値をセット
    for key, entry in gui.entries.items():
        if key in data and data[key]:
            entry.insert(0, data[key])
    root.mainloop()
    # 保存後、再度読み込む
    return load_secret(SECRET_KEY_FILE, SECRET_DATA_FILE)

def get_onedrive_token_delegated(client_id, scopes):
    # 個人アカウント用エンドポイント
    authority = "https://login.microsoftonline.com/consumers"
    app = msal.PublicClientApplication(client_id, authority=authority)
    accounts = app.get_accounts()
    if accounts:
        print("既存アカウントでトークン取得を試みます...")
        result = app.acquire_token_silent(scopes, account=accounts[0])
        if result and "access_token" in result:
            print("既存アカウントでトークン取得成功")
            return result["access_token"]
        else:
            print("既存アカウントでのトークン取得失敗、デバイスコードフローに進みます。")
    print("デバイスコードフローを初期化中...")
    flow = app.initiate_device_flow(scopes=scopes)
    if "user_code" not in flow:
        print(f"デバイスフロー初期化エラー: {flow}")
        raise Exception("デバイスフローの初期化に失敗しました: " + str(flow))
    print(flow["message"])
    webbrowser.open(flow["verification_uri"])
    print("ブラウザで認証を完了してください。")
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" in result:
        print("デバイスコードフローでトークン取得成功")
        return result["access_token"]
    else:
        print(f"デバイスコードフロー認証エラー: {result}")
        raise Exception("OneDrive認証失敗: " + str(result))

def upload_to_onedrive(access_token, local_path, onedrive_folder):
    file_name = os.path.basename(local_path)
    mime_type, _ = mimetypes.guess_type(local_path)
    if not mime_type:
        mime_type = "application/octet-stream"
    folder = onedrive_folder.lstrip("/") if onedrive_folder else ""
    if folder:
        upload_url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{folder}/{file_name}:/content"
    else:
        upload_url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_name}:/content"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": mime_type
    }
    with open(local_path, "rb") as f:
        response = requests.put(upload_url, headers=headers, data=f)
    if response.status_code in (200, 201):
        file_info = response.json()
        share_url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_info['id']}/createLink"
        share_body = {"type": "view", "scope": "anonymous"}
        share_resp = requests.post(share_url, headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}, json=share_body)
        share_json = share_resp.json()
        # ここを修正
        if "link" in share_json and "webUrl" in share_json["link"]:
            return share_json["link"]["webUrl"]
        else:
            raise Exception(f"共有リンク作成失敗: {share_resp.text}")
    else:
        raise Exception(f"OneDriveアップロード失敗: {response.text}")

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
        handler.upload_to_onedrive_and_notify(file_path, file_name)
    clear_unsent_file()

if __name__ == "__main__":
    secret = get_secret_data()
    MONITOR_FOLDER = secret["monitor_folder"]
    EXCEL_FILE_PATH = secret.get("excel_file_path")
    ONEDRIVE_CLIENT_ID = secret["onedrive_client_id"]
    ONEDRIVE_FOLDER = secret["onedrive_folder"]
    LINE_CHANNEL_ACCESS_TOKEN = secret["line_token"]
    try:
        print("OneDrive認証中...")
        ONEDRIVE_ACCESS_TOKEN = get_onedrive_token_delegated(
            ONEDRIVE_CLIENT_ID,
            scopes=["Files.ReadWrite.All"]
        )
        print("OneDrive認証成功")
    except Exception as e:
        print(f"OneDrive認証エラー: {e}")
        logging.error(f"OneDrive認証エラー: {e}")
        exit(1)

    class PDFEventHandler(FileSystemEventHandler):
        def __init__(self):
            self.notified_files = set()
        def on_created(self, event):
            print(f"on_createdイベント検知: {event.src_path}")
            logging.info(f"on_createdイベント検知: {event.src_path}")
            if not event.is_directory and event.src_path.endswith(".pdf"):
                file_path = event.src_path
                file_name = os.path.basename(file_path)
                if file_name not in self.notified_files:
                    self.notified_files.add(file_name)
                    print(f"PDFファイル検知: {file_name}")
                    logging.info(f"PDFファイル検知: {file_name}")
                    time.sleep(3)
                    self.upload_to_onedrive_and_notify(file_path, file_name)
        def upload_to_onedrive_and_notify(self, file_path, file_name):
            try:
                print(f"OneDriveへアップロード中: {file_name}")
                logging.info(f"OneDriveへアップロード中: {file_name}")
                shared_link = upload_to_onedrive(ONEDRIVE_ACCESS_TOKEN, file_path, ONEDRIVE_FOLDER)
                print(f"OneDriveアップロード成功: {shared_link}")
                logging.info(f"OneDriveアップロード成功: {shared_link}")
                name = self.extract_name(file_name) or file_name
                self.notify_line(name, shared_link)
            except Exception as e:
                print(f"アップロードエラー: {e}")
                logging.error(f"アップロードエラー: {e}")
                write_unsent_file(file_path)
        def extract_name(self, file_name):
            try:
                base_name = file_name[:-22]
                normalized_name = re.sub(r'[\s-]', '', base_name)
                logging.info(f"Normalized name: {normalized_name}")
                if not normalized_name:
                    return None
                if not normalized_name[0].isdigit():
                    return normalized_name
                name = search_fax_destination(normalized_name)
                logging.info(f"Sender name: {name}")
                return name
            except Exception as e:
                print(f"名前抽出エラー: {e}")
                logging.error(f"名前抽出エラー: {e}")
                return None
        def notify_line(self, name, shared_link=None):
            try:
                print(f"LINE通知送信中: {name}")
                logging.info(f"LINE通知送信中: {name}")
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
                print("LINE通知送信成功")
                logging.info("LINE通知送信成功")
            except requests.exceptions.RequestException as e:
                print(f"LINE通知エラー: {e}")
                logging.error(f"LINE通知エラー: {e}")

    def search_fax_destination(phone_number):
        if not EXCEL_FILE_PATH:
            return None
        try:
            normalized_number = f"[{phone_number}]"
            df = pd.read_excel(EXCEL_FILE_PATH, engine='openpyxl')
            for _, row in df.iterrows():
                if normalized_number == row.get('FAX'):
                    return row.get('Name')
            return None
        except FileNotFoundError:
            print(f"Excelファイルが見つかりません: {EXCEL_FILE_PATH}")
            logging.warning(f"Excelファイルが見つかりません: {EXCEL_FILE_PATH}")
            return None
        except Exception as e:
            print(f"Excel処理エラー: {e}")
            logging.error(f"Excel処理エラー: {e}")
            return None

    def start_observer():
        event_handler = PDFEventHandler()
        observer = Observer()
        observer.schedule(event_handler, MONITOR_FOLDER, recursive=False)
        observer.start()
        print(f"監視開始: {MONITOR_FOLDER}")
        logging.info(f"監視開始: {MONITOR_FOLDER}")
        process_unsent_files(event_handler)
        try:
            while True:
                time.sleep(3600)
        except KeyboardInterrupt:
            print("監視停止")
            logging.info("監視停止")
            observer.stop()
        observer.join()

    start_observer()