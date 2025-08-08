import os
import json
import requests
from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters

# ================== CẤU HÌNH ==================
TELEGRAM_TOKEN = "7447323319:AAEJEQ22pR-duNY9cflIiu6ibDNujrdir28"   # Token bot Telegram
APP_ID = "cli_a81868fdf5b8d02d"
APP_SECRET = "baPCiHwN0hh0wtFchs5qSd7gstJI2fHr"
SPREADSHEET_ID = "1oGyWPSlfNqTkORFLcXT-FNc_SOMW-07kwavmkG5xpMI"
CREDENTIAL_FILE = "credentials.json"

# Lưu trạng thái người dùng
user_states = {}  # {user_id: {"token": "", "chats": [], "step": ""}}

# ================== CÁC HÀM LARK ==================
def get_tenant_access_token(app_id, app_secret):
    print("Đang lấy token trên lark...")
    url = "https://open.larksuite.com/open-apis/auth/v3/tenant_access_token/internal"
    payload = {
        "app_id": app_id,
        "app_secret": app_secret
    }
    headers = {"Content-Type": "application/json"}
    resp = requests.post(url, json=payload, headers=headers)
    data = resp.json()
    
    if data.get("code") == 0:
        print("Lấy token trên lark thành công")
        return data["tenant_access_token"]
    else:
        raise Exception(f"Lấy tenant_access_token thất bại: {data}")


def get_user_name_by_open_id(open_id, access_token):
    if not open_id:
        return "System"
    url = f"https://open.larksuite.com/open-apis/contact/v3/users/{open_id}"
    headers = {"Authorization": f"Bearer {access_token}"}
    params = {"user_id_type": "open_id"}
    response = requests.get(url, headers=headers, params=params)
    if response.status_code == 200:
        return response.json().get("data", {}).get("user", {}).get("name", open_id)
    return open_id


def get_chat_list(access_token):
    print("Đang lấy danh sách nhóm chat...")
    url = "https://open.larksuite.com/open-apis/im/v1/chats"
    headers = {"Authorization": f"Bearer {access_token}"}
    params = {"page_size": 50}
    chats = []
    while True:
        resp = requests.get(url, headers=headers, params=params)
        if resp.status_code != 200:
            print("Lỗi khi gọi API:", resp.text)
            break
        data = resp.json().get("data", {})
        items = data.get("items", [])
        chats.extend(items)
        if not data.get("has_more"):
            break
        params["page_token"] = data.get("page_token")
    print(f"Lấy thành công {len(chats)} nhóm chat.")
    return chats


def fetch_messages(chat_id, access_token):
    print("Đang lấy tin nhắn từ Lark...")
    url = "https://open.larksuite.com/open-apis/im/v1/messages"
    headers = {"Authorization": f"Bearer {access_token}"}
    params = {
        "container_id_type": "chat",
        "container_id": chat_id,
        "page_size": 50,
    }
    messages = []
    while True:
        response = requests.get(url, headers=headers, params=params)
        if response.status_code != 200:
            print("Lỗi khi gọi API:", response.text)
            break
        data = response.json().get("data", {})
        items = data.get("items", [])
        messages.extend(items)
        if not data.get("has_more"):
            break
        params["page_token"] = data.get("page_token")
    print(f"Lấy thành công {len(messages)} tin nhắn.")
    return messages


def parse_messages(messages, access_token):
    print("Đang cập nhật nội dung tin nhắn để export...")
    parsed = []
    for msg in reversed(messages):
        sender_id = msg.get("sender", {}).get("id")
        sender_name = get_user_name_by_open_id(sender_id, access_token)
        content = json.loads(msg.get("body", {}).get("content", "{}")).get("text", "")
        timestamp = int(msg.get("create_time", 0)) / 1000
        time_str = datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M:%S")
        parsed.append([time_str, sender_name, content])
    return parsed


def write_to_sheet(sheet_id, values):
    creds = service_account.Credentials.from_service_account_file(
        CREDENTIAL_FILE,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
    )
    service = build("sheets", "v4", credentials=creds)
    body = {
        "values": [["Thời gian", "Người gửi", "Nội dung"]] + values
    }
    service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range="A1",
        valueInputOption="RAW",
        body=body
    ).execute()
    print(f"Nội dung đã được tải lên link gg sheet: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")
    return f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"


# ================== HANDLER TELEGRAM ==================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    # Lấy token
    token = get_tenant_access_token(APP_ID, APP_SECRET)
    chats = get_chat_list(token)
    if not chats:
        await update.message.reply_text("Không tìm thấy nhóm chat nào.")
        return
    # Lưu state
    user_states[user_id] = {
        "token": token,
        "chats": chats,
        "step": "choose_chat"
    }
    # Gửi danh sách nhóm chat
    msg = "Danh sách nhóm chat:\n"
    for i, chat in enumerate(chats, 1):
        msg += f"{i}. {chat.get('name')}\n"
    msg += "\nNhập số thứ tự nhóm cần export tin nhắn:"
    await update.message.reply_text(msg)


async def handle_choice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    state = user_states.get(user_id)
    if not state or state.get("step") != "choose_chat":
        await update.message.reply_text("Vui lòng gõ /start để bắt đầu.")
        return

    try:
        choice = int(update.message.text)
    except ValueError:
        await update.message.reply_text("Vui lòng nhập số thứ tự hợp lệ.")
        return

    chats = state["chats"]
    if not (1 <= choice <= len(chats)):
        await update.message.reply_text("Số thứ tự không hợp lệ.")
        return

    chat_id = chats[choice - 1]["chat_id"]
    token = state["token"]

    # Lấy tin nhắn
    messages = fetch_messages(chat_id, token)
    parsed_data = parse_messages(messages, token)
    sheet_link = write_to_sheet(SPREADSHEET_ID, parsed_data)

    await update.message.reply_text(f"✅ Tin nhắn đã được export lên Google Sheet:\n{sheet_link}")
    # Reset state
    user_states.pop(user_id, None)


# ================== MAIN ==================
if __name__ == "__main__":
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_choice))

    print("Bot Telegram đang chạy...")
    app.run_polling()
