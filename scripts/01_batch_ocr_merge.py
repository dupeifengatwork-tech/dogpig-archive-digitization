#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批次將資料夾內的圖片上傳至 Google Drive，
利用 Drive 內建的「開啟為 Google 文件」功能觸發 OCR，
辨識完成後匯出純文字，並依檔名（去除副檔名）作為標題，
合併成一份輸出檔案（txt 或 docx）。

=== 事前準備（只需做一次） ===
1. 到 https://console.cloud.google.com/ 建立一個專案
2. 啟用「Google Drive API」
   （左側選單 -> API 和服務 -> 程式庫 -> 搜尋 Google Drive API -> 啟用）
3. 建立憑證：API 和服務 -> 憑證 -> 建立憑證 -> OAuth 用戶端 ID
   -> 應用程式類型選「電腦版應用程式」
   -> 下載 JSON，改名為 credentials.json，放在跟本腳本同一個資料夾
4. 如果跳出「未驗證應用程式」畫面，選「進階」->「前往...(不安全)」即可
   （這是你自己的私人專案，僅供自己使用，不影響安全性）
5. 安裝套件：
   pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib python-docx

第一次執行會跳出瀏覽器要求登入 Google 帳號並授權，
授權後會在同資料夾產生 token.json，之後就不用再登入。

=== 使用方式 ===
python batch_ocr_merge.py "/Users/c-lab/Desktop/佩峰/豆皮丸jpg檔(300小)" --output merged.docx

參數：
  folder            要處理的圖片資料夾路徑（必填）
  --output          輸出檔名，預設 merged.txt。副檔名 .docx 會輸出 Word 檔（含標題樣式）
  --keep-drive-docs 加這個參數的話，轉出的 Google 文件會保留在雲端硬碟，
                    不加的話預設會在辨識完後自動丟到垃圾桶（不佔用你的雲端空間）
"""

import os
import sys
import io
import time
import argparse

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

SCOPES = ["https://www.googleapis.com/auth/drive"]
SUPPORTED_EXTS = (".jpg", ".jpeg", ".JPG", ".JPEG", ".png", ".PNG")


def get_drive_service():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists("credentials.json"):
                print("❌ 找不到 credentials.json，請先依照腳本開頭的說明設定 OAuth 憑證。")
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return build("drive", "v3", credentials=creds)


def ocr_via_google_docs(service, image_path, keep_drive_doc=False):
    """上傳圖片並轉存為 Google 文件（觸發 OCR），回傳辨識出的純文字。"""
    filename = os.path.basename(image_path)

    file_metadata = {
        "name": filename,
        "mimeType": "application/vnd.google-apps.document",
    }
    media = MediaFileUpload(image_path, mimetype="image/jpeg", resumable=False)

    uploaded = service.files().create(
        body=file_metadata, media_body=media, fields="id"
    ).execute()
    doc_id = uploaded["id"]

    # 匯出為純文字
    request = service.files().export_media(fileId=doc_id, mimeType="text/plain")
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    text = buf.getvalue().decode("utf-8", errors="ignore").strip()

    if keep_drive_doc:
        # 移出「我的雲端硬碟」根目錄以外沒特別處理，保留在雲端
        pass
    else:
        # 辨識完直接丟垃圾桶，避免累積一堆暫存文件
        service.files().update(fileId=doc_id, body={"trashed": True}).execute()

    return text


def write_txt(results, output_path):
    with open(output_path, "w", encoding="utf-8") as f:
        for name, text in results:
            f.write(f"{name}\n")
            f.write("=" * len(name) + "\n\n")
            f.write(text + "\n\n")
            f.write("-" * 40 + "\n\n")


def write_docx(results, output_path):
    from docx import Document
    doc = Document()
    for name, text in results:
        doc.add_heading(name, level=1)
        for para in text.split("\n"):
            if para.strip():
                doc.add_paragraph(para)
        doc.add_page_break()
    doc.save(output_path)


def main():
    parser = argparse.ArgumentParser(description="批次 OCR（透過 Google 文件）並合併輸出")
    parser.add_argument("folder", help="圖片資料夾路徑")
    parser.add_argument("--output", default="merged.txt", help="輸出檔名（.txt 或 .docx）")
    parser.add_argument("--keep-drive-docs", action="store_true",
                         help="保留上傳到雲端硬碟的暫存 Google 文件（預設會自動刪除）")
    args = parser.parse_args()

    folder = args.folder
    if not os.path.isdir(folder):
        print(f"❌ 找不到資料夾：{folder}")
        sys.exit(1)

    image_files = sorted(
        f for f in os.listdir(folder) if f.endswith(SUPPORTED_EXTS)
    )
    if not image_files:
        print(f"在 '{folder}' 中沒有找到圖片檔！")
        sys.exit(0)

    print(f"找到 {len(image_files)} 張圖片，開始上傳並以 Google 文件 OCR 辨識...\n")
    service = get_drive_service()

    results = []
    for idx, filename in enumerate(image_files, 1):
        heading = os.path.splitext(filename)[0]  # 檔名去除副檔名 -> 例如 01丸Av
        image_path = os.path.join(folder, filename)
        print(f"[{idx}/{len(image_files)}] 處理中: {filename} ...", end="", flush=True)
        try:
            text = ocr_via_google_docs(service, image_path, keep_drive_doc=args.keep_drive_docs)
            results.append((heading, text))
            print(f" ✅ 完成（{len(text)} 字）")
        except Exception as e:
            print(f" ❌ 出錯: {e}")
            results.append((heading, f"[辨識失敗: {e}]"))
        time.sleep(0.5)  # 避免短時間內請求過於密集

    output_path = os.path.join(os.getcwd(), args.output)
    if output_path.lower().endswith(".docx"):
        write_docx(results, output_path)
    else:
        write_txt(results, output_path)

    print("\n" + "=" * 50)
    print(f"全部完成！已合併輸出至：{output_path}")
    print("=" * 50)


if __name__ == "__main__":
    main()
