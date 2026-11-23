#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
掃描一個已經完成整理（文字／圖片／影音三條線都確認過）的資料夾，
把裡面每一個檔案記一列進「豆皮檔案清冊.xlsx」總清冊。

這支腳本只做「總清冊」這一步：把已經處理完成的資料夾登記起來。
文字重複比對、圖片查重建議、影音衍生關係判斷，都不是這支腳本的工作，
那些交給 04_coverage_compare.py／digikam_smart_review.py／video_edit_detect.py，
判斷完、人工確認完，才輪到這支腳本把結果記一列進總清冊。

=== 欄位規則（依過去整理 01~05 資料夾時確認過的規則）===
- 每個檔案一列
- 空間名稱：固定填「豆皮」
- 原始位置：這次掃描的資料夾（--folder 參數）的絕對路徑；同一次執行、每一列都一樣
  （這個欄位使用者之後可能會拿掉，目前先照舊保留）
- 資料夾名稱：檔案相對於掃描資料夾的「所在子資料夾路徑」，規則如下：
    - 檔案直接在掃描資料夾底下（沒有子資料夾）→ 填掃描資料夾自己的名稱
    - 有 1 層子資料夾 → 填那層子資料夾的名稱
    - 有 2 層以上子資料夾 → 全部層級用底線 _ 串起來
      （例：中層資料夾/小層資料夾/檔案 → 「中層資料夾_小層資料夾」）
- 檔案名稱：檔案的完整檔名（含副檔名）
- 編號：留空，之後人工填
- 內容物：只填副檔名（例如 .jpg、.docx），統一小寫
- 數量：固定填 1（每列都是一個檔案）
- 狀況描述：留空，之後人工填
- 備注：嘗試自動掃描文件類檔案（.docx/.doc/.rtf/.txt）內文最上方，找「編者附註／紅字
  說明」這類段落並填入；.doc/.rtf 會先用 macOS 內建 textutil 轉檔再讀取；.pages 檔案
  目前無法自動讀取內文（Pages 格式非公開格式，textutil 不支援轉檔），一律留空，並在
  結束時列出清單提醒人工開啟查看

系統檔（.DS_Store、以 ._ 開頭、Thumbs.db）不列入清冊。

=== 追加寫入、不覆蓋 ===
--output 指向已存在的檔案時，會用「(原始位置, 資料夾名稱, 檔案名稱)」三者當唯一鍵，
跟舊清冊比對：已經在清冊裡的列完全不動（人工填過的編號/狀況描述/備注都不會被覆蓋
或洗掉），只有沒登記過的檔案才會新增成一列。所以同一個資料夾要重跑（例如發現漏掃
了幾個檔案）可以直接對同一個 --output 路徑再跑一次。

=== 事前準備 ===
pip3 install --break-system-packages openpyxl python-docx

=== 使用方式 ===
# 掃描一個已經處理完成的資料夾，登記進總清冊（第一次執行，清冊還不存在就會建立）
python3 05_build_inventory.py "/path/to/folder" \\
    --output "/path/to/inventory.xlsx"

# 之後處理下一個資料夾，--output 指向同一份清冊，會自動追加、不覆蓋
python3 05_build_inventory.py "同上層路徑/03_2003勞動藝術" --output "同一份清冊路徑"

參數：
  folder              已經完成整理、要登記進總清冊的資料夾
  --output            豆皮檔案清冊.xlsx 的路徑（不存在就新建，存在就追加合併）
  --original-location 「原始位置」欄要填的值，預設用 folder 的絕對路徑
  --space-name        「空間名稱」欄要填的值，預設「豆皮」
  --no-notes-scan     不掃描備注（跳過文件類檔案的編者附註偵測，加快速度）
  --dry-run           只印出會新增幾筆、不寫入任何檔案
"""

import argparse
import os
import re
import subprocess
import sys

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    print("缺少 openpyxl，請先執行：pip3 install --break-system-packages openpyxl")
    sys.exit(1)

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


SYSTEM_FILE_NAMES = {"Thumbs.db", ".DS_Store"}
CONVERTIBLE_EXTS = (".doc", ".rtf")
TXT_EXTS = (".txt",)
DOCX_EXT = ".docx"
UNREADABLE_EXTS = (".pages",)

HEADERS = [
    "空間名稱", "原始位置", "資料夾名稱", "檔案名稱", "編號",
    "內容物", "數量", "狀況描述", "備注",
]
KEY_COLUMNS = ("原始位置", "資料夾名稱", "檔案名稱")  # 合併去重用的唯一鍵

# 備注掃描：段落文字裡出現這些字樣，就當作「編者附註／紅字說明」段落
NOTE_KEYWORDS = ("紅字", "編者", "附註", "備註說明", "與《", "不同之處")
MAX_NOTE_SCAN_PARAGRAPHS = 15  # 只看文件最前面幾段，避免整份文件內容都被誤判


def is_system_file(name):
    if name in SYSTEM_FILE_NAMES:
        return True
    if name.startswith("._"):
        return True
    return False


def compute_folder_label(root_folder, filepath):
    """依規則算出「資料夾名稱」欄的值。"""
    rel_dir = os.path.relpath(os.path.dirname(filepath), root_folder)
    if rel_dir == ".":
        return os.path.basename(os.path.normpath(root_folder))
    parts = [p for p in rel_dir.split(os.sep) if p not in ("", ".")]
    return "_".join(parts)


def convert_to_docx(filepath):
    """用 macOS 內建 textutil 把 .doc/.rtf 轉成同資料夾的 .docx，回傳新檔案路徑。"""
    base, _ = os.path.splitext(filepath)
    output_path = base + ".docx"
    if os.path.exists(output_path):
        return output_path
    subprocess.run(
        ["textutil", "-convert", "docx", filepath, "-output", output_path],
        check=True,
        capture_output=True,
    )
    return output_path


def read_txt(filepath):
    for enc in ("utf-8", "utf-8-sig", "big5", "gb18030"):
        try:
            with open(filepath, encoding=enc) as fh:
                return fh.read()
        except (UnicodeDecodeError, LookupError):
            continue
    with open(filepath, encoding="utf-8", errors="ignore") as fh:
        return fh.read()


def extract_paragraphs(filepath):
    """回傳這份文件最前面幾段文字（list of str）；讀不到就回傳 []。"""
    ext = os.path.splitext(filepath)[1].lower()
    try:
        if ext == DOCX_EXT:
            if not HAS_DOCX:
                return None  # 表示「應該能讀但缺套件」，跟「真的讀不到」區分
            doc = Document(filepath)
            return [p.text for p in doc.paragraphs if p.text.strip()]
        if ext in TXT_EXTS:
            text = read_txt(filepath)
            return [line for line in text.splitlines() if line.strip()]
        if ext in CONVERTIBLE_EXTS:
            if not HAS_DOCX:
                return None
            converted = convert_to_docx(filepath)
            doc = Document(converted)
            return [p.text for p in doc.paragraphs if p.text.strip()]
    except Exception:
        return None
    return None


def detect_note(filepath):
    """嘗試在文件最前面幾段抓出編者附註／紅字說明段落，抓不到回傳 ''。"""
    ext = os.path.splitext(filepath)[1].lower()
    if ext in UNREADABLE_EXTS:
        return "", "unreadable_format"
    if ext not in (DOCX_EXT,) + CONVERTIBLE_EXTS + TXT_EXTS:
        return "", "not_applicable"

    paragraphs = extract_paragraphs(filepath)
    if paragraphs is None:
        return "", "read_failed"
    if not paragraphs:
        return "", "empty"

    for para in paragraphs[:MAX_NOTE_SCAN_PARAGRAPHS]:
        stripped = para.strip()
        if not stripped:
            continue
        if any(keyword in stripped for keyword in NOTE_KEYWORDS):
            return stripped, "found"
    return "", "not_found"


def find_target_files(root_folder):
    """回傳所有要登記的檔案完整路徑（排除系統檔），依路徑排序。"""
    files = []
    for dirpath, dirnames, filenames in os.walk(root_folder):
        for name in sorted(filenames):
            if is_system_file(name):
                continue
            files.append(os.path.join(dirpath, name))
    return sorted(files)


def load_existing_keys(output_path):
    """回傳 (workbook, worksheet, 既有唯一鍵集合)；檔案不存在就建立新的活頁簿。"""
    if os.path.exists(output_path):
        wb = load_workbook(output_path)
        ws = wb.active
        header_row = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        if header_row != HEADERS:
            raise ValueError(
                f"既有清冊 {output_path} 的欄位跟預期不同（{header_row}），"
                "請確認是不是同一份「豆皮檔案清冊」，避免誤寫。"
            )
        existing_keys = set()
        key_idx = [HEADERS.index(c) for c in KEY_COLUMNS]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and any(row):
                existing_keys.add(tuple(row[i] for i in key_idx))
        return wb, ws, existing_keys

    wb = Workbook()
    ws = wb.active
    ws.title = "豆皮檔案清冊"
    ws.append(HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    return wb, ws, set()


def main():
    parser = argparse.ArgumentParser(description="掃描已完成整理的資料夾，登記進豆皮檔案清冊.xlsx")
    parser.add_argument("folder", help="要登記的資料夾（會自動往下搜尋所有子資料夾）")
    parser.add_argument("--output", required=True, help="豆皮檔案清冊.xlsx 路徑（不存在就新建，存在就追加）")
    parser.add_argument("--original-location", default=None, help="「原始位置」欄要填的值，預設用 folder 的絕對路徑")
    parser.add_argument("--space-name", default="豆皮", help="「空間名稱」欄要填的值，預設「豆皮」")
    parser.add_argument("--no-notes-scan", action="store_true", help="不掃描備注欄（跳過編者附註偵測）")
    parser.add_argument("--dry-run", action="store_true", help="只印出統計，不寫入任何檔案")
    args = parser.parse_args()

    root_folder = os.path.abspath(args.folder)
    if not os.path.isdir(root_folder):
        print(f"❌ 找不到資料夾：{root_folder}")
        sys.exit(1)

    original_location = args.original_location or root_folder
    output_path = os.path.abspath(args.output)

    files = find_target_files(root_folder)
    # 如果清冊本身剛好存在於被掃描的資料夾裡，排除掉它自己
    files = [f for f in files if os.path.abspath(f) != output_path]

    if not files:
        print(f"在 '{root_folder}'（含子資料夾）中沒有找到任何檔案（系統檔已排除）。")
        return

    print(f"共掃描到 {len(files)} 個檔案，讀取既有清冊...")
    wb, ws, existing_keys = load_existing_keys(output_path)

    new_rows = []
    unreadable_notes = []
    skipped_existing = 0

    for filepath in files:
        filename = os.path.basename(filepath)
        folder_label = compute_folder_label(root_folder, filepath)
        ext = os.path.splitext(filename)[1].lower()

        key = (original_location, folder_label, filename)
        if key in existing_keys:
            skipped_existing += 1
            continue

        note = ""
        if not args.no_notes_scan:
            note, status = detect_note(filepath)
            if status in ("read_failed", "unreadable_format"):
                unreadable_notes.append((filepath, status))

        new_rows.append({
            "空間名稱": args.space_name,
            "原始位置": original_location,
            "資料夾名稱": folder_label,
            "檔案名稱": filename,
            "編號": None,
            "內容物": ext,
            "數量": 1,
            "狀況描述": None,
            "備注": note or None,
        })
        existing_keys.add(key)

    print(f"新增 {len(new_rows)} 筆，{skipped_existing} 筆已在清冊中（略過，不覆蓋）。")

    if args.dry_run:
        print("（--dry-run，未寫入任何檔案）")
        return

    for row in new_rows:
        ws.append([row[h] for h in HEADERS])

    col_widths = [10, 46, 34, 30, 8, 10, 6, 30, 40]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    wb.save(output_path)

    print("\n" + "=" * 50)
    print(f"完成！已寫入 {output_path}")
    print(f"清冊目前總筆數（不含表頭）：{ws.max_row - 1}")
    if unreadable_notes:
        print(f"\n⚠️  {len(unreadable_notes)} 份文件類檔案無法自動讀取備注，建議人工開啟查看：")
        for path, status in unreadable_notes:
            reason = "Pages 格式，textutil 不支援轉檔" if status == "unreadable_format" else "讀取失敗"
            print(f"  - {os.path.relpath(path, root_folder)}（{reason}）")
    print("=" * 50)


if __name__ == "__main__":
    main()
