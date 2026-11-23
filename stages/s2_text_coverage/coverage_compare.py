#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批次計算「獨立歷史文件」與「豆皮丸 merged 資料庫」的內容覆蓋率，
自動分類大幅刊載／部分刊載／零星相符／未刊載，並回報每份文件最可能
刊登在第幾期、第幾面，輸出成 Excel（三分頁：比對總表、頁面命中明細、
說明）。

這支腳本把 Emily 09/16 手動整理的「v10 比對表」背後的判斷邏輯，重寫成
一份可重複執行的批次工具，比對方法完全比照她表格「說明」分頁記錄的規則：

=== 比對方法 ===
以連續 6 個中文字（可用 --min-match 調整）為單位，檢查文件裡的每一個
連續片段是否曾出現在 merged 資料庫的任何地方；命中的位置合併起來，算出
「這份文件有多少比例的內容出現在資料庫裡」。比對前只保留中文字（去除所有
標點、空白與英數），消除 OCR 標點差異造成的誤判。

=== 判定門檻（沿用 v10 版本，可用參數調整）===
  大幅刊載        整體命中率 >= 50%
  部分刊載        整體命中率 >= 12%，或最長連續相符 >= 25 字
  零星相符        5% <= 整體命中率 < 12%（多半是人名/機構名等通用詞）
  未刊載          整體命中率 < 5%
  內容過短，略過   中文字數 <= 6（不足以組成一個比對單位）
  讀取失敗        文件無法開啟/解析（例如檔案損毀）

=== 事前準備 ===
pip3 install --break-system-packages python-docx openpyxl
（處理 .doc / .rtf 需要 macOS 內建的 textutil，Windows/Linux 無法轉檔，
 遇到這類檔案會直接標記「讀取失敗」，並在結束時列出清單。）

=== 使用方式 ===
python3 04_coverage_compare.py "要比對的資料夾" \
    --database "豆皮丸辨識結果_merged.docx" \
    --output "比對結果.xlsx"

--output 指向已存在的檔案時，會用「相對路徑」當唯一鍵跟舊結果合併：
已經在 Excel 裡勾選過「已複核」的列會保留勾選狀態，只有新出現/內容有
變動的檔案才會回到未勾選；如果同一份檔案重新掃到但雜湊值不同（內容被
改過），命中率等欄位會用新結果覆蓋，但複核勾選狀態保留讓妳自己決定要
不要重新看一次。

參數：
  folder            要比對的資料夾（會自動往下搜尋所有子資料夾）
  --database        豆皮丸資料庫合併文件路徑（必填，例如 merged.docx）
  --output          輸出/合併的 Excel 路徑（必填）
  --min-match       比對單位的字元長度，預設 6
  --min-content-len 中文字數低於等於這個值就跳過比對，預設等於 --min-match
  --large-threshold 大幅刊載的命中率門檻（%），預設 50
  --partial-threshold 部分刊載的命中率門檻（%），預設 12
  --partial-run-threshold 部分刊載的最長連續相符門檻，預設 25
  --scatter-threshold 零星相符的命中率下限（%），預設 5
  --no-recursive    只掃描指定的這一層資料夾，不搜尋子資料夾
  --source-pattern  辨識資料庫出處標題的正規表示式
                    （預設抓「檔案：xxx.jpg」，跟 03_mark_duplicates_new.py 共用邏輯）
"""

import argparse
import os
import re
import subprocess
import sys
from collections import defaultdict
from datetime import datetime

try:
    from docx import Document
except ImportError:
    print("缺少 python-docx，請先執行：pip3 install --break-system-packages python-docx")
    sys.exit(1)

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    print("缺少 openpyxl，請先執行：pip3 install --break-system-packages openpyxl")
    sys.exit(1)


CONVERTIBLE_EXTS = (".doc", ".rtf")
TXT_EXTS = (".txt",)
DOCX_EXT = ".docx"

IMAGE_EXTS = (".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp", ".heic")

DEFAULT_SOURCE_PATTERN = r"^\s*檔案\s*[：:]\s*(\S.*?)\s*$"
MAX_SOURCE_LINE_LEN = 60

# 只保留中文字（CJK 統一表意文字 + 擴充區塊），比對前去除所有標點、空白、英數，
# 跟 v10 表格「說明」分頁記載的規則一致。
CJK_RE = re.compile(r"[㐀-䶿一-鿿豈-﫿]")

LABEL_RE = re.compile(r"^(\d+)丸([A-Za-z])v?$")

CATEGORY_LARGE = "大幅刊載"
CATEGORY_PARTIAL = "部分刊載"
CATEGORY_SCATTER = "零星相符（多為通用詞）"
CATEGORY_NONE = "未刊載"
CATEGORY_TOO_SHORT = "內容過短，略過"
CATEGORY_FAILED = "讀取失敗"

CATEGORY_FILL = {
    CATEGORY_LARGE: "F4CCCC",
    CATEGORY_PARTIAL: "FCE5CD",
    CATEGORY_SCATTER: "FFF2CC",
    CATEGORY_NONE: "FFFFFF",
    CATEGORY_TOO_SHORT: "D9D9D9",
    CATEGORY_FAILED: "CFE2F3",
}

REVIEW_HEADER = "已複核"


def normalize(text):
    """只留下中文字，回傳純中文字串。"""
    return "".join(CJK_RE.findall(text))


def is_heading_paragraph(paragraph):
    try:
        style_name = paragraph.style.name or ""
    except Exception:
        return False
    lowered = style_name.lower()
    return lowered.startswith("heading") or lowered.startswith("title") or "標題" in style_name


def clean_label(label):
    label = label.strip()
    lowered = label.lower()
    for ext in IMAGE_EXTS:
        if lowered.endswith(ext):
            return label[: -len(ext)].strip()
    return label


def detect_source_label(paragraph, regex):
    text = paragraph.text.strip()
    if not text or len(text) > MAX_SOURCE_LINE_LEN:
        return None
    if regex is not None:
        m = regex.search(text)
        if m:
            label = m.group(1) if m.groups() else m.group(0)
            label = clean_label(label)
            if label:
                return label
    return None


def label_to_issue_page(label):
    """把「31丸Av」這種出處名稱拆成 (期數字串, 頁面檔名)。無法辨識就回傳 (label, label)。"""
    m = LABEL_RE.match(label)
    if not m:
        return label, label
    digits, side = m.groups()
    issue = f"第{int(digits)}期"
    page_file = f"{label}.jpg"
    return issue, page_file


def load_reference(database_path, source_pattern, min_match):
    """
    讀取資料庫，回傳：
      db_text        資料庫整體正規化後的中文字串
      segments       [(start, end, label), ...]
      global_ngrams  資料庫裡出現過的所有 n-gram 字串集合（用來算整體命中率）
      ngram_to_labels  n-gram -> 出處集合（用來算每頁命中率、判斷主要刊載期數）
    """
    doc = Document(database_path)
    regex = re.compile(source_pattern) if source_pattern else None

    parts = []
    segments = []
    total = 0
    current_label = None
    current_start = 0

    for p in doc.paragraphs:
        label = detect_source_label(p, regex)
        if label:
            if label != current_label:
                if current_label is not None and total > current_start:
                    segments.append((current_start, total, current_label))
                current_label = label
                current_start = total
            continue  # 「檔案：xxx.jpg」這行是合併時加的標題，不列入比對內容
        if is_heading_paragraph(p):
            continue  # 例如「豆皮丸辨識結果」這種大標題，不是原始內容

        norm = normalize(p.text)
        if norm:
            parts.append(norm)
            total += len(norm)

    if current_label is not None and total > current_start:
        segments.append((current_start, total, current_label))

    db_text = "".join(parts)

    global_ngrams = set()
    ngram_to_labels = defaultdict(set)
    for start, end, label in segments:
        seg_text = db_text[start:end]
        for i in range(len(seg_text) - min_match + 1):
            gram = seg_text[i : i + min_match]
            global_ngrams.add(gram)
            ngram_to_labels[gram].add(label)

    return db_text, segments, global_ngrams, ngram_to_labels


def longest_run_of_ones(covered):
    best = 0
    cur = 0
    for v in covered:
        if v:
            cur += 1
            best = max(best, cur)
        else:
            cur = 0
    return best


def read_txt(filepath):
    for enc in ("utf-8", "utf-8-sig", "big5", "gb18030"):
        try:
            with open(filepath, encoding=enc) as fh:
                return fh.read()
        except (UnicodeDecodeError, LookupError):
            continue
    with open(filepath, encoding="utf-8", errors="ignore") as fh:
        return fh.read()


def read_docx(filepath):
    doc = Document(filepath)
    return "\n".join(p.text for p in doc.paragraphs)


def convert_to_docx(filepath):
    """用 macOS 內建 textutil 把 .doc/.rtf 轉成同資料夾的 .docx。非 macOS 會直接失敗。"""
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


def extract_text(filepath):
    """回傳 (原始文字, 抽取方式)；失敗時丟出例外，由呼叫端標記讀取失敗。"""
    ext = os.path.splitext(filepath)[1].lower()
    if ext == DOCX_EXT:
        return read_docx(filepath), "textutil"
    if ext in TXT_EXTS:
        return read_txt(filepath), "txt"
    if ext in CONVERTIBLE_EXTS:
        converted = convert_to_docx(filepath)
        return read_docx(converted), "textutil"
    raise ValueError(f"不支援的副檔名：{ext}")


def find_target_files(root_folder, recursive):
    files = []
    if recursive:
        walker = os.walk(root_folder)
    else:
        walker = [(root_folder, [], os.listdir(root_folder))]
    for dirpath, _, filenames in walker:
        for name in filenames:
            if name.startswith("~$") or name.startswith("."):
                continue
            ext = os.path.splitext(name)[1].lower()
            if ext in (DOCX_EXT,) + CONVERTIBLE_EXTS + TXT_EXTS:
                files.append(os.path.join(dirpath, name))
    return sorted(files)


def classify(n, hit_rate, longest_run, min_content_len,
             large_threshold, partial_threshold, partial_run_threshold, scatter_threshold):
    if n <= min_content_len:
        return CATEGORY_TOO_SHORT
    if hit_rate >= large_threshold:
        return CATEGORY_LARGE
    if hit_rate >= partial_threshold or longest_run >= partial_run_threshold:
        return CATEGORY_PARTIAL
    if hit_rate >= scatter_threshold:
        return CATEGORY_SCATTER
    return CATEGORY_NONE


def analyze_file(filepath, root_folder, global_ngrams, ngram_to_labels, min_match, args):
    rel_path = os.path.relpath(filepath, root_folder)
    try:
        raw_text, method = extract_text(filepath)
    except Exception as e:
        return {
            "rel_path": rel_path,
            "filename": os.path.basename(filepath),
            "ext": os.path.splitext(filepath)[1].lower(),
            "char_count": 0,
            "category": CATEGORY_FAILED,
            "hit_rate": None,
            "longest_run": None,
            "primary_issue": None,
            "primary_page": None,
            "all_issues": None,
            "method": "FAILED",
            "page_rows": [],
            "error": str(e),
        }

    normalized = normalize(raw_text)
    n = len(normalized)

    if n <= args.min_content_len:
        return {
            "rel_path": rel_path,
            "filename": os.path.basename(filepath),
            "ext": os.path.splitext(filepath)[1].lower(),
            "char_count": n,
            "category": CATEGORY_TOO_SHORT,
            "hit_rate": None,
            "longest_run": None,
            "primary_issue": None,
            "primary_page": None,
            "all_issues": None,
            "method": method,
            "page_rows": [],
            "error": None,
        }

    covered = bytearray(n)
    label_covered = defaultdict(lambda: bytearray(n))

    if n >= min_match:
        for i in range(n - min_match + 1):
            window = normalized[i : i + min_match]
            if window in global_ngrams:
                for k in range(i, i + min_match):
                    covered[k] = 1
                labels = ngram_to_labels.get(window)
                if labels:
                    for label in labels:
                        arr = label_covered[label]
                        for k in range(i, i + min_match):
                            arr[k] = 1

    hit_count = sum(covered)
    hit_rate = hit_count / n * 100
    longest_run = longest_run_of_ones(covered)

    page_rows = []
    for label, arr in label_covered.items():
        pc = sum(arr)
        if pc == 0:
            continue
        page_rate = pc / n * 100
        page_longest = longest_run_of_ones(arr)
        issue, page_file = label_to_issue_page(label)
        page_rows.append((issue, page_file, page_rate, page_longest))
    page_rows.sort(key=lambda x: -x[2])

    category = classify(
        n, hit_rate, longest_run, args.min_content_len,
        args.large_threshold, args.partial_threshold,
        args.partial_run_threshold, args.scatter_threshold,
    )

    if category in (CATEGORY_NONE,):
        primary_issue = primary_page = all_issues = None
    else:
        if page_rows:
            primary_issue, primary_page = page_rows[0][0], page_rows[0][1]
            seen = []
            for issue, _page, _rate, _run in page_rows:
                if issue not in seen:
                    seen.append(issue)
            all_issues = "、".join(seen)
        else:
            primary_issue = primary_page = all_issues = None

    return {
        "rel_path": rel_path,
        "filename": os.path.basename(filepath),
        "ext": os.path.splitext(filepath)[1].lower(),
        "char_count": n,
        "category": category,
        "hit_rate": round(hit_rate, 1) if category not in (CATEGORY_TOO_SHORT,) else None,
        "longest_run": longest_run if category not in (CATEGORY_TOO_SHORT,) else None,
        "primary_issue": primary_issue,
        "primary_page": primary_page,
        "all_issues": all_issues,
        "method": method,
        "page_rows": page_rows,
        "error": None,
    }


def load_existing_review_state(output_path):
    """回傳 {相對路徑: 已複核布林值}，供合併用。"""
    if not os.path.exists(output_path):
        return {}
    try:
        wb = load_workbook(output_path)
        if "比對總表" not in wb.sheetnames:
            return {}
        ws = wb["比對總表"]
        headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        if REVIEW_HEADER not in headers or "相對路徑" not in headers:
            return {}
        review_idx = headers.index(REVIEW_HEADER)
        path_idx = headers.index("相對路徑")
        result = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[path_idx]:
                result[row[path_idx]] = bool(row[review_idx])
        return result
    except Exception as e:
        print(f"⚠️  讀取舊 Excel 複核狀態時發生問題（將視為沒有舊狀態）：{e}")
        return {}


SUMMARY_HEADERS = [
    REVIEW_HEADER, "相對路徑", "檔名", "副檔名", "中文字數", "比對結果",
    "整體命中率%", "最長連續相符字數", "主要刊載期數", "主要刊載頁面",
    "所有相符期數", "文字抽取方式",
]

DETAIL_HEADERS = ["相對路徑", "期數", "面", "頁面檔名", "該頁命中率%", "最長連續相符字數"]

EXPLANATION_ROWS = [
    ["比對方法", None],
    ["以連續 6 個中文字為單位，計算文件有多少比例出現在 merged 全文中。", None],
    ["比對前移除所有標點、空白與英數，以消除 OCR 的標點差異。", None],
    [None, None],
    ["判定門檻", None],
    ["大幅刊載", "整體命中率 ≥ 50%"],
    ["部分刊載", "命中率 ≥ 12%，或最長連續相符 ≥ 25 字（刊物只節錄一段時屬此類）"],
    ["最長連續相符字數", "整份文件中與 merged 連續吻合最長的一段有多少字。這是「確實刊登過」的硬證據——人名機構名等通用詞湊不出連續 25 字以上的相符。"],
    ["零星相符", "5% ≤ 命中率 < 12%（多半只是人名、機構名等通用詞）"],
    ["未刊載", "命中率 < 5%"],
    [None, None],
    ["注意", None],
    ["merged 為 OCR 結果，錯字會壓低命中率；命中率偏低不代表一定沒刊載，", None],
    ["「部分刊載」與「零星相符」兩類建議人工覆核。", None],
    ["期數取自頁面檔名（如 31丸Av.jpg → 第31期 A 面）。", None],
]


def write_excel(output_path, results, review_state, folder):
    wb = Workbook()

    ws = wb.active
    ws.title = "比對總表"
    ws.append(SUMMARY_HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    r = 2
    for res in results:
        reviewed = review_state.get(res["rel_path"], False)
        ws.cell(row=r, column=1, value=reviewed)
        ws.cell(row=r, column=2, value=res["rel_path"])
        ws.cell(row=r, column=3, value=res["filename"])
        ws.cell(row=r, column=4, value=res["ext"])
        ws.cell(row=r, column=5, value=res["char_count"])
        cat_cell = ws.cell(row=r, column=6, value=res["category"])
        ws.cell(row=r, column=7, value=res["hit_rate"])
        ws.cell(row=r, column=8, value=res["longest_run"])
        ws.cell(row=r, column=9, value=res["primary_issue"])
        ws.cell(row=r, column=10, value=res["primary_page"])
        ws.cell(row=r, column=11, value=res["all_issues"])
        ws.cell(row=r, column=12, value=res["method"])
        fill_hex = CATEGORY_FILL.get(res["category"])
        if fill_hex:
            cat_cell.fill = PatternFill("solid", fgColor=fill_hex)
        r += 1

    col_widths = [8, 46, 26, 8, 10, 12, 12, 14, 12, 14, 30, 12]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    detail = wb.create_sheet("頁面命中明細")
    detail.append(DETAIL_HEADERS)
    for cell in detail[1]:
        cell.font = Font(bold=True)
    r = 2
    for res in results:
        for issue, page_file, rate, run in res["page_rows"]:
            m = LABEL_RE.match(page_file[:-4]) if page_file else None
            side = m.group(2).upper() if m else ""
            detail.cell(row=r, column=1, value=res["rel_path"])
            detail.cell(row=r, column=2, value=issue)
            detail.cell(row=r, column=3, value=side)
            detail.cell(row=r, column=4, value=page_file)
            detail.cell(row=r, column=5, value=round(rate, 1))
            detail.cell(row=r, column=6, value=float(run))
            r += 1
    detail_widths = [46, 10, 6, 14, 14, 16]
    for i, w in enumerate(detail_widths, start=1):
        detail.column_dimensions[get_column_letter(i)].width = w

    explain = wb.create_sheet("說明")
    for row in EXPLANATION_ROWS:
        explain.append(row)
    explain.column_dimensions["A"].width = 40
    explain.column_dimensions["B"].width = 70

    summary = wb.create_sheet("摘要", 0)
    counts = defaultdict(int)
    for res in results:
        counts[res["category"]] += 1
    summary.append(["產生時間", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    summary.append(["比對資料夾", folder])
    summary.append(["檔案總數", len(results)])
    summary.append([])
    summary.append(["分類", "檔案數"])
    for cat in (CATEGORY_LARGE, CATEGORY_PARTIAL, CATEGORY_SCATTER, CATEGORY_NONE,
                CATEGORY_TOO_SHORT, CATEGORY_FAILED):
        if counts.get(cat):
            summary.append([cat, counts[cat]])
    summary.column_dimensions["A"].width = 22
    summary.column_dimensions["B"].width = 40

    wb.save(output_path)


def main():
    parser = argparse.ArgumentParser(description="批次計算獨立文件與豆皮丸 merged 資料庫的內容覆蓋率")
    parser.add_argument("folder", help="要比對的資料夾（會自動往下搜尋所有子資料夾）")
    parser.add_argument("--database", required=True, help="豆皮丸資料庫合併文件路徑（例如 merged.docx）")
    parser.add_argument("--output", required=True, help="輸出/合併的 Excel 路徑")
    parser.add_argument("--min-match", type=int, default=6, help="比對單位的字元長度，預設 6")
    parser.add_argument("--min-content-len", type=int, default=None,
                        help="中文字數低於等於這個值就跳過比對，預設等於 --min-match")
    parser.add_argument("--large-threshold", type=float, default=50.0)
    parser.add_argument("--partial-threshold", type=float, default=12.0)
    parser.add_argument("--partial-run-threshold", type=float, default=25.0)
    parser.add_argument("--scatter-threshold", type=float, default=5.0)
    parser.add_argument("--no-recursive", action="store_true")
    parser.add_argument("--source-pattern", default=DEFAULT_SOURCE_PATTERN)
    args = parser.parse_args()

    if args.min_content_len is None:
        args.min_content_len = args.min_match

    print("讀取豆皮丸資料庫...")
    db_text, segments, global_ngrams, ngram_to_labels = load_reference(
        args.database, args.source_pattern, args.min_match
    )
    print(f"資料庫共 {len(db_text)} 個中文字，辨識到 {len(segments)} 個出處分頁。\n")

    files = find_target_files(args.folder, recursive=not args.no_recursive)
    if not files:
        print(f"在 '{args.folder}'（含子資料夾）中沒有找到任何可處理的文件（.docx/.doc/.rtf/.txt）。")
        return

    print(f"共 {len(files)} 份文件，開始比對...\n")
    results = []
    failed = []
    for idx, filepath in enumerate(files, 1):
        rel = os.path.relpath(filepath, args.folder)
        print(f"[{idx}/{len(files)}] {rel} ...", end="", flush=True)
        res = analyze_file(filepath, args.folder, global_ngrams, ngram_to_labels, args.min_match, args)
        results.append(res)
        if res["category"] == CATEGORY_FAILED:
            failed.append((rel, res["error"]))
            print(f" ❌ 讀取失敗：{res['error']}")
        else:
            hr = f"{res['hit_rate']}%" if res["hit_rate"] is not None else "-"
            print(f" {res['category']}（{hr}）")

    review_state = load_existing_review_state(args.output)
    write_excel(args.output, results, review_state, args.folder)

    print("\n" + "=" * 50)
    print(f"完成！已寫入 {args.output}")
    counts = defaultdict(int)
    for res in results:
        counts[res["category"]] += 1
    for cat, n in counts.items():
        print(f"  {cat}: {n}")
    if failed:
        print(f"\n⚠️  {len(failed)} 份檔案讀取失敗，建議人工開啟確認：")
        for rel, err in failed:
            print(f"  - {rel}（{err}）")
    review_carry = sum(1 for res in results if review_state.get(res["rel_path"]))
    if review_carry:
        print(f"\n已從舊檔案沿用 {review_carry} 筆「已複核」勾選狀態。")
    print("=" * 50)


if __name__ == "__main__":
    main()
