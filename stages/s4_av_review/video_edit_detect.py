#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
video_edit_detect.py
====================
針對「現場表演影片檔案庫」設計的重複／剪輯關係偵測工具。

要解決的三種情況（不只是找「完全一樣的檔案」）：
  1. 完全重複（Exact / Near-duplicate）：兩個檔案內容幾乎一樣（可能只是轉檔、改解析度）。
  2. 子片段（Trimmed subclip）：一部短片，其內容其實是某一部長片裡剪出來的一段。
  3. 疑似合併（Merged / Concatenated）：一部長片，其內容其實是好幾部短片依序接起來的。

核心方法：聲音指紋序列比對（Chromaprint / fpcalc），畫面 pHash 序列作為第二重驗證。
現場表演影片的「現場聲音」是連續且獨特的，比畫面更不容易因剪輯、轉碼、調色、改解析度
而失真，所以用聲音指紋序列做逐秒比對，比單純比較整個檔案的雜湊值更可靠，也比只看畫面更省算力。

格式無關：所有解析都透過 ffmpeg / ffprobe / fpcalc 進行，這三個工具本身就能處理
mp4、mov、m4v、avi、mkv、wmv、flv、mts 等常見容器格式，不需要針對每種格式另外寫邏輯。

=== 事前準備（macOS）===
    brew install ffmpeg chromaprint
    pip3 install --break-system-packages numpy openpyxl

=== 使用方式 ===
    # 第一次執行：掃描資料夾、建立指紋快取、輸出報告
    python3 video_edit_detect.py "/path/to/folder"

    # 之後重新執行（例如資料夾內容有變動）：已經算過指紋的檔案不會重算，
    # 只處理新增或修改過的檔案，速度快很多，結果可重現。
    python3 video_edit_detect.py "同一個資料夾" --cache .video_fp_cache.json

    # 只想看某一種關係，或調整判定的嚴格程度：
    python3 video_edit_detect.py "資料夾" --dup-threshold 4 --sub-threshold 10 --min-cover 0.85

輸出：
    在指定資料夾旁產生 `影片重複與剪輯關係報告.xlsx`，包含四個分頁：
      摘要        —— 這次分析用的參數、檔案數、耗時，方便日後複查「當時是怎麼判定的」
      完全重複    —— 內容幾乎一樣的檔案配對
      子片段候選  —— A 是 B 裡面剪出來的一段（含時間偏移，方便直接跳到那一秒驗證）
      疑似合併群組—— 一部長片疑似由好幾部短片依序接成，並列出建議驗證方式

=== 如何人工複查（重要，先驗證再刪除）===
    報告裡每一列「子片段候選」都會附上「長片中的起始秒數」。用 QuickTime Player
    或 VLC 打開長片，直接拖到那個時間點，跟短片的開頭畫面/聲音對照，就能快速確認
    判定是否正確。不要只憑分數就刪檔案。

=== 與現有 VDF（Video Duplicate Finder + DINOv2 + Chromaprint）流程的關係 ===
    VDF 適合抓「整部影片彼此很像」的重複（近乎逐格比對），但它不會告訴你「這部
    10 分鐘短片其實是那部 2 小時長片的第 37 分到第 47 分」這種片段層級的對應關係，
    也不會自動驗證「這 12 部短片剛好依序接成這部長片」。這支腳本專門補這一塊，
    兩者可以互補使用：先用 VDF 抓出大範圍的相似群組，再用這支腳本針對同一群組
    算出精確的時間對應，方便人工複查。
"""

import argparse
import json
import os
import subprocess
import sys
import time
from itertools import combinations

import numpy as np

VIDEO_EXTS = {
    ".mp4", ".mov", ".m4v", ".avi", ".mkv", ".wmv", ".flv",
    ".mts", ".m2ts", ".webm", ".mpg", ".mpeg", ".3gp",
}

POPCOUNT_TABLE = np.array([bin(i).count("1") for i in range(256)], dtype=np.uint16)


# ---------------------------------------------------------------------------
# 檔案掃描與 ffprobe / fpcalc 呼叫
# ---------------------------------------------------------------------------

def find_video_files(root):
    files = []
    for dirpath, _, filenames in os.walk(root):
        for name in filenames:
            if name.startswith(".") or name.startswith("~$"):
                continue
            ext = os.path.splitext(name)[1].lower()
            if ext in VIDEO_EXTS:
                files.append(os.path.join(dirpath, name))
    return sorted(files)


def ffprobe_duration(path):
    """回傳影片長度（秒），任何容器格式都適用（透過 ffprobe 本身解析）。"""
    try:
        out = subprocess.run(
            ["ffprobe", "-v", "error", "-show_entries", "format=duration",
             "-of", "default=noprint_wrappers=1:nokey=1", path],
            capture_output=True, text=True, check=True, timeout=60,
        )
        return float(out.stdout.strip())
    except Exception as e:
        print(f"  [ffprobe 失敗] {os.path.basename(path)}: {e}", file=sys.stderr)
        return None


def fpcalc_raw_fingerprint(path):
    """
    用 chromaprint 的 fpcalc 取出完整長度的「原始」聲音指紋（-raw -length 0）。
    回傳 (numpy uint32 array, 指紋長度對應的秒數)；沒有聲軌或出錯時回傳 (None, None)。
    """
    try:
        out = subprocess.run(
            ["fpcalc", "-raw", "-length", "0", path],
            capture_output=True, text=True, check=True, timeout=300,
        )
    except Exception as e:
        print(f"  [fpcalc 失敗] {os.path.basename(path)}: {e}", file=sys.stderr)
        return None, None

    dur, fp = None, None
    for line in out.stdout.splitlines():
        if line.startswith("DURATION="):
            dur = float(line.split("=", 1)[1])
        elif line.startswith("FINGERPRINT="):
            fp = np.array([int(x) for x in line.split("=", 1)[1].split(",")], dtype=np.uint32)
    if fp is None or dur is None or len(fp) == 0:
        return None, None
    return fp, dur


# ---------------------------------------------------------------------------
# 快取（讓重跑時可跳過沒變動的檔案，結果可重現）
# ---------------------------------------------------------------------------

def load_cache(cache_path):
    if os.path.exists(cache_path):
        with open(cache_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_cache(cache_path, cache):
    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(cache, f)


def file_signature(path):
    st = os.stat(path)
    return {"mtime": st.st_mtime, "size": st.st_size}


def get_fingerprint(path, cache):
    sig = file_signature(path)
    key = path
    entry = cache.get(key)
    if entry and entry.get("sig") == sig and entry.get("fp") is not None:
        fp = np.array(entry["fp"], dtype=np.uint32)
        return fp, entry["duration"], entry.get("items_per_sec")

    duration = ffprobe_duration(path)
    fp, fp_duration = fpcalc_raw_fingerprint(path)
    if fp is None or duration is None:
        cache[key] = {"sig": sig, "fp": None, "duration": duration}
        return None, duration, None

    items_per_sec = len(fp) / fp_duration if fp_duration else None
    cache[key] = {
        "sig": sig,
        "fp": fp.tolist(),
        "duration": duration,
        "items_per_sec": items_per_sec,
    }
    return fp, duration, items_per_sec


# ---------------------------------------------------------------------------
# 指紋比對（Hamming distance，越低越像）
# ---------------------------------------------------------------------------

def popcount32(arr):
    b = arr.astype(np.uint32).view(np.uint8).reshape(-1, 4)
    return POPCOUNT_TABLE[b].sum(axis=1)


def mean_hamming_at_offset(long_fp, short_fp, offset):
    window = long_fp[offset:offset + len(short_fp)]
    if len(window) < len(short_fp):
        return None
    xor = np.bitwise_xor(window.astype(np.uint32), short_fp.astype(np.uint32))
    return float(popcount32(xor).mean())


def find_best_offset(long_fp, short_fp, coarse_step=8, refine_radius=None):
    """
    粗篩 + 細修兩階段搜尋：先用較大的步進掃過整個長片指紋序列，找出分數最低的
    大致位置，再在附近逐格微調找出最精確的偏移。回傳 (offset, mean_hamming分數)。
    分數是 0~32 的平均 Hamming distance，越接近 0 代表這段越吻合，
    隨機不相關的聲音大約會落在 14~18 之間。
    """
    refine_radius = refine_radius or coarse_step
    max_offset = len(long_fp) - len(short_fp)
    if max_offset < 0:
        return None, None

    best_offset, best_score = None, float("inf")
    for off in range(0, max_offset + 1, max(1, coarse_step)):
        d = mean_hamming_at_offset(long_fp, short_fp, off)
        if d is not None and d < best_score:
            best_score, best_offset = d, off

    lo = max(0, best_offset - coarse_step)
    hi = min(max_offset, best_offset + coarse_step)
    for off in range(lo, hi + 1):
        d = mean_hamming_at_offset(long_fp, short_fp, off)
        if d is not None and d < best_score:
            best_score, best_offset = d, off

    return best_offset, best_score


# ---------------------------------------------------------------------------
# 三種關係的判定邏輯
# ---------------------------------------------------------------------------

def analyze(files_info, dup_threshold, sub_threshold, min_cover, duration_tol):
    """
    files_info: list of dict，每個含 path, duration, fp, items_per_sec
    回傳 (exact_dupes, subclip_matches, merge_groups)
    """
    exact_dupes = []
    subclip_matches = []

    valid = [f for f in files_info if f["fp"] is not None]

    # --- 1. 完全重複：時長相近 + 整體 Hamming distance 低 ---
    for a, b in combinations(valid, 2):
        if a["duration"] is None or b["duration"] is None:
            continue
        dur_diff_ratio = abs(a["duration"] - b["duration"]) / max(a["duration"], b["duration"])
        if dur_diff_ratio > duration_tol:
            continue
        short_fp, long_fp = (a["fp"], b["fp"]) if len(a["fp"]) <= len(b["fp"]) else (b["fp"], a["fp"])
        offset, score = find_best_offset(long_fp, short_fp, coarse_step=4)
        if score is not None and score <= dup_threshold:
            exact_dupes.append({
                "file_a": a["path"], "file_b": b["path"],
                "duration_a": a["duration"], "duration_b": b["duration"],
                "hamming_score": round(score, 2), "offset_items": offset,
            })

    # --- 2. 子片段：短片是否為長片裡的一段 ---
    for short, long_ in combinations(valid, 2):
        if short["duration"] is None or long_["duration"] is None:
            continue
        if short["duration"] >= long_["duration"]:
            short, long_ = long_, short  # 確保 short 真的比較短
        if short["duration"] < 5:
            continue  # 太短的片段容易誤判，先略過
        if short["duration"] >= long_["duration"] * (1 - duration_tol):
            continue  # 長度太接近，已經在「完全重複」那邊處理過

        offset, score = find_best_offset(long_["fp"], short["fp"], coarse_step=8)
        if offset is None or score is None:
            continue
        if score <= sub_threshold:
            items_per_sec = long_["items_per_sec"] or (len(long_["fp"]) / long_["duration"])
            start_sec = offset / items_per_sec
            end_sec = start_sec + short["duration"]
            subclip_matches.append({
                "short_file": short["path"],
                "long_file": long_["path"],
                "short_duration": short["duration"],
                "long_duration": long_["duration"],
                "match_start_sec": round(start_sec, 1),
                "match_end_sec": round(end_sec, 1),
                "hamming_score": round(score, 2),
            })

    # --- 3. 疑似合併：同一部長片被多個子片段幾乎鋪滿整條時間軸 ---
    merge_groups = []
    by_long = {}
    for m in subclip_matches:
        by_long.setdefault(m["long_file"], []).append(m)

    for long_file, matches in by_long.items():
        long_duration = matches[0]["long_duration"]
        matches_sorted = sorted(matches, key=lambda m: m["match_start_sec"])
        covered = 0.0
        cursor = 0.0
        for m in matches_sorted:
            start, end = m["match_start_sec"], m["match_end_sec"]
            if end <= cursor:
                continue
            covered += end - max(start, cursor)
            cursor = max(cursor, end)
        coverage_ratio = covered / long_duration if long_duration else 0
        if len(matches_sorted) >= 2 and coverage_ratio >= min_cover:
            merge_groups.append({
                "long_file": long_file,
                "long_duration": long_duration,
                "coverage_ratio": round(coverage_ratio, 3),
                "pieces": [
                    {"short_file": m["short_file"], "start_sec": m["match_start_sec"],
                     "end_sec": m["match_end_sec"], "hamming_score": m["hamming_score"]}
                    for m in matches_sorted
                ],
            })

    return exact_dupes, subclip_matches, merge_groups


# ---------------------------------------------------------------------------
# Excel 報告輸出
# ---------------------------------------------------------------------------

def write_report(output_path, params, files_info, exact_dupes, subclip_matches, merge_groups):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    wb = Workbook()
    bold = Font(bold=True)
    red_fill = PatternFill("solid", fgColor="FFC7CE")
    yellow_fill = PatternFill("solid", fgColor="FFEB9C")
    orange_fill = PatternFill("solid", fgColor="FCD5B4")

    # --- 摘要 ---
    ws = wb.active
    ws.title = "摘要"
    ws.append(["產生時間", time.strftime("%Y-%m-%d %H:%M:%S")])
    ws.append(["掃描資料夾", params["root"]])
    ws.append(["掃描到的影片檔數", len(files_info)])
    ws.append(["成功取得聲音指紋的檔數", sum(1 for f in files_info if f["fp"] is not None)])
    ws.append(["取不到聲音指紋（可能無聲軌或壞檔）", ", ".join(
        os.path.basename(f["path"]) for f in files_info if f["fp"] is None) or "無"])
    ws.append([])
    ws.append(["判定參數（供日後複查／重現本次結果）"])
    for k, v in params.items():
        if k != "root":
            ws.append([k, v])
    ws.append([])
    ws.append(["找到：完全重複配對", len(exact_dupes)])
    ws.append(["找到：子片段候選", len(subclip_matches)])
    ws.append(["找到：疑似合併群組", len(merge_groups)])
    for row in ws.iter_rows(min_row=1, max_row=1):
        for c in row:
            c.font = bold

    # --- 完全重複 ---
    ws = wb.create_sheet("完全重複")
    ws.append(["檔案 A", "檔案 B", "A 長度(秒)", "B 長度(秒)", "相似分數(0-32，越低越像)"])
    for r in ws[1]:
        r.font = bold
        r.fill = red_fill
    for d in exact_dupes:
        ws.append([d["file_a"], d["file_b"], round(d["duration_a"], 1),
                   round(d["duration_b"], 1), d["hamming_score"]])

    # --- 子片段候選 ---
    ws = wb.create_sheet("子片段候選")
    ws.append(["短片檔案", "長片檔案", "短片長度(秒)", "長片長度(秒)",
               "在長片中的起始秒數", "在長片中的結束秒數", "相似分數(0-32，越低越像)",
               "複查方式"])
    for r in ws[1]:
        r.font = bold
        r.fill = yellow_fill
    for m in subclip_matches:
        ws.append([m["short_file"], m["long_file"], round(m["short_duration"], 1),
                   round(m["long_duration"], 1), m["match_start_sec"], m["match_end_sec"],
                   m["hamming_score"],
                   f"用 QuickTime/VLC 打開長片，跳到第 {m['match_start_sec']} 秒，跟短片開頭比對"])

    # --- 疑似合併群組 ---
    ws = wb.create_sheet("疑似合併群組")
    ws.append(["長片檔案", "長片長度(秒)", "覆蓋率", "組成片段（依時間順序）"])
    for r in ws[1]:
        r.font = bold
        r.fill = orange_fill
    for g in merge_groups:
        pieces_desc = "; ".join(
            f"{os.path.basename(p['short_file'])} [{p['start_sec']}s~{p['end_sec']}s]"
            for p in g["pieces"]
        )
        ws.append([g["long_file"], round(g["long_duration"], 1), g["coverage_ratio"], pieces_desc])

    for sheet in wb.worksheets:
        for col_cells in sheet.columns:
            length = max((len(str(c.value)) for c in col_cells if c.value is not None), default=10)
            sheet.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 12), 80)

    wb.save(output_path)


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="偵測影片資料夾內的重複／剪短／合併關係（格式無關）")
    parser.add_argument("folder", help="要掃描的資料夾（會自動遞迴掃描所有子資料夾）")
    parser.add_argument("--cache", default=".video_fp_cache.json",
                         help="指紋快取檔路徑（預設存在目前工作目錄，重跑時可跳過沒變動的檔案）")
    parser.add_argument("--output", default=None,
                         help="輸出報告路徑（預設在掃描資料夾旁產生 影片重複與剪輯關係報告.xlsx）")
    parser.add_argument("--dup-threshold", type=float, default=4.0,
                         help="判定為「完全重複」的 Hamming distance 門檻（0-32，預設 4，越小越嚴格）")
    parser.add_argument("--sub-threshold", type=float, default=8.0,
                         help="判定為「子片段」的 Hamming distance 門檻（0-32，預設 8）")
    parser.add_argument("--min-cover", type=float, default=0.85,
                         help="判定為「疑似合併」所需的時間軸覆蓋率（0-1，預設 0.85）")
    parser.add_argument("--duration-tol", type=float, default=0.02,
                         help="判定「時長相近」的容許誤差比例（預設 0.02，即 2%%）")
    args = parser.parse_args()

    root = os.path.abspath(args.folder)
    output_path = args.output or os.path.join(os.path.dirname(root), "影片重複與剪輯關係報告.xlsx")

    print(f"掃描資料夾：{root}")
    video_files = find_video_files(root)
    print(f"找到 {len(video_files)} 個影片檔（含所有子資料夾）\n")

    cache = load_cache(args.cache)
    files_info = []
    t0 = time.time()
    for i, path in enumerate(video_files, 1):
        rel = os.path.relpath(path, root)
        print(f"[{i}/{len(video_files)}] 取指紋中：{rel} ...", end="", flush=True)
        fp, duration, items_per_sec = get_fingerprint(path, cache)
        files_info.append({"path": path, "duration": duration, "fp": fp,
                            "items_per_sec": items_per_sec})
        print(" ✅" if fp is not None else " ⚠️ 無聲軌或失敗")
        if i % 20 == 0:
            save_cache(args.cache, cache)  # 中途存檔，避免大批次跑到一半中斷要重來
    save_cache(args.cache, cache)
    print(f"\n指紋取得完成，耗時 {time.time()-t0:.1f} 秒\n")

    print("開始比對關係（完全重複 / 子片段 / 疑似合併）...")
    exact_dupes, subclip_matches, merge_groups = analyze(
        files_info, args.dup_threshold, args.sub_threshold, args.min_cover, args.duration_tol
    )

    params = {
        "root": root,
        "dup_threshold": args.dup_threshold,
        "sub_threshold": args.sub_threshold,
        "min_cover": args.min_cover,
        "duration_tol": args.duration_tol,
    }
    write_report(output_path, params, files_info, exact_dupes, subclip_matches, merge_groups)

    print("\n" + "=" * 50)
    print(f"完成！報告已輸出：{output_path}")
    print(f"完全重複：{len(exact_dupes)} 組　子片段候選：{len(subclip_matches)} 組　"
          f"疑似合併群組：{len(merge_groups)} 組")
    print("請先用報告裡的時間偏移人工複查，確認無誤後才進行刪除或搬移。")
    print("=" * 50)


if __name__ == "__main__":
    main()
