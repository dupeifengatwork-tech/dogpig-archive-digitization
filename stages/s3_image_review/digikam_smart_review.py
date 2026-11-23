#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
digiKam 重複照片智能檢查清單產生器
====================================

不依賴 digiKam 自己的資料庫（之前踩過路徑格式不匹配、80-90% 區間比對出 0 組
的坑），改成直接掃描資料夾、自己算圖片感知雜湊（perceptual hash）來分組，
所以只要給資料夾路徑就能跑，跟 digiKam 內部怎麼存資料完全無關。

=== 判斷邏輯（依妳確認過的規則） ===

保留優先層級：
  1) 原始檔優先（不是複製後再編修的版本）
  2) 解析度較佳（寬 x 高）
  3) 檔案較大（清晰度）
  4) 檔名資訊較豐富

「原始檔」訊號怎麼判斷（分數 0-5，越高越像原始檔）：
  - EXIF 沒有 Software 標籤（代表沒被 Photoshop/Lightroom/Preview 等軟體動過）
  - 檔案修改時間 等於 拍攝時間（代表存檔後沒再被存過一次）
  - 檔名沒有 edited/copy/修改/編修/(1)/-2 這類編修痕跡

50-90% 區間的四條例外規則：
  a) 連續拍攝（EXIF 拍攝時間差在 --burst-seconds 秒內）→ 整組不建議刪除
  b) 判定為編修版的檔案，就算檔案比原始檔大，仍然優先保留原始檔
     （這條其實已經內含在「原始檔訊號」優先於「檔案大小」的排序裡）
  c) 黑白版 vs 彩色原始版 → 用 HSV 平均飽和度判斷，強制保留彩色版
  d) 修圖版檔名比原始檔名多出真正的資訊（不只是 edited/copy 這類編修字樣）
     → 兩者都保留，不列入刪除建議

=== 追蹤進度 ===
每次重新掃描都會用「檔案路徑」當唯一鍵，跟舊的 Excel 合併：已經標記過
狀態（已確認/已刪除/保留待查）的列會保留原狀態，只有新出現的檔案才會
是「尚未確認」，所以可以放心重複執行、不會每次都要整批重看。

=== 事前準備 ===
pip3 install --break-system-packages pillow imagehash openpyxl

=== 分組方式：預設「全部自動一起跑」，另外可疊加人眼確認過的結果 ===
（2026-09-23 改版：以前四種分組方式互斥、只能挑一種下指令，這是常見的
誤用來源——常常忘記另外再跑一次雜湊分組，導致 Excel 裡只有檔名家族分組
的結果。現在改成：只要沒特別關掉，下面 1)、2) 兩種自動分組每次都會一起
跑、結果自動合併去重，不用再分好幾次下指令。）

自動一起跑的兩種（預設一定會跑）：
  1) 檔名家族分組   同資料夾 + 檔名核心編號相同（例如 034.JPG /
                    034(高斯效果).jpg 這種），最快、最準，但只在檔名本來
                    就有規律時才抓得到。
  2) 內建感知雜湊分組  對所有圖片兩兩比對雜湊距離，分成「100%完全相同」
                    跟「50-90%近似」兩種標籤（用 --exact-max-dist /
                    --similar-max-dist 調整門檻）。注意：對「顏色/色調被
                    大幅調過」或「轉黑白」的版本辨識力很弱，這種情況
                    digiKam 自己的演算法也會漏，需要用下面的人工輸入補。
  想跳過雜湊比對（例如資料夾圖片數量很多、只想先看檔名家族分組結果）可以
  加 --skip-hash。內建雜湊分組預設只在「同一個資料夾內」兩兩比對（活動辦
  在不同資料夾，重複照片幾乎不會跨資料夾出現，這樣能大幅減少比對量、避免
  資料夾/圖片數量一多就跑很久）；真的需要跨資料夾找重複時可以加
  --cross-folder 改成整批全部兩兩比對，但要有心理準備會明顯變慢。

可疊加的人工輸入（來自 digiKam 肉眼複查後手動整理，選配、可與上面同時給）：
  3) --groups-csv   吃外部已經分好的組（group_id,file_path 兩欄），
                    例如你在 digiKam 畫面上人眼看出來該歸成一組、
                    但 digiKam 自己沒抓對時，手動整理成 CSV 餵給程式。
                    同一組內重複列出的路徑會自動去重。
  4) --pairs-csv    吃「兩兩很像」的配對清單（file_a,file_b 兩欄），
                    程式會自動用 Union-Find 把有遞移關係的配對合併成同一組。
                    這是專門解決 digiKam 把同一批檔案拆成好幾組顯示的問題：
                    妳不用自己先想「這些總共該分成幾組」，看到哪兩張像
                    就記一行，程式會自己接起來。

所有來源（檔名家族＋雜湊＋groups-csv＋pairs-csv）會合併進同一個 Union-Find：
同一張照片不管被哪幾種方式抓到都只會在 Excel 出現一次，「相似度分級」欄
會列出這組是被哪些方式抓到的（例如「檔名家族分組、100%（完全相同）」）。

=== 操作步驟（建議順序）===
1. 直接跑（不用加任何分組參數），檔名家族分組＋內建雜湊分組會自動一起跑，
   一次產生 Excel：
   python3 digikam_smart_review.py "資料夾路徑" --output 查重清單.xlsx

2. 再跑 digiKam 內建的相似度搜尋（100-100% → 80-90% → 50-80%），把畫面上
   看到的參考組人眼複查一次；如果 digiKam 抓到、但上面 Excel 沒有的組
   （通常是檔名沒規律、色調被大幅調整、或彩轉黑白的版本——這是內建雜湊
   分組已知辨識力較弱的地方），把妳看出來該歸成一組的檔案整理成
   --pairs-csv（兩兩配對即可，不用想要分幾組）或 --groups-csv（已經分好
   組的話），再跑一次：
   python3 digikam_smart_review.py "資料夾路徑" --output 查重清單.xlsx --pairs-csv 配對.csv
   （這次還是會照常自動跑檔名家族分組＋內建雜湊分組，--pairs-csv 是疊加
   進去的額外來源，不是取代。）

3. 每次都輸出到「同一個」--output 路徑，程式會用檔案路徑合併新舊結果，
   已經在 Excel 標過「已確認/已刪除」的列不會被洗掉，只有新出現的檔案
   會是「尚未確認」。

4. 打開 Excel，只需要專心看「建議動作＝待複核」的列（其餘 保留/刪除/
   連拍不建議刪除/兩者都保留 是程式判斷完的建議，仍建議抽查但不必逐列看）。

5.（可選）加 --write-labels，把建議動作同步寫成 digiKam 的色彩標籤／
   揀選標籤（XMP），這樣就能直接在 digiKam 縮圖畫面用篩選器看建議結果，
   不用切去看 Excel。細節見下方「digiKam 標籤整合」段落。

參數：
  folder                要掃描的資料夾（自動遞迴掃描子資料夾）
  --output               輸出的 Excel 路徑；若已存在會合併狀態欄，不覆蓋
  --skip-hash             跳過內建感知雜湊分組，只做檔名家族分組（＋人工輸入，
                          如果有給的話）；圖片數量非常多、雜湊兩兩比對太慢
                          時可以用這個先出一版結果
  --hash-size            感知雜湊精細度，預設 16（越大越精細，速度略慢）
  --exact-max-dist       「100%完全相同」分級的雜湊距離上限，預設 3
  --similar-max-dist     「50-90%近似」分級的雜湊距離上限，預設 14
  --burst-seconds        判定為連續拍攝的拍攝時間差，預設 10 秒
  --grayscale-sat        判定為黑白的平均飽和度上限（0-255），預設 18
  --near-tie-pct          解析度/檔案大小差距在這個百分比以內視為打平、
                          轉人工複核，避免 0.6% 這種雜訊等級差距就自動選邊，
                          預設 3.0
  --no-thumbnails        不嵌入縮圖（檔案很多、想加快產生速度時使用）
  --groups-csv            疊加外部分組結果（group_id,file_path），跟自動
                          分組的結果合併，不是取代
  --pairs-csv             疊加外部兩兩配對清單（file_a,file_b），自動
                          Union-Find 合併成組，跟自動分組的結果合併
  --write-labels          把本次建議動作寫成 digiKam 色彩/揀選標籤（XMP），
                          需要系統已安裝 exiftool（brew install exiftool）
  --family, --tier        舊版參數，保留相容但已不需要使用：檔名家族分組
                          跟內建雜湊分組現在預設就會一起跑，不用再另外指定

=== 這份腳本「準」在哪裡、可能「漏」在哪裡 ===
準確度來源（可信度由高到低）：
  - --pairs-csv / --groups-csv：完全來自妳人眼確認的分組，程式只負責套用
    決策規則跟合併去重，分組本身的正確性 100% 取決於妳給的資料，不是程式判斷的。
  - 檔名家族分組：純文字規則（去括號、去空白、去「拷貝/副本/copy」字樣），
    不看圖片內容，所以不會誤判「長得像但其實不同」的照片，但也只在
    檔名真的有規律時才抓得到，抓不到不會報錯，只是那組不會出現在清單裡。
  - 原始檔訊號分數（EXIF Software／修改時間＝拍攝時間／檔名無編修字樣）：
    三個獨立訊號取分數加總，任一個 EXIF 被清除（例如網路下載、截圖）都會
    讓對應那一分消失，分數會變低但不會判斷錯方向；三個訊號同時消失時
    （EXIF 全無），才會退到只比解析度/檔案大小。
  - 黑白/彩色判斷（HSV 平均飽和度）：對「明顯黑白」跟「明顯彩色」很準，
    但對「刻意做舊、低飽和度調色」這種灰階邊界的圖，門檻（--grayscale-sat）
    可能需要依實際狀況調整。
  - 近似容忍度（--near-tie-pct）：這是刻意做的「不確定就不要裝確定」機制，
    差距在容許範圍內就換下一層比較，全部都打平就直接轉人工複核，不會
    因為 0.6% 這種雜訊硬選一個。
  - 內建感知雜湊分組：對「一般的重複/裁切/縮圖」判斷還算穩，但對「顏色/
    色調大幅調整」或「彩轉黑白」辨識力很弱（同源測試中量到黑白版跟原始
    彩色版的雜湊距離，比兩張本來就不同的照片還遠）——這一點連 digiKam
    自己的演算法也有同樣的問題（已在妳提供的黑白版漏抓案例中驗證過），
    不是這支腳本特有的缺陷，而是感知雜湊這個技術本身的局限。

需要人眼複核、程式無法/不建議自動判斷的狀況：
  - 「建議動作＝待複核」的所有列（所有比較層都打平，或缺 EXIF 又解析度/
    檔案大小都很接近）。
  - 檔名只有純數字尾碼（01/02/03，沒有拷貝/副本/copy 這類字樣）：刻意
    不會被檔名家族分組自動合併，因為這種編號有時代表真的不同張照片
    （例如海報01/02 案例），有時又是同一張的不同檔（例如烏克麗麗01/02/03
    案例），無法只憑檔名規則安全判斷，需要人眼看內容決定要不要用
    --pairs-csv 手動歸組。
  - 顏色/色調被大幅調整、或彩轉黑白的版本：內建雜湊分組可能連不起來，
    需要人眼發現後用 --pairs-csv 或 --groups-csv 補進去。
  - EXIF 完全被清除的圖片（網路下載、截圖、社群平台重新編碼過的檔）：
    原始檔訊號分數會失真，決策會退到只看解析度/檔案大小，這種情況下
    「解析度/檔案大小最大 ≠ 一定是原始檔」，建議對這類組別提高警覺、
    多花一點時間複核。
  - 這支腳本只會產生「建議」，不會刪除任何檔案；決定刪除前一定要自己
    看過 Excel 的判斷依據欄。

=== digiKam 標籤整合（--write-labels，選配）===
背景：digiKam 本身沒有「顯示文字判斷依據」的簡單擴充管道，但它原生支援
色彩標籤（Color Label）跟揀選標籤（Pick Label），並且會讀寫在檔案的 XMP
metadata 裡（前提是 digiKam 設定裡有開「Settings > Configure digiKam >
Metadata > 也把中繼資料寫入圖片檔案 (sidecar)」，這是 digiKam 常見的預設
開啟項，但建議先自己確認一次）。--write-labels 讓這支腳本直接用 exiftool
把本次判斷結果寫進這兩個欄位，不用寫、編譯任何 digiKam 外掛。

對照表（建議動作 → 色彩標籤 / 揀選標籤）：
  ✅ 建議保留              → 綠色 / 已接受(Accepted)
  🟡 資訊不同，兩者皆留     → 黃色 / 已接受(Accepted)
  📸 連續拍攝，不建議刪除   → 藍色 / 已接受(Accepted)
  ❌ 建議刪除              → 紅色 / 已拒絕(Rejected)
  ⚠️ 信心不足，建議人工複核 → 灰色 / 待定(Pending)

這個對照表用的是 digiKam 色彩/揀選標籤長年沒變過的固定數值順序（見程式碼
COLOR_LABEL_* / PICK_LABEL_* 常數），但這不是 digiKam 官方公開文件保證的
API，第一次使用建議先驗證一次：在 digiKam 裡手動把一張測試照片點成紅色
標籤、存檔，再對它跑
  exiftool -XMP-digiKam:ColorLabel 測試照片路徑
確認讀出來是 1，之後再放心對整批資料夾跑 --write-labels。

寫入後的步驟（這是 digiKam 的正常機制，腳本無法代勞）：
  1. 在 digiKam 裡對該資料夾按右鍵 →「重新讀取中繼資料」
     （或 Tools > Maintenance > Sync Metadata），digiKam 才會把檔案裡的
     標籤讀回資料庫、顯示在縮圖畫面上。
  2. 用「篩選 > 色彩標籤」或「篩選 > 揀選標籤」就能直接看到建議結果。

注意事項：
  - 這個功能會覆蓋檔案「目前」的色彩/揀選標籤，包括妳在 digiKam 裡手動
    調整過的——如果某個資料夾裡有不想被腳本蓋掉的手動標籤，先不要對那個
    資料夾跑 --write-labels。
  - 需要系統已安裝 exiftool：macOS 上用 `brew install exiftool`。
  - 跟 Excel 一樣，這仍然只是「建議」，最終要不要刪除還是妳自己在
    digiKam 或檔案系統裡動手。

=== 可重複驗證的基準 ===
隨腳本附上 test_digikam_smart_review.py，是不依賴妳實際照片、完全自產
測試資料（假特徵字典 + PIL 合成帶 EXIF 的圖片）的回歸測試，涵蓋這次
對話中驗證過的每一種真實案例（原始檔 vs 編修版、黑白 vs 彩色、連拍、
同拍攝時間不算連拍、EXIF 全無時的近似容忍度、完全打平轉人工複核、
兩種檔名家族分組、刻意不合併純數字尾碼、pairs-csv 跨組合併、
groups-csv 重複列去重、檔名多資訊兩者都保留、EXIF 讀取正確性、
建議動作 → digiKam 標籤的對照表、多種分組來源合併去重不重複列出）。
之後如果改規則、調參數，跑一次這個測試檔案可以立刻知道有沒有改壞
既有邏輯：
  python3 test_digikam_smart_review.py
全部通過會印出「全部通過 ✅」，任何一項壞掉會列出壞掉的項目名稱並
以非 0 狀態碼結束。
"""

import argparse
import hashlib
import os
import re
import shutil
import subprocess
import sys
import tempfile
from collections import defaultdict
from datetime import datetime

from PIL import Image, ExifTags

try:
    import imagehash
except ImportError:
    print("缺少 imagehash，請先執行：pip3 install --break-system-packages imagehash")
    sys.exit(1)

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.drawing.image import Image as XLImage
except ImportError:
    print("缺少 openpyxl，請先執行：pip3 install --break-system-packages openpyxl")
    sys.exit(1)


IMAGE_EXTS = (".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp")

TIER_LABEL_FAMILY = "檔名家族分組"
TIER_LABEL_EXACT = "100%（完全相同）"
TIER_LABEL_SIMILAR = "50-90%（近似）"
TIER_LABEL_GROUPS_CSV = "外部分組（digiKam 匯出）"
TIER_LABEL_PAIRS_CSV = "外部配對（digiKam 匯出）"

# 檔名中常見的「編修痕跡」字樣（出現這些字樣不算是檔名資訊變豐富，只是編修標記）
EDIT_MARKERS = [
    "edited", "edit", "copy", "副本", "修改", "編修", "修圖", "調色",
    "-2", "_2", "(1)", " 1)", "v2", "final", "fixed",
]
EDIT_MARKER_RE = re.compile(
    "|".join(re.escape(m) for m in EDIT_MARKERS), re.IGNORECASE
)

EDIT_SOFTWARE_HINTS = [
    "photoshop", "lightroom", "capture one", "preview", "gimp",
    "snapseed", "vsco", "美图", "剪映", "faststone", "picasa",
    "affinity", "luminar",
]

STATUS_UNCONFIRMED = "尚未確認"
STATUS_OK_COLORS = {
    "尚未確認": "FFFFFF",
    "已確認": "D9EAD3",
    "已刪除": "F4CCCC",
    "保留待查": "FFF2CC",
}

ACTION_KEEP = "✅ 建議保留"
ACTION_DELETE = "❌ 建議刪除"
ACTION_KEEP_BOTH = "🟡 資訊不同，兩者皆留"
ACTION_BURST = "📸 連續拍攝，不建議刪除"
ACTION_REVIEW = "⚠️ 信心不足，建議人工複核"

ACTION_FILL = {
    ACTION_KEEP: "C6E0B4",
    ACTION_DELETE: "F4CCCC",
    ACTION_KEEP_BOTH: "FFE599",
    ACTION_BURST: "CFE2F3",
    ACTION_REVIEW: "FFF2CC",
}


# ---------------------------------------------------------------------------
# digiKam 色彩／揀選標籤對照（--write-labels 用）
# ---------------------------------------------------------------------------
# 對照來源：digiKam 色彩/揀選標籤長年沒變過的固定數值順序（非官方公開文件
# 保證的 API）。第一次使用前建議先驗證一次：在 digiKam 手動把一張測試照片
# 點成紅色標籤、存檔，跑 `exiftool -XMP-digiKam:ColorLabel 該檔案` 確認讀
# 出來是 1，再放心對整批資料夾跑 --write-labels。
COLOR_LABEL_NONE = 0
COLOR_LABEL_RED = 1
COLOR_LABEL_ORANGE = 2
COLOR_LABEL_YELLOW = 3
COLOR_LABEL_GREEN = 4
COLOR_LABEL_BLUE = 5
COLOR_LABEL_MAGENTA = 6
COLOR_LABEL_GRAY = 7
COLOR_LABEL_BLACK = 8

PICK_LABEL_NONE = 0
PICK_LABEL_REJECTED = 1
PICK_LABEL_PENDING = 2
PICK_LABEL_ACCEPTED = 3

# 建議動作 → (色彩標籤, 揀選標籤)。色彩標籤負責讓妳一眼看出「為什麼」，
# 揀選標籤負責最重要的「該不該刪」，兩者互為備援：只用其中一種篩選也夠用。
ACTION_TO_LABELS = {
    ACTION_KEEP: (COLOR_LABEL_GREEN, PICK_LABEL_ACCEPTED),
    ACTION_KEEP_BOTH: (COLOR_LABEL_YELLOW, PICK_LABEL_ACCEPTED),
    ACTION_BURST: (COLOR_LABEL_BLUE, PICK_LABEL_ACCEPTED),
    ACTION_DELETE: (COLOR_LABEL_RED, PICK_LABEL_REJECTED),
    ACTION_REVIEW: (COLOR_LABEL_GRAY, PICK_LABEL_PENDING),
}


def action_to_labels(action):
    """回傳 (色彩標籤整數, 揀選標籤整數)；未知動作一律回傳 (無標籤, 待定)，
    保守處理、不亂猜一個看起來像「保留」或「刪除」的標籤。"""
    return ACTION_TO_LABELS.get(action, (COLOR_LABEL_NONE, PICK_LABEL_PENDING))


def build_exiftool_args(path, color_label, pick_label):
    """組出寫入 digiKam 色彩/揀選標籤用的 exiftool 參數列表（不含 exiftool
    執行檔本身）。拆成獨立函式方便測試：不用真的安裝 exiftool、不用真的
    寫檔案，就能驗證參數組得對不對。"""
    return [
        f"-XMP-digiKam:ColorLabel={color_label}",
        f"-XMP-digiKam:PickLabel={pick_label}",
        "-overwrite_original",
        "-P",  # 保留檔案原本的作業系統修改時間，避免寫 metadata 動到它
        path,
    ]


def write_labels_to_files(rows):
    """把每一列的建議動作寫成 digiKam 色彩/揀選標籤（XMP），需要系統已安裝
    exiftool。寫入之後要在 digiKam 裡對該資料夾「重新讀取中繼資料」，
    digiKam 才會把標籤讀回資料庫、顯示在縮圖畫面上——這是 digiKam 的正常
    機制，這支腳本無法代勞。

    注意：這個功能會覆蓋檔案「目前」的色彩/揀選標籤，包括妳在 digiKam 裡
    手動調整過的；如果某個資料夾有不想被覆蓋的手動標籤，不要對它跑
    --write-labels。
    """
    if shutil.which("exiftool") is None:
        print("\n⚠️  找不到 exiftool，無法寫入色彩/揀選標籤。")
        print("    macOS 安裝方式：brew install exiftool")
        return

    ok, failed = 0, []
    for row in rows:
        path = row.get("path")
        if not path or not os.path.exists(path):
            continue  # 舊 Excel 帶進來、實際檔案已經不在的列，略過
        color_label, pick_label = action_to_labels(row["action"])
        args = build_exiftool_args(path, color_label, pick_label)
        try:
            result = subprocess.run(
                ["exiftool"] + args,
                capture_output=True, text=True, timeout=30,
            )
            if result.returncode == 0:
                ok += 1
            else:
                failed.append((path, result.stderr.strip()))
        except Exception as e:
            failed.append((path, str(e)))

    print(f"\n色彩/揀選標籤寫入完成：成功 {ok} 個"
          + (f"、失敗 {len(failed)} 個" if failed else ""))
    if failed:
        print("失敗清單（最多列前 10 筆）：")
        for path, err in failed[:10]:
            print(f"  - {path}: {err}")
    print(
        "\n下一步：在 digiKam 裡對這個資料夾按右鍵 →「重新讀取中繼資料」\n"
        "（或 Tools > Maintenance > Sync Metadata），標籤才會顯示在縮圖畫面上。\n"
        "之後可以用「篩選 > 色彩標籤」或「篩選 > 揀選標籤」快速看出建議結果：\n"
        "🟢綠=建議保留　🟡黃=兩者皆留　🔵藍=連拍不建議刪除　🔴紅=建議刪除　"
        "⚪灰=待複核（揀選標籤：已接受=不刪、已拒絕=建議刪除、待定=待複核）。"
    )


# ---------------------------------------------------------------------------
# 第一步：掃描資料夾，抽取每張圖的特徵
# ---------------------------------------------------------------------------

def find_images(root_folder):
    for dirpath, _, filenames in os.walk(root_folder):
        for name in filenames:
            if name.startswith(".") or name.startswith("~$"):
                continue
            if os.path.splitext(name)[1].lower() in IMAGE_EXTS:
                yield os.path.join(dirpath, name)


def read_exif(im):
    try:
        raw = im._getexif() or {}
    except Exception:
        return {}
    out = {}
    for k, v in raw.items():
        tag = ExifTags.TAGS.get(k, k)
        out[tag] = v
    return out


def mean_saturation(im):
    """回傳 0-255 的平均飽和度；越接近 0 越接近黑白。"""
    try:
        hsv = im.convert("HSV")
        hist = hsv.getchannel("S").histogram()
        total = sum(hist)
        if total == 0:
            return 0.0
        return sum(i * c for i, c in enumerate(hist)) / total
    except Exception:
        return None


def parse_exif_datetime(value):
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y:%m:%d %H:%M:%S")
    except (ValueError, TypeError):
        return None


def extract_feature(filepath, hash_size):
    try:
        with Image.open(filepath) as im:
            im.load()
            width, height = im.size
            phash = imagehash.phash(im, hash_size=hash_size)
            sat = mean_saturation(im)
            exif = read_exif(im)
    except Exception as e:
        return {"path": filepath, "error": str(e)}

    software = (exif.get("Software") or "").strip()
    dt_original = parse_exif_datetime(exif.get("DateTimeOriginal"))
    dt_modified = parse_exif_datetime(exif.get("DateTime"))

    return {
        "path": filepath,
        "filename": os.path.basename(filepath),
        "size_bytes": os.path.getsize(filepath),
        "width": width,
        "height": height,
        "megapixels": round(width * height / 1e6, 2),
        "phash": phash,
        "saturation": sat,
        "software": software,
        "dt_original": dt_original,
        "dt_modified": dt_modified,
        "error": None,
    }


# ---------------------------------------------------------------------------
# 第二步：分組（多種來源，用 Union-Find 合併成最終組別）
# ---------------------------------------------------------------------------

class UnionFind:
    def __init__(self, items):
        self.parent = {i: i for i in items}

    def find(self, x):
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.parent[ra] = rb


def group_by_csv(features, csv_path):
    """從外部 CSV 讀取分組（欄位：group_id,file_path）。

    用途：如果妳已經有從 digiKam 資料庫（或 digiKam 的「尋找重複項」搜尋結果）
    匯出的分組清單，直接用那份分組結果最準確——本程式自己算的感知雜湊
    （phash）對「大幅調色/濾鏡處理過」的編修版本不夠敏感，可能抓不到
    原始檔跟編修版是同一組（實測發現：同一張照片，重度調色後 phash 距離
    可以遠到不會被判定為相似）。有 digiKam 自己分好的組最可靠，本程式只
    負責在組內判斷「哪個該留、哪個該刪」。
    """
    import csv as csv_module
    by_path = {f["path"]: f for f in features}
    groups = defaultdict(dict)  # group_id -> {path: feature}，用 dict 天然去重
    with open(csv_path, encoding="utf-8-sig", newline="") as fh:
        reader = csv_module.DictReader(fh)
        for row in reader:
            path = row["file_path"].strip()
            if path in by_path:
                # 用 path 當 key 自動去重：digiKam 匯出的清單有時會把同一個檔案
                # 在同一組裡重複列出兩次（實測遇過），這裡確保每個檔案在同一組
                # 只算一次，不會在 Excel 裡出現一模一樣的兩列。
                groups[row["group_id"].strip()][path] = by_path[path]
            else:
                print(f"  ⚠️ CSV 裡的路徑掃描時沒找到，略過：{path}")
    return [list(g.values()) for g in groups.values() if len(g) > 1]


FAMILY_BRACKET_RE = re.compile(r"[\(（\[【][^\)）\]】]*[\)）\]】]")
FAMILY_WS_RE = re.compile(r"[\s　_\-]+")
FAMILY_MIN_KEY_LEN = 4

# 檔名「尾巴」常見的複製/副本標記（不一定被括號包住），例如：
#   Chulki tour dogpig拷貝.jpg / Chulki tour dogpig拷貝2.jpg / Chulki-tour-dogpig.jpg
# 這三個都是同一份檔案重複「另存新檔」出來的，拿掉尾巴的「拷貝／拷貝2」才會
# 現出同一個核心名稱。注意：這只處理「複製」語意的字樣，不處理 01/02 這種
# 序號——序號常常代表「真的不同兩張」（跟連拍規則一樣的道理），不能自動當
# 成同一組，避免把本來不同的檔案誤判成重複。
FAMILY_TRAILING_COPY_RE = re.compile(
    r"(拷貝\d*|副本\d*|複製\d*|copy\s*\d*|copie\d*)$", re.IGNORECASE
)


def normalize_family_key(path):
    """把檔名去掉括號註記（如「(高斯效果)」「(高斯效果再處理)」）、空白/底線/
    連字號、大小寫差異、以及尾巴的「拷貝/副本/copy」標記，再去掉開頭跟所在
    資料夾同名的重複前綴（有些檔案會把活動名稱重複打在檔名裡，有些又省略），
    剩下的核心名稱就是拿來判斷「是不是同一份檔案的不同存檔／版本」的依據。

    例 1：
      200804膠帶畫廊 20080000 034(高斯效果再處理).jpg
      200804膠帶畫廊20080000 034.JPG
      20080000 034(高斯效果).jpg
    三個都會變成同一個核心編號「20080000034」。

    例 2：
      Chulki tour dogpig拷貝.jpg
      Chulki tour dogpig拷貝2.jpg
      Chulki-tour-dogpig.jpg
    三個都會變成同一個核心名稱「chulkitourdogpig」，即使 digiKam 在 80-90%
    這個區間只抓到其中一部分（實測發現「拷貝2」完全沒被抓進任何一組）。

    「惠笙20100307表演-海報01.jpg」跟「...-海報02.jpg」這種只有尾碼數字不同、
    沒有複製語意字樣的檔名，這個函式刻意不會把它們判成同一組——01/02 常常是
    真的不同兩張（例如兩頁海報、兩個版本設計），要靠視覺相似度而不是檔名去
    判斷，交給檔名家族分組以外的分組方式（digiKam 分組 CSV 或內建雜湊）處理。
    """
    parent = os.path.basename(os.path.dirname(path))
    stem = os.path.splitext(os.path.basename(path))[0]
    stem_clean = FAMILY_BRACKET_RE.sub("", stem)
    stem_norm = FAMILY_WS_RE.sub("", stem_clean).lower()
    parent_norm = FAMILY_WS_RE.sub("", parent).lower()
    if parent_norm and stem_norm.startswith(parent_norm):
        stem_norm = stem_norm[len(parent_norm):]
    prev = None
    while prev != stem_norm:
        prev = stem_norm
        stem_norm = FAMILY_TRAILING_COPY_RE.sub("", stem_norm)
    return stem_norm


def group_by_filename_family(features):
    """依「同資料夾 + 檔名核心編號相同」分組，完全不需要 digiKam 或相似度演算法。

    適合妳的檔案庫這種習慣：同一張照片的不同版本會用「原編號 + 括號註記效果」
    來命名（如 034.JPG / 034(高斯效果).jpg / 034(高斯效果再處理).jpg）。這種
    情況下用檔名分組比相似度分組更準，也不會因為 digiKam 把同一組拆到不同
    相似度區間、要分三次跑才找得齊。
    """
    groups = defaultdict(list)
    for f in features:
        key = normalize_family_key(f["path"])
        if len(key) < FAMILY_MIN_KEY_LEN:
            continue  # 核心編號太短，太容易誤判成同組，跳過不分組
        dir_key = os.path.dirname(f["path"])
        groups[(dir_key, key)].append(f)
    return [g for g in groups.values() if len(g) > 1]


def group_by_pairs_csv(features, csv_path):
    """從外部「兩兩配對」CSV 讀取分組（欄位：file_a,file_b），用 Union-Find
    自動把有遞移關係的配對合併成同一組。

    這是專門解決「digiKam 自己在畫面上把同一批重複檔案拆成好幾組」的問題：
    digiKam 的重複項清單是以「參考圖片」為中心兩兩配對顯示的，A 可能只跟 B
    配對顯示在一組、C 只跟 D 配對顯示在另一組，但如果 B 跟 C 其實也是同一張
    （只是 digiKam 沒把它們放在同一個參考組裡），這幾組其實應該是同一組。
    只要把 digiKam 畫面上看到的「這兩個很像」逐一記成一行 file_a,file_b
    （不用先手動判斷該分成幾組），本程式會自動用 Union-Find 算出正確的組別
    ——A-B、B-C、C-D 只要輾轉相連，就會被合併成一組 {A,B,C,D}，不會再各自
    散落在原本 digiKam 顯示的那幾個「參考組」裡。
    """
    import csv as csv_module
    by_path = {f["path"]: f for f in features}
    all_paths_in_csv = set()
    with open(csv_path, encoding="utf-8-sig", newline="") as fh:
        reader = csv_module.DictReader(fh)
        rows = list(reader)
    for row in rows:
        all_paths_in_csv.add(row["file_a"].strip())
        all_paths_in_csv.add(row["file_b"].strip())
    known_paths = [p for p in all_paths_in_csv if p in by_path]
    missing = all_paths_in_csv - set(known_paths)
    for p in missing:
        print(f"  ⚠️ 配對 CSV 裡的路徑掃描時沒找到，略過：{p}")

    uf = UnionFind(known_paths)
    for row in rows:
        a, b = row["file_a"].strip(), row["file_b"].strip()
        if a in by_path and b in by_path:
            uf.union(a, b)

    groups = defaultdict(list)
    for p in known_paths:
        groups[uf.find(p)].append(by_path[p])
    return [g for g in groups.values() if len(g) > 1]


def group_by_hash(features, max_dist):
    """單一距離門檻的內建感知雜湊分組（提供給測試/獨立使用）。"""
    valid = [f for f in features if f.get("phash") is not None]
    uf = UnionFind([f["path"] for f in valid])
    n = len(valid)
    for i in range(n):
        for j in range(i + 1, n):
            dist = valid[i]["phash"] - valid[j]["phash"]
            if dist <= max_dist:
                uf.union(valid[i]["path"], valid[j]["path"])

    groups = defaultdict(list)
    for f in valid:
        groups[uf.find(f["path"])].append(f)

    return [g for g in groups.values() if len(g) > 1]


# ---------------------------------------------------------------------------
# 多來源合併：把「檔名家族／內建雜湊／groups-csv／pairs-csv」的分組結果
# 全部倒進同一個 Union-Find，同一張照片不管被幾種方式抓到都只算一次。
# ---------------------------------------------------------------------------

def build_family_edges(features):
    """把檔名家族分組轉成邊清單：組內全部連到第一個成員（星狀連法），
    足夠讓 Union-Find 把整組連通，不需要兩兩全連。"""
    edges = []
    for group in group_by_filename_family(features):
        anchor = group[0]["path"]
        for f in group[1:]:
            edges.append((anchor, f["path"], TIER_LABEL_FAMILY))
    return edges


def build_hash_edges(features, exact_max_dist, similar_max_dist, cross_folder=False):
    """計算圖片兩兩之間的雜湊距離，依距離分成「完全相同」跟「近似」兩種標籤
    的邊（只算一次距離，不用像以前一樣為兩個 tier 各跑一次 O(n²)）。

    預設只在「同一個資料夾內」兩兩比對（cross_folder=False）：豆皮的檔案庫
    是按活動分資料夾（一個講座/派對一個資料夾），實務上重複照片幾乎都出現
    在同一場活動的資料夾裡，不會跨兩個不相干的活動誤判成重複。這樣可以把
    比對量從「整批全部兩兩比對」（例如 2000 張圖片 ≈ 200 萬組比對）降到
    「各資料夾內部比對後加總」（例如 100 個資料夾、每個 20 張 ≈ 2 萬組），
    對這種資料夾數量多、單一資料夾內圖片數量少的檔案庫是關鍵的效能改善——
    不這樣做的話，一個資料夾有兩千張圖片的規模，兩兩比對可能要跑好幾分鐘
    甚至更久。真的需要跨資料夾找重複（例如懷疑同一張照片被放進兩個不同
    活動的資料夾）時，用 --cross-folder 開啟全域比對，但資料夾/圖片數量多
    時要有心理準備會跑比較久。
    """
    valid = [f for f in features if f.get("phash") is not None]
    edges = []

    if cross_folder:
        buckets = {None: valid}
    else:
        buckets = defaultdict(list)
        for f in valid:
            buckets[os.path.dirname(f["path"])].append(f)

    for bucket in buckets.values():
        n = len(bucket)
        for i in range(n):
            for j in range(i + 1, n):
                dist = bucket[i]["phash"] - bucket[j]["phash"]
                if dist <= exact_max_dist:
                    edges.append((bucket[i]["path"], bucket[j]["path"], TIER_LABEL_EXACT))
                elif dist <= similar_max_dist:
                    edges.append((bucket[i]["path"], bucket[j]["path"], TIER_LABEL_SIMILAR))
    return edges


def build_groups_csv_edges(features, csv_path):
    edges = []
    for group in group_by_csv(features, csv_path):
        anchor = group[0]["path"]
        for f in group[1:]:
            edges.append((anchor, f["path"], TIER_LABEL_GROUPS_CSV))
    return edges


def build_pairs_csv_edges(features, csv_path):
    """跟 group_by_pairs_csv 讀同一份 CSV，但直接回傳每一行原始配對當一條邊
    （不用先跑一次 Union-Find 再攤平成星狀邊，比較單純）。"""
    import csv as csv_module
    by_path = {f["path"] for f in features}
    edges = []
    with open(csv_path, encoding="utf-8-sig", newline="") as fh:
        reader = csv_module.DictReader(fh)
        for row in reader:
            a, b = row["file_a"].strip(), row["file_b"].strip()
            if a in by_path and b in by_path:
                edges.append((a, b, TIER_LABEL_PAIRS_CSV))
            else:
                for p in (a, b):
                    if p not in by_path:
                        print(f"  ⚠️ 配對 CSV 裡的路徑掃描時沒找到，略過：{p}")
    return edges


def combine_all_groups(features, skip_hash=False, exact_max_dist=3,
                        similar_max_dist=14, groups_csv=None, pairs_csv=None,
                        cross_folder=False):
    """把所有分組來源合併成最終組別清單，回傳 [(members, tier_label), ...]。

    tier_label 是這一組「被哪些方式抓到」的組合說明（例如「檔名家族分組、
    100%（完全相同）」），用逗號／頓號分隔多個來源，方便在 Excel 裡一眼
    看出這組重複是靠什麼判斷出來的。同一張照片不管被幾種方式抓到，最終
    只會出現在一組、一列，不會因為同時符合檔名規則又符合雜湊距離而重複
    列出兩次——這是先前「四選一」設計改成「自動全部一起跑」後特別要處理
    的重點，避免同一批照片被家族分組跟雜湊分組各自抓一次、在 Excel 裡
    出現兩份幾乎一樣的列。
    """
    all_paths = [f["path"] for f in features]
    uf = UnionFind(all_paths)
    contributing_labels = defaultdict(set)  # path -> 這個檔案被哪些來源連過

    def apply_edges(edges):
        for a, b, label in edges:
            uf.union(a, b)
            contributing_labels[a].add(label)
            contributing_labels[b].add(label)

    apply_edges(build_family_edges(features))
    if not skip_hash:
        apply_edges(build_hash_edges(features, exact_max_dist, similar_max_dist,
                                      cross_folder=cross_folder))
    if groups_csv:
        apply_edges(build_groups_csv_edges(features, groups_csv))
    if pairs_csv:
        apply_edges(build_pairs_csv_edges(features, pairs_csv))

    by_path = {f["path"]: f for f in features}
    members_by_root = defaultdict(list)
    labels_by_root = defaultdict(set)
    for p in all_paths:
        root = uf.find(p)
        members_by_root[root].append(by_path[p])
        labels_by_root[root] |= contributing_labels.get(p, set())

    result = []
    for root, members in members_by_root.items():
        if len(members) < 2:
            continue
        labels = labels_by_root[root]
        tier_label = "、".join(sorted(labels)) if labels else "未知來源"
        result.append((members, tier_label))
    return result


# ---------------------------------------------------------------------------
# 第三步：判斷規則
# ---------------------------------------------------------------------------

def clean_stem(filename):
    stem = os.path.splitext(filename)[0]
    return EDIT_MARKER_RE.sub("", stem).strip(" _-")


def has_edit_marker(filename):
    return bool(EDIT_MARKER_RE.search(filename))


def has_edit_software(software):
    if not software:
        return False
    low = software.lower()
    return any(hint in low for hint in EDIT_SOFTWARE_HINTS) or True
    # 只要 Software 欄位非空就視為編修過；上面的 hint 清單只是額外參考，
    # 不縮小判斷範圍（很多修圖 App 會填自己的名字，清單不可能列全）。


def original_signal_score(f):
    """分數 0-5，越高越像「原始檔、沒被動過」。"""
    score = 0
    if not has_edit_software(f["software"]):
        score += 2
    if f["dt_original"] and f["dt_modified"] and f["dt_original"] == f["dt_modified"]:
        score += 2
    elif f["dt_modified"] is None:
        score += 1  # 沒有修改時間欄位，保守給一半分數
    if not has_edit_marker(f["filename"]):
        score += 1
    return score


def is_grayscale(f, threshold):
    return f["saturation"] is not None and f["saturation"] <= threshold


def detect_burst(group, burst_seconds):
    times = [f["dt_original"] for f in group if f["dt_original"]]
    if len(times) != len(group):
        return False  # 缺 EXIF 拍攝時間的組不判定為連拍，交給人工看
    spread = (max(times) - min(times)).total_seconds()
    # 注意：spread 必須「大於 0」才算連拍。如果同組所有檔案的拍攝時間完全
    # 相同（差 0 秒），代表這些檔案其實是同一張照片的不同編修版本（編修
    # 軟體通常不會改動 DateTimeOriginal，只會改 DateTime/修改時間），
    # 不是真的連續按了好幾次快門，這種情況要交給後面的原始檔判斷邏輯處理，
    # 不能當成連拍放行。
    return 0 < spread <= burst_seconds


def filename_adds_info(short_name, long_name):
    """long_name 的檔名是否比 short_name 多了「真正的資訊」（不只是編修字樣/數字）。"""
    short_stem = clean_stem(short_name)
    long_stem_raw = os.path.splitext(long_name)[0]
    if short_stem and short_stem in long_stem_raw:
        extra = long_stem_raw.replace(short_stem, "", 1)
        extra_wo_markers = EDIT_MARKER_RE.sub("", extra)
        extra_meaningful = re.sub(r"[\s\-_\d]", "", extra_wo_markers)
        return len(extra_meaningful) >= 2
    return False


def rank_key(f):
    """由高到低排序用的 key：原始檔訊號 → 解析度 → 檔案大小 → 檔名長度。
    （只用來給「黑白 vs 彩色」那個分支挑最佳彩色版，一般分支改用 pick_best，
    因為一般分支需要「接近就算打平」的容許誤差，不能用精確比較。）"""
    return (
        original_signal_score(f),
        f["width"] * f["height"],
        f["size_bytes"],
        len(os.path.splitext(f["filename"])[0]),
    )


def bytes_per_pixel(f):
    px = f["width"] * f["height"]
    return f["size_bytes"] / px if px else 0.0


def pick_best(candidates, near_tie_pct):
    """依「原始檔訊號 → 解析度 → 檔案大小 → 檔名長度」逐層篩選，每一層只有
    「贏得夠明顯」（超過 near_tie_pct% 的差距）才算數，贏得不夠明顯就當作
    這一層打平、進下一層比。所有層都打平，回傳 None，交給人工複核。

    這是為了處理像「兩張照片解析度只差 0.6%」這種情況——差距這麼小很可能
    只是不同次存檔/裁切造成的雜訊，不是真的哪張比較好，不該讓程式自己選，
    要交給人看。EXIF 完全缺失（Software、拍攝時間都沒有）時，原始檔訊號這
    一層本來就分不出高下，更容易走到這種「解析度/檔案大小都很接近」的情況。
    """
    # 第一層：原始檔訊號（離散分數 0-5，本來就沒有「接近」的問題，精確比較即可）
    max_signal = max(original_signal_score(f) for f in candidates)
    pool = [f for f in candidates if original_signal_score(f) == max_signal]
    if len(pool) == 1:
        return pool[0], pool, f"原始檔訊號分數最高（{max_signal}/5）"

    # 第二層：解析度，差距要超過 near_tie_pct% 才算贏
    max_px = max(f["width"] * f["height"] for f in pool)
    next_pool = [f for f in pool
                 if max_px and (max_px - f["width"] * f["height"]) / max_px * 100 <= near_tie_pct]
    if len(next_pool) == 1:
        return next_pool[0], next_pool, f"解析度明顯最佳（其餘檔案低於 {near_tie_pct}% 容許範圍）"
    pool = next_pool

    # 第三層：檔案大小，差距要超過 near_tie_pct% 才算贏
    max_size = max(f["size_bytes"] for f in pool)
    next_pool = [f for f in pool
                 if max_size and (max_size - f["size_bytes"]) / max_size * 100 <= near_tie_pct]
    if len(next_pool) == 1:
        return next_pool[0], next_pool, f"檔案大小明顯最佳（其餘檔案低於 {near_tie_pct}% 容許範圍）"
    pool = next_pool

    # 第四層：檔名長度（資訊豐富度），這層是離散整數，精確比較即可
    max_len = max(len(os.path.splitext(f["filename"])[0]) for f in pool)
    next_pool = [f for f in pool if len(os.path.splitext(f["filename"])[0]) == max_len]
    if len(next_pool) == 1:
        return next_pool[0], next_pool, "檔名資訊最豐富"

    return None, next_pool, (
        f"有 {len(next_pool)} 個檔案在原始檔訊號/解析度（{near_tie_pct}% 內）/"
        f"檔案大小（{near_tie_pct}% 內）/檔名長度上都很接近，程式無法自動判斷"
    )


def decide_group(group, grayscale_sat_threshold, burst_seconds, near_tie_pct=3.0):
    """回傳每個檔案的 (action, reason)。"""
    decisions = {}

    if detect_burst(group, burst_seconds):
        for f in group:
            decisions[f["path"]] = (ACTION_BURST, "同組拍攝時間差在門檻內，判定為連續拍攝")
        return decisions

    # 黑白 vs 彩色：先找出組內是否同時有黑白與明顯彩色的檔案
    grays = [f for f in group if is_grayscale(f, grayscale_sat_threshold)]
    colors = [f for f in group if not is_grayscale(f, grayscale_sat_threshold)]
    force_keep_color = bool(grays) and bool(colors)

    # 檔名多資訊：兩兩比較，標記出「兩者皆留」的檔案
    keep_both_paths = set()
    for i in range(len(group)):
        for j in range(len(group)):
            if i == j:
                continue
            a, b = group[i], group[j]
            if filename_adds_info(a["filename"], b["filename"]):
                keep_both_paths.add(a["path"])
                keep_both_paths.add(b["path"])

    for f in keep_both_paths:
        decisions[f] = (ACTION_KEEP_BOTH, "修圖版檔名比原始檔多了實質資訊，兩者都保留")

    remaining = [f for f in group if f["path"] not in decisions]
    if not remaining:
        return decisions

    if force_keep_color:
        color_remaining = [f for f in remaining if f["path"] not in
                            {g["path"] for g in grays}]
        gray_remaining = [f for f in remaining if f["path"] in
                           {g["path"] for g in grays}]
        for f in gray_remaining:
            decisions[f["path"]] = (ACTION_DELETE, "黑白修圖版，同組已有彩色版本")
        if color_remaining:
            best = max(color_remaining, key=rank_key)
            for f in color_remaining:
                if f["path"] == best["path"]:
                    decisions[f["path"]] = (ACTION_KEEP, "彩色原始版，同組黑白版已標記刪除")
                else:
                    decisions[f["path"]] = (ACTION_DELETE, "彩色但非組內最佳版本")
        return decisions

    # 一般情況：逐層挑最佳，差距在容許範圍內就當打平，轉人工複核
    best, tied_pool, reason = pick_best(remaining, near_tie_pct)

    if best is None:
        bpp_note = "；每像素位元組(壓縮密度)：" + "、".join(
            f'{f["filename"]}={bytes_per_pixel(f):.3f}' for f in tied_pool
        )
        for f in remaining:
            decisions[f["path"]] = (ACTION_REVIEW, reason + bpp_note)
        return decisions

    for f in remaining:
        if f["path"] == best["path"]:
            decisions[f["path"]] = (ACTION_KEEP, reason)
        else:
            decisions[f["path"]] = (ACTION_DELETE, "同組內非最佳版本")
    return decisions


# ---------------------------------------------------------------------------
# 第四步：輸出 / 合併 Excel
# ---------------------------------------------------------------------------

def group_key(paths):
    """用組內所有檔案路徑算一個穩定的群組 ID，重新掃描時能對上同一組。"""
    joined = "|".join(sorted(paths))
    return hashlib.md5(joined.encode("utf-8")).hexdigest()[:10]


def load_existing_status(output_path):
    """回傳 {檔案路徑: 狀態} 的字典，供合併用。"""
    if not os.path.exists(output_path):
        return {}
    try:
        wb = load_workbook(output_path)
        if "檢查清單" not in wb.sheetnames:
            return {}
        ws = wb["檢查清單"]
        headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        path_idx = headers.index("檔案路徑")
        status_idx = headers.index("狀態")
        result = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[path_idx]:
                result[row[path_idx]] = row[status_idx] or STATUS_UNCONFIRMED
        return result
    except Exception as e:
        print(f"⚠️  讀取舊 Excel 狀態時發生問題（將視為沒有舊狀態）：{e}")
        return {}


def build_thumbnail(path, tmp_dir, max_dim=120):
    """把縮圖存到「獨立的暫存資料夾」（tmp_dir），絕對不寫進原始照片所在的
    資料夾——早期版本曾經把縮圖存成 <原始檔名>.__thumb__.jpg、跟原始照片放
    在同一個資料夾，一旦程式中途被中斷（例如逾時被強制關閉），縮圖檔案就
    會留在妳的豆皮檔案庫裡當垃圾，甚至可能在下次掃描時被誤判成一張新照片
    （2026-09-23 實測案例：v11 資料夾因為逾時中斷，留下 86 個這種殘留檔）。
    改成寫進呼叫端用 tempfile.mkdtemp() 建立的獨立暫存資料夾後，就算程式
    中途被砍掉，殘留的也只會是系統暫存目錄裡的檔案，不會混進妳的原始檔案
    庫；呼叫端在整批處理完（不管成功或失敗）都要用 shutil.rmtree 清掉這個
    暫存資料夾。"""
    try:
        with Image.open(path) as im:
            im = im.convert("RGB")
            im.thumbnail((max_dim, max_dim))
            thumb_path = os.path.join(tmp_dir, f"{hashlib.md5(path.encode('utf-8')).hexdigest()}.jpg")
            im.save(thumb_path, "JPEG", quality=70)
        return thumb_path
    except Exception:
        return None


def write_excel(output_path, rows, with_thumbnails=True):
    """rows: 每列是一個 dict，鍵對應下面 HEADERS。"""
    HEADERS = [
        "縮圖", "群組ID", "相似度分級", "檔名", "檔案路徑", "解析度", "檔案大小(KB)",
        "每像素位元組", "拍攝時間", "修改時間", "平均飽和度", "原始檔訊號分數",
        "建議動作", "判斷依據", "狀態",
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "檢查清單"
    ws.append(HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    col_widths = [16, 12, 10, 28, 46, 12, 12, 12, 18, 18, 12, 14, 24, 34, 12]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    dv = DataValidation(
        type="list",
        formula1=f'"{",".join(STATUS_OK_COLORS.keys())}"',
        allow_blank=True,
    )
    ws.add_data_validation(dv)

    # 縮圖一律寫進系統暫存資料夾（tempfile.mkdtemp()），絕對不寫進原始照片
    # 所在的資料夾——用 try/finally 確保不管中間有沒有出錯（包含程式被強制
    # 中斷），這個暫存資料夾最後都會被整個刪掉，不會在妳的檔案庫裡留下任何
    # 殘留檔案。早期版本把縮圖存成 <原始檔名>.__thumb__.jpg、跟原始照片放
    # 在同一個資料夾，一旦程式中途被中斷（例如逾時被強制關閉）就會留下垃圾
    # 檔案，甚至可能在下次掃描時被誤判成一張新照片（2026-09-23 實測案例：
    # v11 資料夾因為逾時中斷，留下 86 個這種殘留檔）。
    thumb_tmp_dir = tempfile.mkdtemp(prefix="digikam_review_thumbs_") if with_thumbnails else None
    try:
        r = 2
        for row in rows:
            ws.cell(row=r, column=2, value=row["group_id"])
            ws.cell(row=r, column=3, value=row["tier_label"])
            ws.cell(row=r, column=4, value=row["filename"])
            ws.cell(row=r, column=5, value=row["path"])
            ws.cell(row=r, column=6, value=f'{row["width"]}x{row["height"]}')
            ws.cell(row=r, column=7, value=round(row["size_bytes"] / 1024, 1))
            ws.cell(row=r, column=8, value=round(bytes_per_pixel(row), 3)
                     if row["width"] and row["height"] else "")
            ws.cell(row=r, column=9, value=row["dt_original"].strftime("%Y-%m-%d %H:%M:%S")
                     if row["dt_original"] else "")
            ws.cell(row=r, column=10, value=row["dt_modified"].strftime("%Y-%m-%d %H:%M:%S")
                     if row["dt_modified"] else "")
            ws.cell(row=r, column=11, value=round(row["saturation"], 1)
                     if row["saturation"] is not None else "")
            ws.cell(row=r, column=12, value=original_signal_score(row))
            action_cell = ws.cell(row=r, column=13, value=row["action"])
            ws.cell(row=r, column=14, value=row["reason"])
            status_cell = ws.cell(row=r, column=15, value=row["status"])
            dv.add(status_cell)

            fill_hex = ACTION_FILL.get(row["action"])
            if fill_hex:
                action_cell.fill = PatternFill("solid", fgColor=fill_hex)
            status_fill = STATUS_OK_COLORS.get(row["status"], "FFFFFF")
            status_cell.fill = PatternFill("solid", fgColor=status_fill)

            if with_thumbnails:
                thumb = build_thumbnail(row["path"], thumb_tmp_dir)
                if thumb:
                    try:
                        img = XLImage(thumb)
                        img.width, img.height = 90, 68
                        ws.add_image(img, f"A{r}")
                        ws.row_dimensions[r].height = 54
                    except Exception:
                        pass
            r += 1

        ws.freeze_panes = "B2"

        # 摘要頁
        summary = wb.create_sheet("摘要", 0)
        total_groups = len({row["group_id"] for row in rows})
        total_files = len(rows)
        action_counts = defaultdict(int)
        for row in rows:
            action_counts[row["action"]] += 1
        unconfirmed = sum(1 for row in rows if row["status"] == STATUS_UNCONFIRMED)

        summary.append(["產生時間", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
        summary.append(["組數", total_groups])
        summary.append(["涉及檔案數", total_files])
        summary.append(["尚未確認的檔案數", unconfirmed])
        summary.append([])
        summary.append(["建議動作", "檔案數"])
        for action, count in action_counts.items():
            summary.append([action, count])
        for col, w in zip("AB", (22, 40)):
            summary.column_dimensions[col].width = w

        wb.save(output_path)
    finally:
        if thumb_tmp_dir:
            shutil.rmtree(thumb_tmp_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="digiKam 重複照片智能檢查清單產生器")
    parser.add_argument("folder", help="要掃描的資料夾（含子資料夾）")
    parser.add_argument("--output", required=True, help="輸出/合併的 Excel 路徑")
    parser.add_argument("--skip-hash", action="store_true",
                        help="跳過內建感知雜湊分組，只做檔名家族分組（＋人工輸入，"
                             "如果有給的話）；圖片數量非常多、雜湊兩兩比對太慢時使用")
    parser.add_argument("--cross-folder", action="store_true",
                        help="內建雜湊分組預設只在同一個資料夾內兩兩比對（活動辦在"
                             "不同資料夾，重複照片幾乎不會跨資料夾出現，這樣可以大幅"
                             "減少比對量）；加這個參數改成整批資料夾全部兩兩比對，"
                             "圖片數量多時會明顯變慢")
    parser.add_argument("--hash-size", type=int, default=16)
    parser.add_argument("--exact-max-dist", type=int, default=3)
    parser.add_argument("--similar-max-dist", type=int, default=14)
    parser.add_argument("--burst-seconds", type=int, default=10)
    parser.add_argument("--grayscale-sat", type=float, default=18)
    parser.add_argument("--near-tie-pct", type=float, default=3.0,
                        help="解析度/檔案大小差距在這個百分比以內視為打平、轉人工複核"
                             "（避免因為 0.6%% 這種雜訊等級的差距就自動選邊），預設 3.0")
    parser.add_argument("--no-thumbnails", action="store_true")
    parser.add_argument("--groups-csv", default=None,
                        help="疊加外部分組結果（欄位 group_id,file_path），例如從 digiKam "
                             "資料庫匯出的重複組清單；跟自動分組（檔名家族＋內建雜湊）的"
                             "結果合併，不是取代")
    parser.add_argument("--pairs-csv", default=None,
                        help="疊加外部「兩兩配對」清單（欄位 file_a,file_b），程式會自動"
                             "用 Union-Find 把有遞移關係的配對合併成同一組——專門解決"
                             "digiKam 自己把同一批重複檔案拆成好幾個參考組顯示的問題，"
                             "跟自動分組的結果合併，不是取代")
    parser.add_argument("--write-labels", action="store_true",
                        help="把本次建議動作寫成 digiKam 色彩/揀選標籤（XMP metadata），"
                             "需要系統已安裝 exiftool（brew install exiftool）；寫入後要在"
                             "digiKam 裡對該資料夾「重新讀取中繼資料」才會顯示在縮圖畫面上，"
                             "細節見腳本開頭「digiKam 標籤整合」說明")
    # 舊版參數，保留相容：現在檔名家族分組跟內建雜湊分組預設就會一起跑，
    # 這兩個參數已經不影響行為，只是避免舊指令直接報錯。
    parser.add_argument("--family", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--tier", choices=["exact", "similar"], default=None,
                        help=argparse.SUPPRESS)
    args = parser.parse_args()

    if args.family or args.tier:
        print("ℹ️  --family / --tier 是舊版參數：現在檔名家族分組跟內建雜湊分組"
              "預設就會一起跑，不需要再指定，這裡直接忽略。")

    print(f"掃描資料夾：{args.folder} ...")
    paths = list(find_images(args.folder))
    print(f"共 {len(paths)} 張圖片，開始計算特徵（解析度/雜湊/飽和度/EXIF）...")

    features = []
    for i, p in enumerate(paths, 1):
        f = extract_feature(p, args.hash_size)
        if f.get("error"):
            print(f"  [{i}/{len(paths)}] ⚠️ 跳過（讀取失敗）：{p} ({f['error']})")
            continue
        features.append(f)
        if i % 200 == 0:
            print(f"  已處理 {i}/{len(paths)} 張")

    source_desc = "檔名家族分組" + ("" if args.skip_hash else "＋內建雜湊分組")
    if args.groups_csv:
        source_desc += "＋外部分組 CSV"
    if args.pairs_csv:
        source_desc += "＋外部配對 CSV"
    print(f"分組中（{source_desc}，自動合併去重）...")
    groups = combine_all_groups(
        features,
        skip_hash=args.skip_hash,
        exact_max_dist=args.exact_max_dist,
        similar_max_dist=args.similar_max_dist,
        groups_csv=args.groups_csv,
        pairs_csv=args.pairs_csv,
        cross_folder=args.cross_folder,
    )
    print(f"找到 {len(groups)} 組重複／近似照片")

    existing_status = load_existing_status(args.output)

    rows = []
    seen_paths = set()
    for group, tier_label in groups:
        gid = group_key([f["path"] for f in group])
        decisions = decide_group(group, args.grayscale_sat, args.burst_seconds, args.near_tie_pct)
        for f in group:
            action, reason = decisions.get(f["path"], (ACTION_REVIEW, "未能判斷"))
            status = existing_status.get(f["path"], STATUS_UNCONFIRMED)
            seen_paths.add(f["path"])
            rows.append({
                **f,
                "group_id": gid,
                "tier_label": tier_label,
                "action": action,
                "reason": reason,
                "status": status,
            })

    # 如果 --output 已存在，把「這次沒掃到、但上次還在清單裡」的列也留著
    # （例如檔案已被搬走/刪除，或這次不再被判定為重複組的一部分）——
    # 用讀舊檔的方式最單純：舊清單裡「這次沒有重新產生」的路徑，原樣保留。
    # 2026-09-23 改版：以前這裡只保留「上次是別種相似度分級」的舊列（因為
    # 以前一次只跑一種分組方式，其他分級的結果本來就不會在這次重新出現）；
    # 現在每次都會自動跑齊所有分組方式，所以改成單純比對路徑有沒有出現在
    # 這次結果裡，不再依賴相似度分級是否相同。
    if os.path.exists(args.output):
        try:
            wb_old = load_workbook(args.output)
            if "檢查清單" in wb_old.sheetnames:
                ws_old = wb_old["檢查清單"]
                headers = [c.value for c in next(ws_old.iter_rows(min_row=1, max_row=1))]
                idx = {h: i for i, h in enumerate(headers)}
                for row in ws_old.iter_rows(min_row=2, values_only=True):
                    old_path = row[idx["檔案路徑"]]
                    if old_path and old_path not in seen_paths:
                        rows.append({
                            "group_id": row[idx["群組ID"]],
                            "tier_label": row[idx.get("相似度分級", 2)],
                            "filename": row[idx["檔名"]],
                            "path": old_path,
                            "width": 0, "height": 0,
                            "size_bytes": (row[idx["檔案大小(KB)"]] or 0) * 1024,
                            "dt_original": None, "dt_modified": None,
                            "saturation": None, "software": "",
                            "action": row[idx["建議動作"]],
                            "reason": row[idx["判斷依據"]],
                            "status": row[idx["狀態"]] or STATUS_UNCONFIRMED,
                        })
        except Exception as e:
            print(f"⚠️  合併舊清單時發生問題（將只寫入本次結果）：{e}")

    write_excel(args.output, rows, with_thumbnails=not args.no_thumbnails)

    if args.write_labels:
        write_labels_to_files(rows)

    new_unconfirmed = sum(1 for r in rows if r["status"] == STATUS_UNCONFIRMED)
    print("\n" + "=" * 50)
    print(f"完成！已寫入 {args.output}")
    print(f"本次分組來源：{source_desc}，共 {len(groups)} 組、{len(paths)} 張圖片掃描")
    print(f"目前清單中「尚未確認」的檔案數：{new_unconfirmed}")
    print("=" * 50)


if __name__ == "__main__":
    main()
