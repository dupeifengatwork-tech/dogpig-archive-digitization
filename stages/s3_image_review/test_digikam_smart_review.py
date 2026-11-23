#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
digikam_smart_review.py 的回歸測試（可重複複驗的基準）
========================================================

這份測試不依賴任何真實照片——所有測試圖片都是程式當場合成出來的
（純色小圖 + 手動寫入的 EXIF），每次執行都會產生一模一樣的結果，
在任何一台機器上都能重跑、對答案。

涵蓋範圍：把這次討論過程中，用使用者實際提供的真實案例驗證過的每一種
情境（原始檔 vs 修圖版、黑白 vs 彩色、連拍、EXIF 完全缺失時的近似容忍度
判斷、檔名家族分組、digiKam 拆組後用配對清單合併、分組 CSV 裡的重複列、
建議動作 → digiKam 色彩/揀選標籤的對照表、2026-09-23 改版後「多種分組
來源自動合併、同一張照片不重複列出」的行為)，各自轉成一個獨立、快速、
不需要真實檔案就能驗證的測試案例，記錄下「應該要判斷成什麼」的基準答案。

=== 使用方式 ===
python3 test_digikam_smart_review.py

跟 digikam_smart_review.py 放在同一個資料夾執行即可。每次修改主程式
的判斷邏輯後，都應該重新跑一次這份測試，全部 PASS 才代表沒有把之前
用真實案例驗證過的行為改壞。

=== 這份測試「能」保證什麼、「不能」保證什麼 ===
能保證：判斷邏輯（decide_group / pick_best / 分組函式 / 多來源合併邏輯 /
建議動作到 digiKam 標籤的對照表）在這幾種已知情境下的行為跟預期一致，
未來改動邏輯時能立刻抓到回歸（原本對的變錯）。
不能保證：對妳實際檔案庫裡所有沒出現過的檔名/EXIF 組合都判斷正確——
那些終究要靠實際跑過、人工複核 Excel 才能驗證，這份測試只是最基本的
安全網，不是「保證全對」的證明。也不能保證 --write-labels 寫進去的
XMP 標籤在妳那台 digiKam 上顯示出來的顏色跟預期一致——這部分需要
妳照著使用指南的說明手動驗證一次 exiftool 讀寫是否符合預期。
"""

import os
import shutil
import sys
import tempfile
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import digikam_smart_review as dsr  # noqa: E402

PASS = []
FAIL = []


def check(name, condition, detail=""):
    if condition:
        PASS.append(name)
        print(f"  ✅ {name}")
    else:
        FAIL.append(name)
        print(f"  ❌ {name}  {detail}")


def feat(path, width=1280, height=960, size_bytes=200_000, software="",
         dt_original=None, dt_modified=None, saturation=130.0, phash=None):
    """建構一個跟 extract_feature() 回傳格式相同的假特徵字典，
    不需要真的去讀圖片檔案，讓判斷邏輯測試可以快速、精準地控制每個變因。"""
    return {
        "path": path,
        "filename": os.path.basename(path),
        "size_bytes": size_bytes,
        "width": width,
        "height": height,
        "megapixels": round(width * height / 1e6, 2),
        "phash": phash,
        "saturation": saturation,
        "software": software,
        "dt_original": dt_original,
        "dt_modified": dt_modified,
        "error": None,
    }


# ---------------------------------------------------------------------------
# 1. 原始檔 vs 修圖版（對應「膠帶畫廊」那組：1 原始 + 2 個 Photoshop 調色版）
# ---------------------------------------------------------------------------
def test_original_vs_edited():
    t0 = datetime(2008, 3, 8, 0, 52, 0)
    group = [
        feat("orig.jpg", software="", dt_original=t0, dt_modified=t0,
             size_bytes=560_000, saturation=159.3),
        feat("edit1.jpg", software="Adobe Photoshop 7.0", dt_original=t0,
             dt_modified=datetime(2008, 4, 3, 1, 56, 38),
             size_bytes=131_000, saturation=110.1),
        feat("edit2.jpg", software="Adobe Photoshop 7.0", dt_original=t0,
             dt_modified=datetime(2008, 4, 3, 1, 52, 40),
             size_bytes=128_000, saturation=155.3),
    ]
    decisions = dsr.decide_group(group, grayscale_sat_threshold=18, burst_seconds=10)
    check("原始檔判斷：無編修痕跡的檔案被保留",
          decisions["orig.jpg"][0] == dsr.ACTION_KEEP, decisions["orig.jpg"])
    check("原始檔判斷：兩個 Photoshop 版本都建議刪除",
          decisions["edit1.jpg"][0] == dsr.ACTION_DELETE
          and decisions["edit2.jpg"][0] == dsr.ACTION_DELETE)


# ---------------------------------------------------------------------------
# 2. 黑白版 vs 彩色原始版（對應「烏克麗利」那組：漏抓的黑白版）
# ---------------------------------------------------------------------------
def test_grayscale_vs_color():
    t0 = datetime(2008, 3, 8, 0, 52, 0)
    group = [
        feat("P1080353.JPG", software="", dt_original=t0, dt_modified=t0,
             size_bytes=560_000, saturation=159.3),
        feat("colored_edit.jpg", software="Photoshop", dt_original=t0,
             dt_modified=datetime(2008, 4, 3, 1, 52, 40),
             size_bytes=128_000, saturation=155.3),
        feat("grayscale.jpg", software="", dt_original=None, dt_modified=None,
             size_bytes=147_000, saturation=0.0),
    ]
    decisions = dsr.decide_group(group, grayscale_sat_threshold=18, burst_seconds=10)
    check("黑白版即使檔案不小，仍判定刪除",
          decisions["grayscale.jpg"][0] == dsr.ACTION_DELETE
          and "黑白" in decisions["grayscale.jpg"][1])
    check("彩色原始檔保留，且理由提到彩色",
          decisions["P1080353.JPG"][0] == dsr.ACTION_KEEP)


# ---------------------------------------------------------------------------
# 3. 連續拍攝不誤刪
# ---------------------------------------------------------------------------
def test_burst_not_deleted():
    group = [
        feat("burst1.jpg", dt_original=datetime(2010, 3, 7, 19, 30, 0)),
        feat("burst2.jpg", dt_original=datetime(2010, 3, 7, 19, 30, 3),
             width=1275, height=955),
    ]
    decisions = dsr.decide_group(group, grayscale_sat_threshold=18, burst_seconds=10)
    check("拍攝時間差在門檻內，兩張都判定為連拍、不建議刪除",
          decisions["burst1.jpg"][0] == dsr.ACTION_BURST
          and decisions["burst2.jpg"][0] == dsr.ACTION_BURST)


def test_same_timestamp_is_not_burst():
    """回歸測試：同一張照片的不同編修版拍攝時間會完全相同（spread=0），
    不能被誤判成連拍——這是先前修過的一個 bug。"""
    t0 = datetime(2008, 3, 8, 0, 52, 0)
    group = [
        feat("orig.jpg", software="", dt_original=t0, dt_modified=t0,
             size_bytes=560_000, saturation=159.3),
        feat("edit.jpg", software="Photoshop", dt_original=t0,
             dt_modified=datetime(2008, 4, 3, 1, 56, 38),
             size_bytes=131_000, saturation=110.1),
    ]
    decisions = dsr.decide_group(group, grayscale_sat_threshold=18, burst_seconds=10)
    check("拍攝時間完全相同（同一原始檔的編修版）不會被誤判成連拍",
          decisions["orig.jpg"][0] != dsr.ACTION_BURST)


# ---------------------------------------------------------------------------
# 4. EXIF 完全缺失時的近似容忍度判斷（對應「屋頂合照」那組）
# ---------------------------------------------------------------------------
def test_near_tie_falls_back_to_size():
    # 注意：檔名刻意避開 EDIT_MARKERS 清單（例如 "v2" 本身就是一個編修標記，
    # 用它當檔名會汙染 original_signal_score，讓這個測試變成在測別的東西）。
    group = [
        feat("IMG_2201.jpg", width=552, height=413, size_bytes=162_520, saturation=100.1),
        feat("IMG_2202.jpg", width=550, height=412, size_bytes=200_234, saturation=107.5),
        feat("IMG_2203.jpg", width=543, height=403, size_bytes=185_724, saturation=108.6),
    ]
    decisions = dsr.decide_group(group, grayscale_sat_threshold=18, burst_seconds=10,
                                  near_tie_pct=3.0)
    check("解析度差距在容許範圍內（0.6%）時改比檔案大小，選出壓縮密度最高的版本",
          decisions["IMG_2202.jpg"][0] == dsr.ACTION_KEEP, decisions)


def test_exact_tie_goes_to_review():
    """完全無法分辨優劣時，要老實承認、轉人工複核，不能硬選一個。"""
    group = [
        feat("a.jpg", width=1000, height=1000, size_bytes=100_000),
        feat("b.jpg", width=1000, height=1000, size_bytes=100_000),
    ]
    decisions = dsr.decide_group(group, grayscale_sat_threshold=18, burst_seconds=10)
    check("所有指標都打平時，轉人工複核而不是隨便選一個",
          decisions["a.jpg"][0] == dsr.ACTION_REVIEW
          and decisions["b.jpg"][0] == dsr.ACTION_REVIEW)


# ---------------------------------------------------------------------------
# 5. 檔名家族分組（對應「034高斯效果」跟「Chulki tour dogpig拷貝」兩組）
# ---------------------------------------------------------------------------
def test_family_key_matches_effect_suffix():
    base = "/x/200804膠帶畫廊"
    paths = [
        f"{base}/200804膠帶畫廊 20080000 034(高斯效果再處理).jpg",
        f"{base}/200804膠帶畫廊20080000 034.JPG",
        f"{base}/20080000 034(高斯效果).jpg",
    ]
    keys = {dsr.normalize_family_key(p) for p in paths}
    check("原始檔／高斯效果／高斯效果再處理 三個檔名歸一化後是同一把鑰匙",
          len(keys) == 1, keys)


def test_family_key_matches_copy_suffix():
    base = "/x/201003張惠笙豆皮表演"
    paths = [
        f"{base}/Chulki tour dogpig拷貝.jpg",
        f"{base}/Chulki tour dogpig拷貝2.jpg",
        f"{base}/Chulki-tour-dogpig.jpg",
    ]
    keys = {dsr.normalize_family_key(p) for p in paths}
    check("拷貝／拷貝2／原始檔（連字號差異、大小寫差異）三個檔名歸一化後是同一把鑰匙",
          len(keys) == 1, keys)


def test_family_key_does_not_merge_bare_sequence_numbers():
    base = "/x/201003張惠笙豆皮表演"
    key1 = dsr.normalize_family_key(f"{base}/惠笙20100307表演-海報01.jpg")
    key2 = dsr.normalize_family_key(f"{base}/惠笙20100307表演-海報02.jpg")
    check("純數字尾碼（01/02，無複製語意字樣）刻意不會被判成同一把鑰匙",
          key1 != key2, (key1, key2))


# ---------------------------------------------------------------------------
# 6. digiKam 拆組後用配對清單合併（對應「烏克麗麗海報」拆成 3 組那次）
# ---------------------------------------------------------------------------
def test_pairs_csv_merges_fragmented_groups():
    tmpdir = tempfile.mkdtemp()
    try:
        paths = [os.path.join(tmpdir, f"p{i}.jpg") for i in range(4)]
        for p in paths:
            open(p, "wb").close()
        features = [feat(p) for p in paths]

        csv_path = os.path.join(tmpdir, "pairs.csv")
        with open(csv_path, "w", encoding="utf-8") as fh:
            fh.write("file_a,file_b\n")
            fh.write(f"{paths[0]},{paths[1]}\n")  # digiKam 顯示的「組一」
            fh.write(f"{paths[2]},{paths[3]}\n")  # digiKam 顯示的「組二」
            fh.write(f"{paths[1]},{paths[2]}\n")  # 兩組之間的橋接關係

        groups = dsr.group_by_pairs_csv(features, csv_path)
        check("三段配對關係（含跨組橋接）正確合併成 1 組、4 個檔案",
              len(groups) == 1 and len(groups[0]) == 4,
              [len(g) for g in groups])
    finally:
        shutil.rmtree(tmpdir)


def test_groups_csv_dedups_repeated_rows():
    tmpdir = tempfile.mkdtemp()
    try:
        paths = [os.path.join(tmpdir, f"p{i}.jpg") for i in range(3)]
        for p in paths:
            open(p, "wb").close()
        features = [feat(p) for p in paths]

        csv_path = os.path.join(tmpdir, "groups.csv")
        with open(csv_path, "w", encoding="utf-8") as fh:
            fh.write("group_id,file_path\n")
            fh.write(f"g1,{paths[0]}\n")
            fh.write(f"g1,{paths[0]}\n")  # 重複列（實測 digiKam 匯出會發生）
            fh.write(f"g1,{paths[1]}\n")
            fh.write(f"g1,{paths[2]}\n")

        groups = dsr.group_by_csv(features, csv_path)
        check("同一組內重複的檔案路徑只算一次，不會在同一組出現兩次",
              len(groups) == 1 and len(groups[0]) == 3,
              [len(g) for g in groups])
    finally:
        shutil.rmtree(tmpdir)


# ---------------------------------------------------------------------------
# 7. 修圖版檔名多了實質資訊 → 兩者皆留
# ---------------------------------------------------------------------------
def test_filename_extra_info_keeps_both():
    group = [
        feat("034.jpg", size_bytes=200_000),
        feat("034 五碗節目現場.jpg", size_bytes=210_000),
    ]
    decisions = dsr.decide_group(group, grayscale_sat_threshold=18, burst_seconds=10)
    check("修圖版檔名多了活動名稱這類實質資訊時，兩者都保留",
          decisions["034.jpg"][0] == dsr.ACTION_KEEP_BOTH
          and decisions["034 五碗節目現場.jpg"][0] == dsr.ACTION_KEEP_BOTH)


# ---------------------------------------------------------------------------
# 8. 端到端：真的合成一張圖片，確認 EXIF／飽和度真的讀得出來（不是只測邏輯）
# ---------------------------------------------------------------------------
def test_extract_feature_reads_real_exif_and_saturation():
    from PIL import Image
    tmpdir = tempfile.mkdtemp()
    try:
        # 高飽和度彩色圖
        color_path = os.path.join(tmpdir, "color.jpg")
        im = Image.new("RGB", (200, 150), (220, 30, 30))
        exif = im.getexif()
        exif[305] = "Adobe Photoshop 7.0"
        im.save(color_path, exif=exif, quality=90)

        # 灰階圖
        gray_path = os.path.join(tmpdir, "gray.jpg")
        Image.new("RGB", (200, 150), (128, 128, 128)).save(gray_path, quality=90)

        f_color = dsr.extract_feature(color_path, hash_size=8)
        f_gray = dsr.extract_feature(gray_path, hash_size=8)

        check("extract_feature 讀到實際寫入的 Software 標籤",
              f_color["software"] == "Adobe Photoshop 7.0", f_color["software"])
        check("extract_feature 算出的解析度跟合成圖片一致",
              (f_color["width"], f_color["height"]) == (200, 150))
        check("純灰色圖片的飽和度接近 0（正確判定為黑白）",
              f_gray["saturation"] < 5, f_gray["saturation"])
        check("高飽和度彩色圖片的飽和度明顯不是 0",
              f_color["saturation"] > 50, f_color["saturation"])
    finally:
        shutil.rmtree(tmpdir)


# ---------------------------------------------------------------------------
# 9. 建議動作 → digiKam 色彩/揀選標籤對照表（--write-labels 用）
# ---------------------------------------------------------------------------
def test_action_to_labels_mapping():
    cases = {
        dsr.ACTION_KEEP: (dsr.COLOR_LABEL_GREEN, dsr.PICK_LABEL_ACCEPTED),
        dsr.ACTION_KEEP_BOTH: (dsr.COLOR_LABEL_YELLOW, dsr.PICK_LABEL_ACCEPTED),
        dsr.ACTION_BURST: (dsr.COLOR_LABEL_BLUE, dsr.PICK_LABEL_ACCEPTED),
        dsr.ACTION_DELETE: (dsr.COLOR_LABEL_RED, dsr.PICK_LABEL_REJECTED),
        dsr.ACTION_REVIEW: (dsr.COLOR_LABEL_GRAY, dsr.PICK_LABEL_PENDING),
    }
    all_ok = True
    for action, expected in cases.items():
        got = dsr.action_to_labels(action)
        if got != expected:
            all_ok = False
    check("五種建議動作各自對到預期的 (色彩標籤, 揀選標籤)", all_ok, cases)

    check("刪除類建議一律用「已拒絕」揀選標籤（不會誤標成已接受）",
          dsr.action_to_labels(dsr.ACTION_DELETE)[1] == dsr.PICK_LABEL_REJECTED)
    check("待複核建議一律用「待定」揀選標籤，不會被誤標成已接受或已拒絕",
          dsr.action_to_labels(dsr.ACTION_REVIEW)[1] == dsr.PICK_LABEL_PENDING)

    unknown_color, unknown_pick = dsr.action_to_labels("不存在的動作")
    check("未知的建議動作保守回傳「無色彩標籤／待定」，不會亂猜",
          (unknown_color, unknown_pick) == (dsr.COLOR_LABEL_NONE, dsr.PICK_LABEL_PENDING))


def test_build_exiftool_args_shape():
    args = dsr.build_exiftool_args("/x/y/照片.jpg", dsr.COLOR_LABEL_RED, dsr.PICK_LABEL_REJECTED)
    check("exiftool 參數含正確的色彩標籤設定",
          "-XMP-digiKam:ColorLabel=1" in args, args)
    check("exiftool 參數含正確的揀選標籤設定",
          "-XMP-digiKam:PickLabel=1" in args, args)
    check("exiftool 參數要求就地覆寫、不留備份檔",
          "-overwrite_original" in args, args)
    check("exiftool 參數最後一項是檔案路徑",
          args[-1] == "/x/y/照片.jpg", args)


def test_write_labels_reports_missing_exiftool_without_crashing(monkeypatch=None):
    """系統沒裝 exiftool 時，write_labels_to_files 應該印出安裝提示、優雅返回，
    而不是丟例外——這是使用者第一次用這個功能、還沒裝 exiftool 時最常見的狀況。"""
    import io
    import contextlib

    original_which = dsr.shutil.which
    dsr.shutil.which = lambda name: None  # 模擬系統沒有 exiftool
    try:
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            dsr.write_labels_to_files([dsr_test_row()])
        output = buf.getvalue()
        check("找不到 exiftool 時會印出 brew install 安裝提示、不會拋例外",
              "exiftool" in output and "brew install" in output, output)
    finally:
        dsr.shutil.which = original_which


# ---------------------------------------------------------------------------
# 11. 2026-09-23 修的 bug：縮圖不能寫進原始照片所在的資料夾
#     （v11 資料夾實測案例：程式被中途中斷，留下 86 個殘留縮圖檔）
# ---------------------------------------------------------------------------
def test_write_excel_does_not_leave_thumbnail_beside_source_photo():
    """write_excel 完成後，原始照片所在的資料夾裡不應該出現任何新檔案——
    縮圖只能寫進獨立的系統暫存資料夾，並在寫完 Excel 後整個清掉。"""
    from PIL import Image as PILImage

    tmpdir = tempfile.mkdtemp()
    try:
        photo_dir = os.path.join(tmpdir, "活動資料夾")
        os.makedirs(photo_dir)
        photo_path = os.path.join(photo_dir, "照片01.jpg")
        PILImage.new("RGB", (100, 80), (200, 50, 50)).save(photo_path, quality=90)

        before = set(os.listdir(photo_dir))
        row = {**feat(photo_path, width=100, height=80), "group_id": "g1",
               "tier_label": "測試", "action": dsr.ACTION_KEEP, "reason": "測試",
               "status": dsr.STATUS_UNCONFIRMED}
        output_path = os.path.join(tmpdir, "查重清單.xlsx")
        dsr.write_excel(output_path, [row], with_thumbnails=True)
        after = set(os.listdir(photo_dir))

        check("write_excel 產生縮圖後，原始照片資料夾裡沒有多出任何檔案",
              before == after, (before, after))
        check("Excel 檔案本身有正常產生",
              os.path.exists(output_path))
    finally:
        shutil.rmtree(tmpdir)


def test_write_excel_cleans_up_temp_dir_even_if_thumbnail_fails():
    """就算某一列的縮圖建立失敗（例如原始檔案已經不存在），暫存資料夾還是
    要被清乾淨，不能留下空的暫存垃圾資料夾。"""
    tmpdir = tempfile.mkdtemp()
    try:
        row = {**feat("/tmp/不存在的照片.jpg"), "group_id": "g1",
               "tier_label": "測試", "action": dsr.ACTION_KEEP, "reason": "測試",
               "status": dsr.STATUS_UNCONFIRMED}
        output_path = os.path.join(tmpdir, "查重清單.xlsx")

        tmp_dirs_created = []
        original_mkdtemp = dsr.tempfile.mkdtemp

        def tracking_mkdtemp(*args, **kwargs):
            d = original_mkdtemp(*args, **kwargs)
            tmp_dirs_created.append(d)
            return d

        dsr.tempfile.mkdtemp = tracking_mkdtemp
        try:
            dsr.write_excel(output_path, [row], with_thumbnails=True)
        finally:
            dsr.tempfile.mkdtemp = original_mkdtemp

        check("縮圖來源檔案不存在時仍不會拋例外、正常寫出 Excel",
              os.path.exists(output_path))
        check("暫存縮圖資料夾用完後會被整個刪掉，不留空資料夾",
              len(tmp_dirs_created) == 1 and not os.path.exists(tmp_dirs_created[0]),
              tmp_dirs_created)
    finally:
        shutil.rmtree(tmpdir)


def dsr_test_row():
    return {**feat("/tmp/不存在的測試檔案.jpg"), "action": dsr.ACTION_KEEP}


# ---------------------------------------------------------------------------
# 10. 2026-09-23 改版：多種分組來源自動合併，不再需要分好幾次下指令
#     （對應這次 v11 資料夾昨天只跑了 --family、忘記另外跑雜湊分組的案例）
# ---------------------------------------------------------------------------
class _FakeHash:
    """假的感知雜湊物件，只需要支援 `a - b` 算距離，不用真的算圖片雜湊，
    讓 combine_all_groups 的測試可以精準控制雜湊距離。"""
    def __init__(self, value):
        self.value = value

    def __sub__(self, other):
        return abs(self.value - other.value)


def test_combine_groups_runs_family_and_hash_together_by_default():
    """以前的痛點：不加 --family 就不會做檔名家族分組、不特別再跑一次就不會
    做雜湊分組，兩種要分開下指令才會都跑到。現在不用給任何分組參數，兩種
    應該自動一起跑：這裡故意設計「只能靠檔名家族分組抓到」跟「只能靠雜湊
    距離抓到」兩組完全獨立的照片，確認一次呼叫就能兩組都抓到。"""
    base = "/x/200804膠帶畫廊"
    features = [
        # 這一組只有檔名有規律（雜湊距離故意設很遠，內建雜湊分組抓不到）
        feat(f"{base}/20080000034.jpg", phash=_FakeHash(0)),
        feat(f"{base}/20080000034(高斯效果).jpg", phash=_FakeHash(200)),
        # 這一組檔名完全不相關（family 抓不到），但雜湊距離很近
        feat(f"{base}/DSC0001.jpg", phash=_FakeHash(50)),
        feat(f"{base}/中秋節活動照片.jpg", phash=_FakeHash(52)),
    ]
    groups = dsr.combine_all_groups(features, exact_max_dist=3, similar_max_dist=14)
    check("不加任何分組參數，就能同時抓到「檔名有規律」跟「雜湊距離近」兩組",
          len(groups) == 2, [(len(m), t) for m, t in groups])

    family_group = next((g for g, t in groups if "20080000034" in g[0]["filename"]), None)
    hash_group = next((g for g, t in groups
                        if any("DSC0001" in f["filename"] for f in g)), None)
    check("檔名家族那組的相似度分級標記有包含「檔名家族分組」",
          family_group is not None
          and any("檔名家族分組" in t for g, t in groups if g is family_group))
    check("雜湊距離那組的相似度分級標記有包含「100%（完全相同）」",
          hash_group is not None
          and any("100%（完全相同）" in t for g, t in groups if g is hash_group))


def test_combine_groups_no_duplicate_when_family_and_hash_agree():
    """同一批照片如果檔名家族分組跟雜湊分組都判定是同一組，最終合併結果裡
    每個檔案應該只出現一次，不會因為被兩種方式各自抓到就在 Excel 裡重複
    列出兩份幾乎一樣的列（這是把「四選一」改成「全部一起跑」後最需要特別
    驗證的地方）。"""
    base = "/x/200804膠帶畫廊"
    features = [
        feat(f"{base}/20080000034.jpg", phash=_FakeHash(0)),
        feat(f"{base}/20080000034(高斯效果).jpg", phash=_FakeHash(1)),  # 雜湊距離也很近
    ]
    groups = dsr.combine_all_groups(features, exact_max_dist=3, similar_max_dist=14)
    check("兩種方式都判定同一組時，最終只會合併成 1 組、2 個檔案（不重複列出）",
          len(groups) == 1 and len(groups[0][0]) == 2, groups)
    check("這組的相似度分級會同時標記出兩種來源（檔名家族分組、100%完全相同）",
          "檔名家族分組" in groups[0][1] and "100%（完全相同）" in groups[0][1],
          groups[0][1])


def test_combine_groups_merges_pairs_csv_with_automatic_sources():
    """--pairs-csv 現在是「疊加」而不是「取代」：一份完全靠人眼比對配對進來
    的檔案，應該能跟自動分組（家族/雜湊）抓到的組橋接合併成同一組。"""
    tmpdir = tempfile.mkdtemp()
    try:
        base = os.path.join(tmpdir, "201107烏克麗麗音樂會")
        os.makedirs(base, exist_ok=True)
        p_family_a = os.path.join(base, "海報01.jpg")
        p_family_b = os.path.join(base, "海報01拷貝.jpg")
        p_manual = os.path.join(base, "kupc2011.png")  # 檔名跟前兩個完全無關
        for p in (p_family_a, p_family_b, p_manual):
            open(p, "wb").close()
        features = [feat(p_family_a), feat(p_family_b), feat(p_manual)]

        csv_path = os.path.join(tmpdir, "pairs.csv")
        with open(csv_path, "w", encoding="utf-8") as fh:
            fh.write("file_a,file_b\n")
            fh.write(f"{p_family_b},{p_manual}\n")  # 人眼看出 b 跟 manual 其實同一張

        groups = dsr.combine_all_groups(features, skip_hash=True, pairs_csv=csv_path)
        check("檔名家族分組抓到的一對，跟 pairs-csv 額外指出的橋接關係合併成 1 組 3 個檔案",
              len(groups) == 1 and len(groups[0][0]) == 3, groups)
        check("合併後的分級標記同時看得出「檔名家族分組」跟「外部配對」兩種來源",
              "檔名家族分組" in groups[0][1] and "外部配對（digiKam 匯出）" in groups[0][1],
              groups[0][1])
    finally:
        shutil.rmtree(tmpdir)


def test_hash_grouping_defaults_to_same_folder_only():
    """效能優化：內建雜湊分組預設只在同一個資料夾內兩兩比對，不會跨資料夾
    誤判成重複——這是為了避免資料夾數量多、每個資料夾圖片不多，但整批
    資料夾合計圖片數很大時，整批 O(n²) 比對跑非常久（實測 v11 資料夾
    2000+ 張圖片、100 個活動資料夾時整批比對逾時未完成）。"""
    features = [
        feat("/x/活動A/DSC0001.jpg", phash=_FakeHash(0)),
        feat("/x/活動B/DSC0002.jpg", phash=_FakeHash(1)),  # 雜湊距離很近，但在不同資料夾
    ]
    groups_same_folder_only = dsr.combine_all_groups(features, exact_max_dist=3, similar_max_dist=14)
    check("預設（未加 --cross-folder）：不同資料夾即使雜湊距離很近也不會被判成同一組",
          len(groups_same_folder_only) == 0, groups_same_folder_only)

    groups_cross_folder = dsr.combine_all_groups(
        features, exact_max_dist=3, similar_max_dist=14, cross_folder=True)
    check("加 --cross-folder 後，跨資料夾但雜湊距離近的照片會被判成同一組",
          len(groups_cross_folder) == 1 and len(groups_cross_folder[0][0]) == 2,
          groups_cross_folder)


def test_hash_grouping_still_finds_matches_within_same_folder():
    """確認限制在同資料夾內比對，不會連「同資料夾內」原本抓得到的重複都漏掉。"""
    features = [
        feat("/x/活動A/DSC0001.jpg", phash=_FakeHash(0)),
        feat("/x/活動A/DSC0002.jpg", phash=_FakeHash(1)),
    ]
    groups = dsr.combine_all_groups(features, exact_max_dist=3, similar_max_dist=14)
    check("同一個資料夾內雜湊距離近的照片，預設（不用 --cross-folder）就能抓到",
          len(groups) == 1 and len(groups[0][0]) == 2, groups)


def test_combine_groups_skip_hash_only_uses_family():
    """--skip-hash 應該真的關掉雜湊分組，只留下檔名家族分組（＋人工輸入），
    避免圖片數量很多時被 O(n²) 的雜湊比對拖慢。"""
    base = "/x/某資料夾"
    features = [
        feat(f"{base}/IMG001.jpg", phash=_FakeHash(0)),
        feat(f"{base}/IMG002.jpg", phash=_FakeHash(1)),  # 雜湊距離很近，但檔名無規律
    ]
    groups = dsr.combine_all_groups(features, skip_hash=True)
    check("--skip-hash 之後，純靠雜湊距離判定相似、檔名無規律的兩張不會被分成一組",
          len(groups) == 0, groups)


def main():
    tests = [
        test_original_vs_edited,
        test_grayscale_vs_color,
        test_burst_not_deleted,
        test_same_timestamp_is_not_burst,
        test_near_tie_falls_back_to_size,
        test_exact_tie_goes_to_review,
        test_family_key_matches_effect_suffix,
        test_family_key_matches_copy_suffix,
        test_family_key_does_not_merge_bare_sequence_numbers,
        test_pairs_csv_merges_fragmented_groups,
        test_groups_csv_dedups_repeated_rows,
        test_filename_extra_info_keeps_both,
        test_extract_feature_reads_real_exif_and_saturation,
        test_action_to_labels_mapping,
        test_build_exiftool_args_shape,
        test_write_labels_reports_missing_exiftool_without_crashing,
        test_combine_groups_runs_family_and_hash_together_by_default,
        test_combine_groups_no_duplicate_when_family_and_hash_agree,
        test_combine_groups_merges_pairs_csv_with_automatic_sources,
        test_combine_groups_skip_hash_only_uses_family,
        test_hash_grouping_defaults_to_same_folder_only,
        test_hash_grouping_still_finds_matches_within_same_folder,
        test_write_excel_does_not_leave_thumbnail_beside_source_photo,
        test_write_excel_cleans_up_temp_dir_even_if_thumbnail_fails,
    ]
    print(f"執行 {len(tests)} 個測試案例...\n")
    for t in tests:
        print(f"[{t.__name__}]")
        t()
        print()

    print("=" * 50)
    print(f"通過：{len(PASS)}　失敗：{len(FAIL)}")
    if FAIL:
        print("失敗的測試：")
        for name in FAIL:
            print(f"  - {name}")
        sys.exit(1)
    print("全部通過 ✅")


if __name__ == "__main__":
    main()
