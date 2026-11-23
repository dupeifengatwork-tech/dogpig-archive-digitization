import os
import logging
import warnings
from PIL import Image

# 忽略 urllib3 與系統 Warning
warnings.filterwarnings("ignore")
logging.getLogger("ppocr").setLevel(logging.ERROR)

from paddleocr import PaddleOCR

def search_keywords_with_paddle(folder_path, target_keywords):
    print("正在載入 PaddleOCR 繁體中文模型...")
    # 移除 max_side_limit 避免新版初始化報錯
    ocr = PaddleOCR(
        use_textline_orientation=True, 
        lang='chinese_cht'
    )

    supported_exts = ('.jpg', '.jpeg', '.JPG', '.JPEG')
    image_files = [f for f in os.listdir(folder_path) if f.endswith(supported_exts)]

    if not image_files:
        print(f"在 '{folder_path}' 資料夾中沒有找到 JPG 圖檔！")
        return

    print(f"找到 {len(image_files)} 張圖片，開始辨識與比對...\n")

    matched_files = {}

    for idx, filename in enumerate(image_files, 1):
        image_path = os.path.join(folder_path, filename)
        temp_path = os.path.join(folder_path, f"_temp_{filename}")

        print(f"[{idx}/{len(image_files)}] 正在掃描: {filename} ...", end="", flush=True)

        try:
            # 用 PIL 將巨型圖片長邊預先縮放到 1200px 以內，解決卡死問題
            with Image.open(image_path) as img:
                img.thumbnail((1200, 1200))
                img.convert('RGB').save(temp_path, "JPEG")

            # 進行 OCR 辨識
            results = ocr.ocr(temp_path)

            # 刪除臨時檔
            if os.path.exists(temp_path):
                os.remove(temp_path)

            extracted_texts = []
            if results:
                for res in results:
                    if isinstance(res, dict) and 'rec_polys' in res:
                        rec_texts = res.get('rec_text', [])
                        extracted_texts.extend(rec_texts)
                    elif isinstance(res, list):
                        for line in res:
                            if isinstance(line, list) and len(line) >= 2:
                                extracted_texts.append(line[1][0])
                    elif hasattr(res, 'str'):
                        extracted_texts.append(str(res))

            full_text = " ".join(extracted_texts)

            # 檢查關鍵字
            found_kws = [kw for kw in target_keywords if kw in full_text]

            if found_kws:
                matched_files[filename] = {
                    'keywords': found_kws,
                    'full_text': full_text
                }
                print(f" 👉 找到關鍵字: {found_kws}")
            else:
                print(" 無匹配")

        except Exception as e:
            if os.path.exists(temp_path):
                os.remove(temp_path)
            print(f" ❌ 出錯: {e}")

    print("\n" + "="*50)
    print(f"比對完成！符合條件的檔案共有 {len(matched_files)} 個：")
    print("="*50)

    for filename, data in matched_files.items():
        print(f"\n檔名：{filename}")
        print(f"命中關鍵字：{', '.join(data['keywords'])}")
        print(f"內文預覽：{data['full_text'][:100]}...")

if __name__ == "__main__":
    # 用法：python 00_search_paddle.py <圖片資料夾> <關鍵字1> [關鍵字2 ...]
    import sys
    if len(sys.argv) < 3:
        sys.exit("用法：python 00_search_paddle.py /path/to/zine-scans 關鍵字 [關鍵字 ...]")
    search_keywords_with_paddle(sys.argv[1], sys.argv[2:])
