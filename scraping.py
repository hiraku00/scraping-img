import requests
from bs4 import BeautifulSoup
import re
import json
from urllib.parse import urljoin, urlparse
import time
import logging
import argparse
from typing import Optional, Tuple, Any
import os
from io import BytesIO
from PIL import Image as PILImage
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import WebDriverException, TimeoutException

# --- 設定 ---
HEADERS = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36'}
DEFAULT_IMAGE_WIDTH = 100
DEBUG_LOG_FILE = "scraping_debug.log"

# --- Selenium Only ドメインリスト ---
SELENIUM_ONLY_DOMAINS = [
    "ebay.com",
    "mercari.com",
]

# --- 固定する列ヘッダー名 ---
URL_HEADER_NAME = "URL"
IMAGE_URL_HEADER_NAME = "(work)画像URL"

# --- WebDriverManager クラス ---
class WebDriverManager:
    def __init__(self):
        self.options = self._default_options()
        self.driver: Optional[webdriver.Chrome] = None

    def _default_options(self):
        options = Options()
        logging.debug("WebDriver オプション初期化")
        options.add_argument("--headless")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--log-level=3")
        options.add_argument(f"user-agent={HEADERS['User-Agent']}")
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        return options

    def __enter__(self) -> Optional[webdriver.Chrome]:
        t_start = time.time()
        print("WebDriverを初期化しています (ヘッドレスモード)...")
        logging.info("WebDriverを初期化しています (ヘッドレスモード)...")
        try:
            self.driver = webdriver.Chrome(options=self.options)
            print("WebDriverの準備が完了しました。")
            logging.info(f"WebDriverの準備が完了しました。({time.time() - t_start:.3f}s)")
            return self.driver
        except WebDriverException as e:
            logging.error(f"WebDriverException: WebDriver準備失敗: {e}", exc_info=True)
            print(f"エラー: WebDriver準備失敗: {e}")
            if "cannot find Chrome binary" in str(e).lower(): print(">>> Chromeブラウザ本体が見つかりません...")
            elif "session not created" in str(e).lower() and "this version of chromedriver only supports chrome version" in str(e).lower(): print(">>> ChromeDriverのバージョンとChrome本体のバージョンが一致していません...")
            else: print(">>> ChromeDriverの準備またはChromeの起動に問題...")
        except Exception as e:
            logging.error(f"予期せぬエラー: WebDriver準備中: {e}", exc_info=True)
            print(f"エラー: WebDriver準備中: {e}")
        return None

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.driver:
            t_start = time.time()
            try:
                self.driver.quit()
                print("WebDriverを終了しました。")
                logging.info(f"WebDriverを終了しました。({time.time() - t_start:.3f}s)")
            except Exception as e:
                logging.error(f"WebDriver終了中のエラー: {e}", exc_info=True)
                print(f"エラー: WebDriver終了中: {e}")
        self.driver = None

# --- HTML/JSON 解析関数 ---
def extract_meta_property(soup: BeautifulSoup, property_name: str) -> Optional[str]:
    tag = soup.find('meta', property=property_name)
    return tag['content'] if tag and tag.get('content') else None

def extract_meta_name(soup: BeautifulSoup, name: str) -> Optional[str]:
    tag = soup.find('meta', attrs={'name': name})
    return tag['content'] if tag and tag.get('content') else None

def find_image_in_json(json_obj: Any) -> Optional[str]:
    if not json_obj: return None
    if isinstance(json_obj, dict):
        if '@graph' in json_obj and isinstance(json_obj['@graph'], list):
            image = find_image_in_json(json_obj['@graph'])
            if image: return image
        if 'image' in json_obj:
            image_prop = json_obj['image']
            if isinstance(image_prop, str): return image_prop
            elif isinstance(image_prop, list) and len(image_prop) > 0:
                if isinstance(image_prop[0], str): return image_prop[0]
                elif isinstance(image_prop[0], dict) and image_prop[0].get('url'): return image_prop[0]['url']
            elif isinstance(image_prop, dict) and image_prop.get('url'): return image_prop['url']
    elif isinstance(json_obj, list):
        for item in json_obj:
            image = find_image_in_json(item)
            if image: return image
    return None

def extract_json_ld_image(soup: BeautifulSoup) -> Optional[str]:
    scripts = soup.find_all('script', type='application/ld+json')
    for script in scripts:
        if script.string:
            try:
                json_data = json.loads(script.string)
                image_url = find_image_in_json(json_data)
                if image_url: return image_url
            except Exception as e: logging.warning(f"JSON-LD解析エラー: {e}")
    return None

def convert_to_absolute_path(base_url: str, target_path: str) -> str:
    if not target_path or target_path.startswith(('http://', 'https://')): return target_path or ""
    if target_path.startswith('//'): return f"{urlparse(base_url).scheme}:{target_path}"
    try: return urljoin(base_url, target_path)
    except Exception as e: logging.error(f"絶対パス変換失敗: {e}"); return target_path

# HTML解析コアロジック
def parse_html_for_image(html_content: str, base_url: str) -> Optional[str]:
    if not html_content: return None
    logging.debug(f"parse_html: Start parsing for {base_url}")
    t_start = time.time()
    try:
        soup = BeautifulSoup(html_content, 'html.parser')
        logging.debug(f"BeautifulSoup parsing completed in {time.time() - t_start:.3f}s")
    except Exception as e: logging.error(f"BeautifulSoup parsing failed: {e}", exc_info=True); return None
    domain = urlparse(base_url).netloc.lower() if base_url else ""

    image_url = extract_meta_property(soup, 'og:image')
    if image_url: logging.info(f"Found og:image"); return convert_to_absolute_path(base_url, image_url)

    image_url = extract_meta_name(soup, 'twitter:image')
    if image_url: logging.info(f"Found twitter:image"); return convert_to_absolute_path(base_url, image_url)

    if "mercari.com" in domain:
        next_data_script = soup.find('script', id='__NEXT_DATA__', type='application/json')
        if next_data_script and next_data_script.string:
            try:
                next_data = json.loads(next_data_script.string)
                photos = next_data.get('props', {}).get('pageProps', {}).get('item', {}).get('photos', [])
                if photos and isinstance(photos, list) and len(photos) > 0 and isinstance(photos[0], str):
                    logging.info(f"Found image URL in __NEXT_DATA__ (Mercari)"); return convert_to_absolute_path(base_url, photos[0])
            except Exception as e: logging.error(f"Error processing __NEXT_DATA__ JSON: {e}")

        mercari_img_alt = soup.find('img', alt='のサムネイル')
        if mercari_img_alt and mercari_img_alt.get('src') and 'static.mercdn.net' in mercari_img_alt['src']:
            img_src = mercari_img_alt['src']; logging.info(f'Found specific img by alt (mercari - fallback)'); return convert_to_absolute_path(base_url, img_src.split("?")[0])
        mercari_pattern = re.compile(r'https://static\.mercdn\.net/item/detail/orig/photos/[^"\']+?')
        img_tag_src = soup.find('img', src=mercari_pattern)
        if img_tag_src: img_src = img_tag_src['src']; logging.info(f"Found specific img src (mercari pattern - fallback)"); return convert_to_absolute_path(base_url, img_src.split("?")[0])

    image_url = extract_json_ld_image(soup)
    if image_url: logging.info(f"Found JSON-LD image"); return convert_to_absolute_path(base_url, image_url)

    if "amazon" in domain:
        main_image_container = soup.find(id='imgTagWrapperId') or soup.find(id='ivLargeImage') or soup.find(id='landingImage')
        if main_image_container:
            main_img = main_image_container.find('img')
            if main_img and main_img.get('src'):
                potential_src = main_img['src']
                if not potential_src.startswith("data:image") and "captcha" not in potential_src.lower():
                    logging.info(f"Found potential main image via ID (Amazon)"); return convert_to_absolute_path(base_url, potential_src.split("?")[0])

    checked_sources = set()
    image_url = None
    for img in soup.find_all('img'):
        potential_src = img.get('src')
        if not potential_src: continue
        absolute_src = convert_to_absolute_path(base_url, potential_src)
        exclude_patterns = [".gif", ".svg", "ads", "icon", "logo", "sprite", "avatar", "spinner", "loading", "pixel", "fls-fe.amazon", "transparent", "spacer", "dummy", "captcha"]
        exclude_extensions = ['.php', '.aspx', '.jsp']
        if (absolute_src and
                absolute_src not in checked_sources and
                not absolute_src.startswith("data:image") and
                not any(pat in absolute_src.lower() for pat in exclude_patterns) and
                not any(absolute_src.lower().endswith(ext) for ext in exclude_extensions) and
                len(absolute_src) > 10):
            checked_sources.add(absolute_src)
            if "/thumb/" not in absolute_src and "favicon" not in absolute_src:
                logging.info(f"Found generic img src (fallback)")
                image_url = absolute_src.split("?")[0]
                break
    logging.debug(f"parse_html completed in {time.time() - t_start:.3f}s")
    return image_url

# Seleniumでの画像URL取得
def get_image_url_from_url_with_selenium(driver: webdriver.Chrome, url: str) -> Optional[str]:
    image_url = None
    try:
        logging.info(f"Attempting to fetch URL with Selenium: {url}")
        t_get = time.time()
        driver.set_page_load_timeout(30)
        driver.get(url)
        logging.debug(f"Selenium driver.get() completed in {time.time() - t_get:.3f}s")
        t_sleep = time.time(); wait_seconds = 0.5; logging.info(f"Waiting for {wait_seconds} seconds...")
        time.sleep(wait_seconds); logging.debug(f"Selenium time.sleep() completed in {time.time() - t_sleep:.3f}s")
        t_source = time.time(); page_source = driver.page_source; current_url = driver.current_url
        logging.debug(f"Selenium driver.page_source completed in {time.time() - t_source:.3f}s")
        logging.info(f"Successfully fetched page source with Selenium (final URL: {current_url})")
        image_url = parse_html_for_image(page_source, current_url)
        if image_url: logging.info("Found image URL using Selenium.")
        else: logging.warning("Could not find image URL even with Selenium.")
    except TimeoutException: logging.error(f"Selenium page load timed out for URL: {url}")
    except WebDriverException as e: logging.error(f"Selenium WebDriver error for URL {url}: {e}", exc_info=True)
    except Exception as e: logging.error(f"Unexpected error during Selenium processing for URL {url}: {e}", exc_info=True)
    return image_url

# 画像URL取得メイン関数 (ハイブリッド, Selenium Onlyドメイン判定)
def get_image_url_from_url(url: str, row_index_for_debug: int, driver: Optional[webdriver.Chrome] = None) -> Tuple[Optional[str], Optional[str]]:
    final_image_url = None
    error_message = None
    request_timeout_seconds = 20
    parsed_url = urlparse(url)
    domain = parsed_url.netloc.lower()

    use_selenium_directly = any(d in domain for d in SELENIUM_ONLY_DOMAINS)

    if use_selenium_directly:
        logging.info(f"Domain '{domain}' is in SELENIUM_ONLY_DOMAINS. Using Selenium directly for {url}")
        if driver:
            t_sel_start = time.time()
            final_image_url = get_image_url_from_url_with_selenium(driver, url)
            logging.debug(f"get_image_url_from_url_with_selenium (Direct) completed in {time.time() - t_sel_start:.3f}s")
            if not final_image_url: error_message = "画像が見つかりません(Sel-Direct)"
        else:
            logging.warning(f"Selenium Only Domain {url}, but Selenium driver is not available.")
            error_message = "画像が見つかりません(NoDriver)"
        return final_image_url, error_message

    logging.info(f"Attempting to fetch URL with requests: {url}")
    t_req_start = time.time(); response = None; html_content = ""; base_url = url
    try:
        response = requests.get(url, headers=HEADERS, timeout=request_timeout_seconds, allow_redirects=True)
        logging.debug(f"requests.get() completed in {time.time() - t_req_start:.3f}s")
        response.raise_for_status(); response.encoding = response.apparent_encoding
        html_content = response.text; base_url = response.url
        logging.info(f"--- URL: {url} (Final: {base_url}) --- Response Code: {response.status_code}")
        final_image_url = parse_html_for_image(html_content, base_url)
        if not final_image_url: error_message = "画像が見つかりません(Req)"; logging.warning(f"Image not found with requests for: {url}")
    except requests.exceptions.Timeout: error_message = f"タイムアウト(>{request_timeout_seconds}s)(Req)"; logging.warning(f"Requests Timeout (>{request_timeout_seconds}s) in {time.time() - t_req_start:.3f}s for {url}.")
    except requests.exceptions.RequestException as e: status_code = e.response.status_code if e.response is not None else "N/A"; error_message = f"アクセス失敗(Code:{status_code})(Req)"; logging.error(f"Requests Access Error ({time.time() - t_req_start:.3f}s) {url}: {e}")
    except Exception as e: error_message = f"エラー(Req): {str(e)[:50]}"; logging.error(f"Requests URL Processing Error ({time.time() - t_req_start:.3f}s) {url}: {e}", exc_info=True)

    if not final_image_url and driver:
        logging.info(f"Requests failed or timed out for {url}. Retrying with Selenium...")
        t_sel_start = time.time()
        selenium_image_url = get_image_url_from_url_with_selenium(driver, url)
        logging.debug(f"get_image_url_from_url_with_selenium completed in {time.time() - t_sel_start:.3f}s")
        if selenium_image_url: final_image_url = selenium_image_url; error_message = None
        else: error_message = "画像が見つかりません(Sel)" if "タイムアウト" in str(error_message) else f"{error_message} / 画像が見つかりません(Sel)"
    elif not final_image_url and not driver: logging.warning(f"Requests failed for {url}, and Selenium retry is disabled or driver is unavailable.")

    return final_image_url, error_message

# 画像ダウンロード＆準備関数
def download_and_prepare_image(image_url: str, target_width: int) -> Optional[Tuple[BytesIO, int, int]]:
    t_dl_start = time.time()
    try:
        logging.debug(f"download_prepare: Starting download for {image_url}")
        img_response = requests.get(image_url, stream=True, timeout=15); img_response.raise_for_status()
        logging.debug(f"download_prepare: Download completed in {time.time() - t_dl_start:.3f}s. Status: {img_response.status_code}")
        t_proc_start = time.time(); content_type = img_response.headers.get('content-type')
        if not content_type or not content_type.lower().startswith('image/'): logging.warning(f"非画像コンテンツ ({content_type}): {image_url}"); return None
        img_data = BytesIO(img_response.content)
        if img_data.getbuffer().nbytes == 0: logging.warning(f"空の画像データ: {image_url}"); return None
        with PILImage.open(img_data) as img:
            img_copy = img.copy();
            if img_copy.mode == 'P': img_copy = img_copy.convert('RGBA')
            elif img_copy.mode == 'CMYK': img_copy = img_copy.convert('RGB')
            elif img_copy.mode == 'LA': img_copy = img_copy.convert('RGBA')
            original_width, original_height = img_copy.size
            if original_width <= 0 or original_height <= 0: logging.warning(f"無効画像サイズ: {image_url}"); return None
            aspect_ratio = original_height / original_width; target_height = max(1, int(target_width * aspect_ratio))
            img_resized = img_copy.resize((target_width, target_height), PILImage.Resampling.LANCZOS)
            output_buffer = BytesIO()
            save_format = img.format if img.format and img.format.upper() in ['JPEG', 'PNG', 'BMP', 'TIFF'] else 'PNG'
            if save_format == 'GIF': save_format = 'PNG'
            if save_format == 'JPEG' and img_resized.mode == 'RGBA': img_resized = img_resized.convert('RGB')
            img_resized.save(output_buffer, format=save_format, quality=85 if save_format == 'JPEG' else None)
        output_buffer.seek(0);
        if output_buffer.closed: logging.error(f"BytesIO閉鎖済み(保存後): {image_url}"); return None
        logging.debug(f"download_prepare: Image processing completed in {time.time() - t_proc_start:.3f}s")
        return output_buffer, target_width, target_height
    except requests.exceptions.RequestException as e: logging.error(f"画像DLエラー {image_url}: {e}", exc_info=True); return None
    except PILImage.UnidentifiedImageError: logging.error(f"画像形式認識不可: {image_url}", exc_info=True); return None
    except Exception as e: logging.error(f"画像処理エラー {image_url}: {e}", exc_info=True); return None

# --- メイン実行ブロック ---
if __name__ == "__main__":
    overall_start_time = time.time()
    parser = argparse.ArgumentParser(description='Excel内のURLから画像を取得 (requests + Selenium 再試行, デフォルト有効・ヘッドレス)')
    parser.add_argument('input_file', help='処理対象Excelファイルパス (.xlsx)')
    parser.add_argument('--process_all', action='store_true', help='空URLで処理中断')
    parser.add_argument('--sheet_name', default=0, help='シート名 or インデックス (0始まり)')
    parser.add_argument('--image_width', type=int, default=DEFAULT_IMAGE_WIDTH, help=f'埋込画像幅(px, デフォルト: {DEFAULT_IMAGE_WIDTH})')
    parser.add_argument('--sleep', type=float, default=1.0, help='各URL処理後の待機時間(秒, デフォルト: 1.0)')
    parser.add_argument('--debug', action='store_true', help='デバッグレベルのログをファイルに出力する')
    args = parser.parse_args()

    # --- ログ設定 ---
    log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - [%(funcName)s] - %(message)s')
    logger = logging.getLogger()
    logger.handlers.clear()
    console_handler = logging.StreamHandler(); console_handler.setFormatter(log_formatter)
    console_log_level = logging.INFO if args.debug else logging.WARNING
    console_handler.setLevel(console_log_level)
    logger.addHandler(console_handler)
    if args.debug:
        try:
            if os.path.exists(DEBUG_LOG_FILE): os.remove(DEBUG_LOG_FILE)
            file_handler = logging.FileHandler(DEBUG_LOG_FILE, mode='w', encoding='utf-8')
            file_handler.setFormatter(log_formatter); file_handler.setLevel(logging.DEBUG)
            logger.addHandler(file_handler); logger.setLevel(logging.DEBUG)
            print(f"デバッグログを有効にし、{DEBUG_LOG_FILE} に出力します。")
        except Exception as e: logging.error(f"デバッグログファイル準備失敗: {e}"); logger.setLevel(logging.INFO)
    else:
        logger.setLevel(logging.WARNING)
    logging.getLogger("urllib3").setLevel(logging.WARNING)
    logging.getLogger("selenium").setLevel(logging.WARNING)
    # --- ログ設定ここまで ---

    # --- Excelファイル処理 ---
    if not os.path.isfile(args.input_file): print(f"エラー: ファイル '{args.input_file}' 未検出"); exit(1)
    if not args.input_file.lower().endswith('.xlsx'): print("エラー: 入力は .xlsx ファイルのみ"); exit(1)
    workbook = None
    try:
        t_excel_load = time.time(); print(f"入力ファイル '{args.input_file}' を読み込み処理開始...")
        try: workbook = openpyxl.load_workbook(args.input_file)
        except Exception as load_e: print(f"\nエラー: Excelファイル読込失敗: {load_e}"); logging.error(f"Excel読込失敗: {args.input_file}", exc_info=True); exit(1)
        logging.info(f"Excelファイル読み込み完了 ({time.time() - t_excel_load:.3f}s)")
        sheet_name = args.sheet_name if isinstance(args.sheet_name, str) else workbook.sheetnames[args.sheet_name]
        if sheet_name not in workbook.sheetnames: raise ValueError(f"シート名 '{sheet_name}' が見つかりません。")
        sheet = workbook[sheet_name]; print(f"処理シート: '{sheet.title}'")
        header_row_index = 1; headers = {cell.value: cell.column for cell in sheet[header_row_index] if cell.value is not None}

        # --- 列インデックスの取得とチェック (固定ヘッダー名を使用) ---
        # 固定ヘッダー名がシートに存在するかチェック
        required_headers = {URL_HEADER_NAME, IMAGE_URL_HEADER_NAME}
        missing_headers = required_headers - set(headers.keys())
        if missing_headers:
            raise ValueError(f"必須列ヘッダー未検出: {missing_headers}. シート上のヘッダー: {list(headers.keys())}")

        # 固定ヘッダー名から列インデックスを取得
        url_col_idx = headers[URL_HEADER_NAME]
        img_url_col_idx = headers[IMAGE_URL_HEADER_NAME]
        # 画像埋め込み列はE列 (列番号5) に固定
        img_embed_col_idx = 5
        # ハイフンを入れる列はD列 (列番号4) に固定
        hyphen_col_idx = 4

        print(f"使用列: URL='{URL_HEADER_NAME}' ({get_column_letter(url_col_idx)}), "
              f"ImgURL='{IMAGE_URL_HEADER_NAME}' ({get_column_letter(img_url_col_idx)}), "
              f"ImgEmbed=E, Hyphen=D")

        # --- 処理対象行数のカウント ---
        processed_count = 0; valid_url_found = False; total_rows_to_process = 0
        for row_idx in range(header_row_index + 1, sheet.max_row + 1):
            url_val = sheet.cell(row=row_idx, column=url_col_idx).value;
            if url_val and str(url_val).strip(): total_rows_to_process += 1
        print(f"処理対象のURLを含む行数: {total_rows_to_process}")

        # ★★★ 注意: 既存の画像をすべて削除します ★★★
        if sheet._images:
             print("注意: シート上の既存の画像をすべて削除します...")
             logging.warning("シート上の既存の画像をすべて削除します。")
             sheet._images = []
        # ★★★ ここまで追加 ★★★

        print("Selenium再試行をデフォルトで有効にします (ヘッドレスモード)。WebDriverを準備します...")
        with WebDriverManager() as driver:
            if driver is None: print("警告: WebDriverの初期化に失敗。Seleniumでの再試行は行われません。")
            row_processed_count = 0
            # --- メインループ ---
            for row_index in range(header_row_index + 1, sheet.max_row + 1):
                # --- セルオブジェクト取得 ---
                url_cell = sheet.cell(row=row_index, column=url_col_idx)
                url = str(url_cell.value).strip() if url_cell.value is not None else ""
                img_url_cell = sheet.cell(row=row_index, column=img_url_col_idx)
                img_embed_cell = sheet.cell(row=row_index, column=img_embed_col_idx) # E列
                hyphen_cell_d = sheet.cell(row=row_index, column=hyphen_col_idx)   # ★ D列のセルを取得

                # --- セルクリア処理 ---
                img_url_cell.value = None
                img_url_cell.hyperlink = None # ★ ハイパーリンクもクリア
                img_embed_cell.value = None
                hyphen_cell_d.value = None    # ★ D列もクリア

                # --- URLチェック ---
                if not url:
                    if args.process_all: print(f"\n行 {row_index}: URL空のため中断"); break
                    else: logging.debug(f"Row {row_index}: Skipping empty URL."); continue
                if not url.lower().startswith(('http://', 'https://')):
                    logging.warning(f"行 {row_index}: 無効URL形式: {url}"); img_url_cell.value = "無効なURL"
                    # 無効なURLの場合でもD列に印をつけるならここに hyphen_cell_d.value = "無効" など追加可能
                    if args.process_all: print(f"\n行 {row_index}: 無効URLのため中断"); break
                    else: continue

                # --- 有効なURLの場合の処理開始 ---
                valid_url_found = True; row_processed_count += 1

                # ★★★ D列にハイフンを設定 ★★★
                hyphen_cell_d.value = "-"

                # --- 処理状況表示 ---
                print(f"\r処理中: {row_processed_count}/{total_rows_to_process} 件目 ({row_index}行目) - {url[:50]}...", end="", flush=True)

                # --- 画像URL取得 ---
                t_geturl_start = time.time()
                image_url, error_message = get_image_url_from_url(url, row_index - 1, driver)
                logging.debug(f"Row {row_index}: get_image_url_from_url finished in {time.time() - t_geturl_start:.3f}s")

                # --- 画像URL取得結果の処理 ---
                if image_url:
                    img_url_cell.value = image_url
                    img_url_cell.hyperlink = image_url
                    img_url_cell.style = "Hyperlink"

                    # --- 画像ダウンロードと準備 ---
                    t_dlprep_start = time.time(); image_result = download_and_prepare_image(image_url, args.image_width)
                    logging.debug(f"Row {row_index}: download_and_prepare_image finished in {time.time() - t_dlprep_start:.3f}s")
                    if image_result:
                        # --- 画像埋め込み (E列固定) ---
                        image_data_buffer, img_width, img_height = image_result; t_embed_start = time.time()
                        try:
                            if not image_data_buffer.closed:
                                img_for_excel = OpenpyxlImage(image_data_buffer); img_for_excel.width = img_width; img_for_excel.height = img_height
                                required_row_height = img_height * 0.75 + 2
                                if sheet.row_dimensions[row_index].height is None or sheet.row_dimensions[row_index].height < required_row_height: sheet.row_dimensions[row_index].height = required_row_height

                                col_letter = 'E' # 固定
                                required_col_width = img_width / 7.0 + 2
                                current_width = sheet.column_dimensions[col_letter].width
                                if current_width is None or current_width < required_col_width: sheet.column_dimensions[col_letter].width = required_col_width

                                cell_anchor = f"E{row_index}" # 固定
                                img_embed_cell.alignment = Alignment(horizontal='center', vertical='center')
                                sheet.add_image(img_for_excel, cell_anchor); logging.info(f"行 {row_index}: 画像埋込完了 -> {cell_anchor}")
                            else: logging.error(f"行 {row_index}: BytesIO閉鎖済"); img_embed_cell.value = "内部エラー(Buffer Closed)"
                        except ValueError as ve: logging.error(f"行 {row_index}: 画像埋込ValueError: {ve}", exc_info=True); img_embed_cell.value = f"画像形式エラー? ({ve})"
                        except Exception as e: logging.error(f"行 {row_index}: 画像埋込 予期せぬエラー: {e}", exc_info=True); img_embed_cell.value = "画像埋込エラー"
                        logging.debug(f"Row {row_index}: Excel embedding finished in {time.time() - t_embed_start:.3f}s")
                    else:
                        # 画像ダウンロード/処理失敗
                        logging.warning(f"行 {row_index}: 画像データ準備失敗 URL: {image_url}");
                        img_embed_cell.value = "画像DL/処理失敗"
                else:
                    # 画像URL取得失敗
                    img_url_cell.value = error_message if error_message else "取得エラー"
                    logging.error(f"行 {row_index}: 画像URL取得失敗 - {error_message}")
                    # 画像URL取得失敗でもD列には "-" が入っている

                # --- 処理状況表示と待機 ---
                current_elapsed_time = time.time() - overall_start_time
                print(f"\r処理完了: {row_processed_count}/{total_rows_to_process} 件目 ({row_index}行目) - {url[:50]}... (経過時間: {current_elapsed_time:.2f} 秒)          ", flush=True)
                t_sleep_start = time.time(); time.sleep(args.sleep); logging.debug(f"Row {row_index}: Post-URL sleep completed in {time.time() - t_sleep_start:.3f}s")
            print() # 改行

        if not valid_url_found and total_rows_to_process > 0 : print("有効なURL未検出")
        elif row_processed_count > 0:
            overall_elapsed_time = time.time() - overall_start_time
            print(f"\n処理完了 (処理URL: {row_processed_count} 件 / 全体時間: {overall_elapsed_time:.2f} 秒)")
        print(f"変更を '{args.input_file}' に保存中...")
        t_save_start = time.time()
        if workbook is None: raise RuntimeError("ワークブックオブジェクトが無効")
        try: workbook.save(args.input_file); print("保存完了")
        except PermissionError: print(f"\nエラー: '{args.input_file}' 書込権限なし")
        except Exception as save_e: print(f"\nエラー: ファイル保存中に問題発生: {save_e}"); logging.error(f"ファイル保存エラー", exc_info=True)
        logging.info(f"Excel file saving completed in {time.time() - t_save_start:.3f}s")
    except FileNotFoundError: print(f"エラー: ファイル '{args.input_file}' 未検出")
    except ValueError as ve: print(f"設定/ファイルエラー: {ve}")
    except ImportError as ie:
        if 'selenium' in str(ie).lower(): print("エラー: Seleniumライブラリ未インストール")
        else: print(f"エラー: 必要なライブラリ未インストール: {ie}")
    except RuntimeError as rte: print(f"\n内部エラー: {rte}")
    except KeyboardInterrupt: print("\n処理が中断されました。")
    except Exception as e: print(f"\n予期せぬエラー発生: {e}"); logging.exception("予期せぬエラー")
    finally:
        if workbook:
            try: workbook.close(); logging.info("ワークブックを閉じました。")
            except Exception as close_e: logging.error(f"ワークブッククローズ中のエラー: {close_e}")
        if 'overall_start_time' in locals():
            overall_final_elapsed_time = time.time() - overall_start_time
            print(f"\nスクリプト実行時間: {overall_final_elapsed_time:.2f} 秒")
