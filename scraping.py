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
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
HEADERS = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36'}
DEFAULT_IMAGE_WIDTH = 100

# --- WebDriverManager クラス ---
class WebDriverManager:
    """WebDriverを自動管理し、コンテキストマネージャーで提供するクラス"""
    def __init__(self):
        self.options = self._default_options()
        self.driver: Optional[webdriver.Chrome] = None

    def _default_options(self):
        options = Options()
        options.add_argument("--headless")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--log-level=3")
        options.add_argument(f"user-agent={HEADERS['User-Agent']}")
        return options

    def __enter__(self) -> Optional[webdriver.Chrome]:
        print("WebDriverを初期化しています (ヘッドレスモード)...")
        try:
            self.driver = webdriver.Chrome(options=self.options)
            print("WebDriverの準備が完了しました。")
            return self.driver
        except WebDriverException as e:
            print(f"WebDriverException: WebDriverの準備に失敗しました: {e}")
            # エラー詳細の表示は残す
            if "cannot find Chrome binary" in str(e).lower(): print(">>> Chromeブラウザ本体が見つかりません...")
            elif "session not created" in str(e).lower() and "this version of chromedriver only supports chrome version" in str(e).lower(): print(">>> ChromeDriverのバージョンとChrome本体のバージョンが一致していません...")
            else: print(">>> ChromeDriverの準備またはChromeの起動に問題...")
        except Exception as e: print(f"予期せぬエラー: WebDriverの準備中にエラーが発生しました: {e}")
        return None

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.driver:
            try:
                self.driver.quit()
                print("WebDriverを終了しました。")
            except Exception as e: print(f"WebDriver終了中のエラー: {e}"); logging.error(f"WebDriver終了エラー: {e}")
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
            except json.JSONDecodeError as e: logging.warning(f"JSON-LD解析失敗: {e}")
            except Exception as e: logging.error(f"JSON-LD処理エラー: {e}")
    return None

def convert_to_absolute_path(base_url: str, target_path: str) -> str:
    if not target_path: return ""
    if target_path.startswith(('http://', 'https://')): return target_path
    if target_path.startswith('//'):
        base_scheme = urlparse(base_url).scheme
        return f"{base_scheme}:{target_path}"
    try: return urljoin(base_url, target_path)
    except Exception as e:
        logging.error(f"絶対パス変換失敗 Base: {base_url}, Target: {target_path}, Error: {e}")
        return target_path

# HTML解析コアロジック
def parse_html_for_image(html_content: str, base_url: str) -> Optional[str]:
    if not html_content: return None
    soup = BeautifulSoup(html_content, 'html.parser')
    domain = urlparse(base_url).netloc.lower() if base_url else ""

    # 優先度順に画像URLを探索
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
                    logging.info(f"Found image URL in __NEXT_DATA__ (item.photos): {photos[0]}")
                    return convert_to_absolute_path(base_url, photos[0])
            except Exception as e: logging.error(f"Error processing __NEXT_DATA__ JSON: {e}")
        else: logging.debug("__NEXT_DATA__ script tag not found or empty.")

        mercari_img_alt = soup.find('img', alt='のサムネイル')
        if mercari_img_alt and mercari_img_alt.get('src') and 'static.mercdn.net/item/detail/orig/photos/' in mercari_img_alt['src']:
            img_src = mercari_img_alt['src']; logging.info(f'Found specific img by alt (mercari - fallback): {img_src}'); return convert_to_absolute_path(base_url, img_src.split("?")[0])

        mercari_pattern = re.compile(r'https://static\.mercdn\.net/item/detail/orig/photos/[^"\']+?')
        img_tag_src = soup.find('img', src=mercari_pattern)
        if img_tag_src: img_src = img_tag_src['src']; logging.info(f"Found specific img src (mercari pattern - fallback): {img_src}"); return convert_to_absolute_path(base_url, img_src.split("?")[0])

    image_url = extract_json_ld_image(soup)
    if image_url: logging.info(f"Found JSON-LD image"); return convert_to_absolute_path(base_url, image_url)

    if "amazon" in domain:
        main_image_container = soup.find(id='imgTagWrapperId') or soup.find(id='ivLargeImage') or soup.find(id='landingImage')
        if main_image_container:
            main_img = main_image_container.find('img')
            if main_img and main_img.get('src'):
                potential_src = main_img['src']
                if not potential_src.startswith("data:image") and not potential_src.endswith("1._AC_UL.") and "captcha" not in potential_src.lower():
                    logging.info(f"Found potential main image via ID (Amazon): {potential_src}"); return convert_to_absolute_path(base_url, potential_src.split("?")[0])

    checked_sources = set()
    for img in soup.find_all('img'):
        potential_src = img.get('src')
        if not potential_src: continue
        absolute_src = convert_to_absolute_path(base_url, potential_src)
        exclude_patterns = [".gif", ".svg", "ads", "icon", "logo", "sprite", "avatar", "spinner", "loading", "pixel", "fls-fe.amazon", "transparent", "spacer", "dummy", "captcha"]
        exclude_extensions = ['.php', '.aspx', '.jsp']
        if absolute_src and absolute_src not in checked_sources and \
            not absolute_src.startswith("data:image") and \
            not any(pat in absolute_src.lower() for pat in exclude_patterns) and \
            not any(absolute_src.lower().endswith(ext) for ext in exclude_extensions) and \
            len(absolute_src) > 10:
            checked_sources.add(absolute_src)
            if "/thumb/" not in absolute_src and "favicon" not in absolute_src:
                logging.info(f"Found generic img src (fallback): {absolute_src}")
                return absolute_src.split("?")[0]

    return None # 見つからなければ None

# --- Seleniumでの画像URL取得 ---
def get_image_url_from_url_with_selenium(driver: webdriver.Chrome, url: str) -> Optional[str]:
    """Seleniumを使って画像URLを取得する"""
    try:
        logging.info(f"Attempting to fetch URL with Selenium: {url}")
        driver.set_page_load_timeout(30)
        driver.get(url)
        wait_seconds = 1
        logging.info(f"Waiting for {wait_seconds} seconds...")
        time.sleep(wait_seconds)
        page_source = driver.page_source
        current_url = driver.current_url
        logging.info(f"Successfully fetched page source with Selenium (final URL: {current_url})")

        image_url = parse_html_for_image(page_source, current_url)

        if image_url:
            logging.info("Found image URL using Selenium.")
            return image_url
        else:
            logging.warning("Could not find image URL even with Selenium.")
            # HTMLデバッグ出力は削除
            return None
    except TimeoutException: logging.error(f"Selenium page load timed out for URL: {url}"); return None
    except WebDriverException as e: logging.error(f"Selenium WebDriver error for URL {url}: {e}"); return None
    except Exception as e: logging.error(f"Unexpected error during Selenium processing for URL {url}: {e}"); return None

# --- 画像URL取得メイン関数 (ハイブリッド) ---
def get_image_url_from_url(url: str, row_index_for_debug: int, driver: Optional[webdriver.Chrome] = None) -> Tuple[Optional[str], Optional[str]]:
    """requestsで試行し、失敗した場合にSeleniumで再試行して画像URLを取得"""
    final_image_url = None
    error_message = None

    logging.info(f"Attempting to fetch URL with requests: {url}")
    try:
        response = requests.get(url, headers=HEADERS, timeout=20, allow_redirects=True)
        response.raise_for_status()
        response.encoding = response.apparent_encoding
        html_content = response.text
        base_url = response.url
        logging.info(f"--- URL: {url} (Final: {base_url}) ---")
        logging.info(f"Response Code: {response.status_code}")
        final_image_url = parse_html_for_image(html_content, base_url)
        if not final_image_url:
            logging.warning(f"Image not found with requests for: {url}")
            error_message = "画像が見つかりません(Req)"
    except requests.exceptions.Timeout:
        logging.error(f"Requests Timeout {url}"); error_message = "タイムアウト(Req)"
    except requests.exceptions.RequestException as e:
        logging.error(f"Requests Access Error {url}: {e}")
        status_code = e.response.status_code if e.response is not None else "N/A"
        error_message = f"アクセス失敗(Code:{status_code})(Req)"
    except Exception as e:
        logging.error(f"Requests URL Processing Error {url}: {e}"); error_message = f"エラー(Req): {str(e)[:50]}"
    # HTMLデバッグ出力は削除

    # requests失敗時 かつ driver利用可能な場合のみSelenium再試行
    if not final_image_url and driver:
        logging.info(f"Requests failed or image not found for {url}. Retrying with Selenium...")
        selenium_image_url = get_image_url_from_url_with_selenium(driver, url)
        if selenium_image_url:
            final_image_url = selenium_image_url
            error_message = None # 成功したのでエラーメッセージクリア
        else:
            selenium_error = "画像が見つかりません(Sel)"
            error_message = f"{error_message} / {selenium_error}" if error_message else selenium_error
    elif not final_image_url and not driver:
        logging.warning(f"Requests failed for {url}, and Selenium retry is disabled or driver is unavailable.")

    return final_image_url, error_message

# --- 画像ダウンロード＆準備関数 ---
def download_and_prepare_image(image_url: str, target_width: int) -> Optional[Tuple[BytesIO, int, int]]:
    """画像をダウンロードし、リサイズしてBytesIOで返す"""
    try:
        img_response = requests.get(image_url, stream=True, timeout=15)
        img_response.raise_for_status()
        content_type = img_response.headers.get('content-type')
        if not content_type or not content_type.lower().startswith('image/'):
            logging.warning(f"非画像コンテンツ ({content_type}): {image_url}"); return None
        img_data = BytesIO(img_response.content)
        if img_data.getbuffer().nbytes == 0: logging.warning(f"空の画像データ: {image_url}"); return None
        with PILImage.open(img_data) as img:
            img_copy = img.copy()
            if img_copy.mode == 'P': img_copy = img_copy.convert('RGBA')
            elif img_copy.mode == 'CMYK': img_copy = img_copy.convert('RGB')
            elif img_copy.mode == 'LA': img_copy = img_copy.convert('RGBA')
            original_width, original_height = img_copy.size
            if original_width <= 0 or original_height <= 0: logging.warning(f"無効画像サイズ: {image_url}"); return None
            aspect_ratio = original_height / original_width
            target_height = max(1, int(target_width * aspect_ratio))
            img_resized = img_copy.resize((target_width, target_height), PILImage.Resampling.LANCZOS)
            output_buffer = BytesIO()
            save_format = img.format if img.format and img.format.upper() in ['JPEG', 'PNG', 'BMP', 'TIFF'] else 'PNG'
            if save_format == 'GIF': save_format = 'PNG'
            if save_format == 'JPEG' and img_resized.mode == 'RGBA': img_resized = img_resized.convert('RGB')
            img_resized.save(output_buffer, format=save_format, quality=85 if save_format == 'JPEG' else None)
        output_buffer.seek(0)
        if output_buffer.closed: logging.error(f"BytesIO閉鎖済み(保存後): {image_url}"); return None
        return output_buffer, target_width, target_height
    except requests.exceptions.RequestException as e: logging.error(f"画像DLエラー {image_url}: {e}"); return None
    except PILImage.UnidentifiedImageError: logging.error(f"画像形式認識不可: {image_url}"); return None
    except Exception as e: logging.error(f"画像処理エラー {image_url}: {e}"); return None

# --- メイン実行ブロック (デフォルトハイブリッド + 時間計測) ---
if __name__ == "__main__":
    overall_start_time = time.time()

    # 引数解析 (--use_selenium を削除)
    parser = argparse.ArgumentParser(description='Excel内のURLから画像を取得 (requests + Selenium 再試行, デフォルト有効・ヘッドレス)')
    parser.add_argument('input_file', help='処理対象Excelファイルパス (.xlsx)')
    parser.add_argument('-u', '--url_column', default='URL', help='URL列ヘッダー名')
    parser.add_argument('-i', '--image_url_column', default='(work)画像URL', help='画像URL出力列ヘッダー名')
    parser.add_argument('-p', '--image_embed_column', default='画像', help='画像埋込列ヘッダー名')
    parser.add_argument('--process_all', action='store_true', help='空URLで処理中断')
    parser.add_argument('--sheet_name', default=0, help='シート名 or インデックス (0始まり)')
    parser.add_argument('--image_width', type=int, default=DEFAULT_IMAGE_WIDTH, help=f'埋込画像幅(px, デフォルト: {DEFAULT_IMAGE_WIDTH})')
    parser.add_argument('--sleep', type=float, default=1.0, help='各URL処理後の待機時間(秒, デフォルト: 1.0)')
    args = parser.parse_args()

    if not os.path.isfile(args.input_file): print(f"エラー: ファイル '{args.input_file}' 未検出"); exit(1)
    if not args.input_file.lower().endswith('.xlsx'): print("エラー: 入力は .xlsx ファイルのみ"); exit(1)

    workbook = None
    try:
        print(f"入力ファイル '{args.input_file}' を読み込み処理開始...")
        try: workbook = openpyxl.load_workbook(args.input_file)
        except Exception as load_e: print(f"\nエラー: Excelファイル読込失敗: {load_e}"); logging.error(f"Excel読込失敗: {args.input_file}", exc_info=True); exit(1)

        # シートとヘッダーの処理 (簡略化)
        sheet_name = args.sheet_name if isinstance(args.sheet_name, str) else workbook.sheetnames[args.sheet_name]
        if sheet_name not in workbook.sheetnames: raise ValueError(f"シート名 '{sheet_name}' が見つかりません。")
        sheet = workbook[sheet_name]
        print(f"処理シート: '{sheet.title}'")
        header_row_index = 1
        headers = {cell.value: cell.column for cell in sheet[header_row_index] if cell.value is not None}
        required_cols = {args.url_column, args.image_url_column, args.image_embed_column}
        missing_cols = required_cols - set(headers.keys())
        if missing_cols: raise ValueError(f"必須列ヘッダー未検出: {missing_cols}. 利用可能: {list(headers.keys())}")
        url_col_idx = headers[args.url_column]; img_url_col_idx = headers[args.image_url_column]; img_embed_col_idx = headers[args.image_embed_column]
        print(f"URL:{get_column_letter(url_col_idx)}, ImgURL:{get_column_letter(img_url_col_idx)}, ImgEmbed:{get_column_letter(img_embed_col_idx)}")
        processed_count = 0; valid_url_found = False; total_rows_to_process = sheet.max_row - header_row_index

        # 常に WebDriverManager を使用
        print("Selenium再試行をデフォルトで有効にします (ヘッドレスモード)。WebDriverを準備します...")
        with WebDriverManager() as driver:
            if driver is None:
                print("警告: WebDriverの初期化に失敗。Seleniumでの再試行は行われません。")
                # driver が None でも処理は続行 (get_image_url_from_url内で処理)

            for row_index in range(header_row_index + 1, sheet.max_row + 1):
                url_cell = sheet.cell(row=row_index, column=url_col_idx)
                url = str(url_cell.value).strip() if url_cell.value is not None else ""
                img_url_cell = sheet.cell(row=row_index, column=img_url_col_idx)
                img_embed_cell = sheet.cell(row=row_index, column=img_embed_col_idx)
                img_url_cell.value = None; img_embed_cell.value = None # 事前クリア

                if not url:
                    if args.process_all: print(f"\n行 {row_index}: URL空のため中断"); break
                    else: continue
                if not url.lower().startswith(('http://', 'https://')):
                    logging.warning(f"行 {row_index}: 無効URL形式: {url}"); img_url_cell.value = "無効なURL"
                    if args.process_all: print(f"\n行 {row_index}: 無効URLのため中断"); break
                    else: continue

                valid_url_found = True; processed_count += 1
                print(f"\r処理中: {processed_count}/{total_rows_to_process} 件目 ({row_index}行目) - {url[:50]}...", end="", flush=True)

                # ハイブリッド関数呼び出し (driverインスタンスを渡す)
                image_url, error_message = get_image_url_from_url(url, row_index - 1, driver)

                # 画像埋め込み処理
                if image_url:
                    img_url_cell.value = image_url
                    image_result = download_and_prepare_image(image_url, args.image_width)
                    if image_result:
                        image_data_buffer, img_width, img_height = image_result
                        try:
                            if not image_data_buffer.closed:
                                img_for_excel = OpenpyxlImage(image_data_buffer)
                                img_for_excel.width = img_width; img_for_excel.height = img_height
                                required_row_height = img_height * 0.75 + 2
                                if sheet.row_dimensions[row_index].height is None or sheet.row_dimensions[row_index].height < required_row_height: sheet.row_dimensions[row_index].height = required_row_height
                                col_letter = get_column_letter(img_embed_col_idx)
                                required_col_width = img_width / 7.0 + 2
                                current_width = sheet.column_dimensions[col_letter].width
                                if current_width is None or current_width < required_col_width: sheet.column_dimensions[col_letter].width = required_col_width
                                cell_anchor = f"{col_letter}{row_index}"
                                img_embed_cell.alignment = Alignment(horizontal='center', vertical='center')
                                sheet.add_image(img_for_excel, cell_anchor)
                                logging.info(f"行 {row_index}: 画像埋込完了 -> {cell_anchor}")
                            else: logging.error(f"行 {row_index}: BytesIO閉鎖済 - 埋込スキップ"); img_embed_cell.value = "内部エラー(Buffer Closed)"
                        except ValueError as ve: logging.error(f"行 {row_index}: 画像埋込ValueError: {ve}"); img_embed_cell.value = f"画像形式エラー? ({ve})"
                        except Exception as e: logging.error(f"行 {row_index}: 画像埋込 予期せぬエラー: {e}"); img_embed_cell.value = "画像埋込エラー"
                    else: logging.warning(f"行 {row_index}: 画像データ準備失敗 URL: {image_url}"); img_embed_cell.value = "画像DL/処理失敗"
                else: img_url_cell.value = error_message if error_message else "取得エラー"

                # 累積経過時間表示
                current_elapsed_time = time.time() - overall_start_time
                print(f"\r処理完了: {processed_count}/{total_rows_to_process} 件目 ({row_index}行目) - {url[:50]}... (経過時間: {current_elapsed_time:.2f} 秒)          ", flush=True)

                time.sleep(args.sleep) # 各URL処理後の待機
            print() # ループ終了後改行

        # 処理結果表示と保存
        if not valid_url_found and total_rows_to_process > 0 : print("有効なURL未検出")
        elif processed_count > 0:
            overall_elapsed_time = time.time() - overall_start_time
            print(f"\n処理完了 (処理URL: {processed_count} 件 / 全体時間: {overall_elapsed_time:.2f} 秒)")
        print(f"変更を '{args.input_file}' に保存中...")
        if workbook is None: raise RuntimeError("ワークブックオブジェクトが無効")
        try: workbook.save(args.input_file); print("保存完了")
        except PermissionError: print(f"\nエラー: '{args.input_file}' 書込権限なし (ファイルが開かれていませんか？)")
        except Exception as save_e: print(f"\nエラー: ファイル保存中に問題発生: {save_e}"); logging.error(f"ファイル保存エラー", exc_info=True)

    # 例外処理
    except FileNotFoundError: print(f"エラー: ファイル '{args.input_file}' 未検出")
    except ValueError as ve: print(f"設定/ファイルエラー: {ve}")
    except ImportError as ie:
        if 'selenium' in str(ie).lower(): print("エラー: Seleniumライブラリ未インストール (pip install selenium)")
        else: print(f"エラー: 必要なライブラリ未インストール: {ie}")
    except RuntimeError as rte: print(f"\n内部エラー: {rte}")
    except KeyboardInterrupt: print("\n処理が中断されました。")
    except Exception as e:
        print(f"\n予期せぬエラー発生: {e}")
        logging.exception("予期せぬエラー")
    finally:
        if workbook:
            try: workbook.close(); logging.info("ワークブックを閉じました。")
            except Exception as close_e: logging.error(f"ワークブッククローズ中のエラー: {close_e}")
        if 'overall_start_time' in locals():
            overall_final_elapsed_time = time.time() - overall_start_time
            print(f"\nスクリプト実行時間: {overall_final_elapsed_time:.2f} 秒")
