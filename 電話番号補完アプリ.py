# app.py (ファイル名: 電話番号補完アプリ_v23_modified.py)
import streamlit as st
import pandas as pd
import time
import random
import re
import io
import zipfile
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from urllib.parse import urljoin, quote_plus
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, WebDriverException, NoSuchElementException, InvalidSessionIdException
import sys

# --- ▼▼▼ 基本設定 ▼▼▼ ---
USER_AGENTS = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36'
]
DECOY_URLS = ['https://www.yahoo.co.jp/', 'https://www.wikipedia.org/', 'https://www.nikkei.com/']

COMPANY_LINK_XPATH = " | ".join([
    "//a[contains(., '会社概要')]",
    "//a[contains(., '企業情報')]",
    "//a[contains(., '会社案内')]",
    "//a[contains(., '私たちについて')]",
    "//a[contains(@href, 'company')]",
    "//a[contains(@href, 'about')]",
    "//a[contains(@href, 'corporate')]",
    "//a[contains(@href, 'profile')]",
    "//a[contains(@href, 'gaiyou')]",
])

SUB_COMPANY_LINK_XPATH = " | ".join([
    "//a[contains(., '概要')]",
    "//a[contains(., '沿革')]",
    "//a[contains(., '拠点')]",
    "//a[contains(., '事業所')]",
    "//a[contains(., 'アクセス')]",
    "//a[contains(@href, 'outline')]",
    "//a[contains(@href, 'access')]",
    "//a[contains(@href, 'location')]",
    "//a[contains(@href, 'base')]",
])

# --- プロキシ設定用関数 ---
def create_proxy_extension(proxy_host, proxy_port, proxy_user, proxy_pass):
    manifest_json = """{"version": "1.0.0","manifest_version": 2,"name": "Chrome Proxy","permissions": ["proxy","tabs","unlimitedStorage","storage","<all_urls>","webRequest","webRequestBlocking"],"background": {"scripts": ["background.js"]}}"""
    background_js = f"""var config = {{mode: "fixed_servers",rules: {{singleProxy: {{scheme: "http",host: "{proxy_host}",port: parseInt({proxy_port})}},bypassList: ["localhost"]}}}};chrome.proxy.settings.set({{value: config, scope: "regular"}}, function() {{}});function callbackFn(details) {{return {{authCredentials: {{username: "{proxy_user}",password: "{proxy_pass}"}}}};}}chrome.webRequest.onAuthRequired.addListener(callbackFn,{{urls: ["<all_urls>"]}},['blocking']);"""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        zf.writestr("manifest.json", manifest_json)
        zf.writestr("background.js", background_js)
    return zip_buffer.getvalue()

# --- ★★★ 電話番号抽出関連関数 ★★★ ---
def extract_phone_number(soup, area_codes_set, sorted_area_codes):
    """HTML(soup)から電話番号を抽出"""
    try:
        def validate_and_add_internal(phone_digits, phones_list):
            if not phone_digits or (phone_digits in phones_list): return False
            if phone_digits.startswith(('0120', '0800')): return False
            if phone_digits.startswith(('050', '070', '080', '090')):
                phones_list.append(phone_digits); return True
            is_valid_area_code = False
            for code in sorted_area_codes:
                if phone_digits.startswith(code): is_valid_area_code = True; break
            if is_valid_area_code: phones_list.append(phone_digits); return True
            return False

        for tag in soup(['script', 'style', 'header', 'nav', 'aside']):
            tag.decompose()
        full_text = soup.get_text()
        translation_table = str.maketrans('０１２３４５６７８９（）－‐　', '0123456789()-- ')
        normalized_text = full_text.translate(translation_table)
        found_phones = []

        pattern1 = r'(?:TEL|電話番号|電話)\s*[.:：]?\s*(0\d{1,4}[-()（）\s]{1,3}\d{1,4}[-()（）\s]{1,3}\d{3,4})'
        matches1 = re.finditer(pattern1, normalized_text, re.IGNORECASE)
        for m in matches1:
            phone_digits = re.sub(r'[^0-9]', '', m.group(1))
            if len(phone_digits) == 10 or len(phone_digits) == 11:
                validate_and_add_internal(phone_digits, found_phones)

        pattern2 = r'0\d{1,4}[-()（）\s]{1,3}\d{1,4}[-()（）\s]{1,3}\d{3,4}'
        matches2 = re.findall(pattern2, normalized_text)
        for candidate in matches2:
            phone_digits = re.sub(r'[^0-9]', '', candidate)
            if len(phone_digits) == 10 or len(phone_digits) == 11:
                validate_and_add_internal(phone_digits, found_phones)

        pattern3 = r'(?<!\d)(0\d{9,10})(?!\d)'
        matches3 = set(re.findall(pattern3, normalized_text))
        for phone_digits in matches3:
            validate_and_add_internal(phone_digits, found_phones)

        if found_phones:
            mobile_phones = [p for p in found_phones if p.startswith(('070', '080', '090'))]
            if mobile_phones: return mobile_phones[0]
            else: return found_phones[0]
        return None
    except Exception as e:
        print(f"電話番号抽出中にエラー: {e}"); return None

# --- ★★★ Yahoo検索(検索結果ページ)から電話番号を探す関数 ★★★ ---
def search_yahoo_search_phone(driver, facility_name, address, status_container):
    """Yahoo検索結果ページから施設名と住所で電話番号を探す"""
    phone_number = 'N/A'
    # ▼▼▼ 修正: 'nan' もチェック対象に追加 ▼▼▼
    if not facility_name or facility_name.lower() in ['n/a', 'アクセスエラー', '抽出エラー', 'nan', ''] or \
       not address or address.lower() in ['n/a', 'アクセスエラー', '抽出エラー', 'nan', '']:
        status_container.info(f" -> 屋号/住所が無効なためYahoo検索(ダイレクト)スキップ。(屋号: {facility_name}, 住所: {address})")
        return phone_number
    # ▲▲▲ ここまで修正 ▲▲▲

    clean_facility_name = re.sub(r'【.*?】|\(.*?\)|（.*?）|の.*?求人.*', '', facility_name).strip()
    if not clean_facility_name:
        clean_facility_name = facility_name

    search_query = f'"{clean_facility_name}" "{address}"'
    status_container.info(f" -> Yahoo検索(ダイレクト)開始: '{search_query}'")

    try:
        search_url = f"https://search.yahoo.co.jp/search?p={quote_plus(search_query)}"
        status_container.info(f" -> Yahoo検索ページに移動します: {search_url}")
        driver.set_page_load_timeout(30)
        driver.get(search_url)
        time.sleep(random.uniform(1.0, 2.0))

        phone_xpath = "//span[contains(@class, 'AnswerLocalSpot__subInfoSpotDetail') and text()='電話：']/following-sibling::span[1]"

        try:
            status_container.info(f" -> 電話番号要素 ({phone_xpath}) を検索...")
            phone_element = driver.find_element(By.XPATH, phone_xpath)
            phone_text = phone_element.text.strip()

            if phone_text and re.fullmatch(r'[\d-]+', phone_text):
                phone_number = phone_text
                status_container.success(f" ----> 電話番号候補 (Yahoo検索結果): {phone_number}")
            else:
                status_container.warning(f" ----> 電話番号要素のテキストが不正または空: '{phone_text}'")

        except NoSuchElementException:
            status_container.warning(f" ----> 電話番号要素 ({phone_xpath}) が見つかりませんでした。")
        except Exception as e_phone:
            status_container.error(f" ----> 電話番号抽出(Yahoo検索結果)エラー: {e_phone}")

    except InvalidSessionIdException as e_sid:
        status_container.error(f" -> Yahoo検索中にセッション無効: {e_sid}"); raise
    except TimeoutException:
        status_container.warning(f" -> Yahoo検索ページ ({search_url}) の読み込みタイムアウト。")
    except Exception as e:
        status_container.error(f" -> Yahoo検索中に予期せぬエラー: {e}")

    return phone_number

# --- ★★★ (従来の)Yahoo検索(検索結果一覧)で電話番号を探す関数 ★★★ ---
def search_yahoo_for_phone(query, driver, area_codes_set, sorted_area_codes, status_container):
    """(従来)Yahoo検索結果一覧から電話番号を抽出する"""
    try:
        status_container.info(f"(予備) Yahoo検索(一覧)を実行: {query}")
        search_url = f"https://search.yahoo.co.jp/search?p={quote_plus(query)}"
        driver.set_page_load_timeout(30)
        driver.get(search_url)
        time.sleep(random.uniform(2.0, 3.0))

        soup = BeautifulSoup(driver.page_source, 'html.parser')
        result_blocks = soup.select('div.sw-CardBase, div.Algo, section.Algo')

        found_phones_yahoo = []

        def validate_and_add_yahoo(phone_digits, phones_list):
            if not phone_digits or (phone_digits in phones_list): return False
            if phone_digits.startswith(('0120', '0800')): return False
            if phone_digits.startswith(('050', '070', '080', '090')):
                phones_list.append(phone_digits); return True
            is_valid_area_code = False
            for code in sorted_area_codes:
                if phone_digits.startswith(code): is_valid_area_code = True; break
            if is_valid_area_code: phones_list.append(phone_digits); return True
            return False

        for block in result_blocks[:5]:
            block_text = block.get_text()
            translation_table = str.maketrans('０１２３４５６７８９（）－‐　', '0123456789()-- ')
            normalized_text = block_text.translate(translation_table)

            pattern1 = r'(?:電話番号|電話|TEL)\s*[.:：]?\s*(0\d{1,4}[-()（）\s]{1,3}\d{1,4}[-()（）\s]{1,3}\d{3,4})'
            matches1 = re.finditer(pattern1, normalized_text, re.IGNORECASE)
            for m in matches1:
                phone_digits = re.sub(r'[^0-9]', '', m.group(1))
                if len(phone_digits) == 10 or len(phone_digits) == 11:
                    validate_and_add_yahoo(phone_digits, found_phones_yahoo)

            pattern2 = r'0\d{1,4}[-()（）\s]{1,3}\d{1,4}[-()（）\s]{1,3}\d{3,4}'
            matches2 = re.findall(pattern2, normalized_text)
            for candidate in matches2:
                phone_digits = re.sub(r'[^0-9]', '', candidate)
                if len(phone_digits) == 10 or len(phone_digits) == 11:
                    validate_and_add_yahoo(phone_digits, found_phones_yahoo)

            pattern3 = r'(?<!\d)(0\d{9,10})(?!\d)'
            matches3 = set(re.findall(pattern3, normalized_text))
            for phone_digits in matches3:
                    validate_and_add_yahoo(phone_digits, found_phones_yahoo)

            if found_phones_yahoo:
                break

        if found_phones_yahoo:
            mobile_phones = [p for p in found_phones_yahoo if p.startswith(('070', '080', '090'))]
            if mobile_phones: return mobile_phones[0]
            else: return found_phones_yahoo[0]

        return None

    except (TimeoutException, WebDriverException) as e:
        status_container.warning(f"(予備) Yahoo検索(一覧)中にタイムアウトまたはエラー: {e}")
        return None
    except InvalidSessionIdException as e_sid:
        status_container.error(f" -> Yahoo検索(一覧)中にセッション無効: {e_sid}"); raise
    except Exception as e:
        status_container.error(f"(予備) Yahoo検索(一覧)中に予期せぬエラー: {e}")
        return None


# --- ★★★ メイン処理: run_scraping_process ★★★ ---
def run_scraping_process(df, status_container, proxy_settings, disable_headless, area_codes_set):

    phone_column_name = '電話番号'
    hp_column_name = 'HP'
    company_name_cols = ['屋号']
    address_cols = ['住所', '所在地']

    actual_company_col = next((col for col in company_name_cols if col in df.columns), None)
    actual_address_col = next((col for col in address_cols if col in df.columns), None)

    if phone_column_name not in df.columns:
         st.error(f"エラー: CSVに '{phone_column_name}' 列が見つかりません。")
         yield 1.0, "列名エラー(電話番号)", df
         return
    
    # ▼▼▼ 修正: 列存在の警告を移動 ▼▼▼
    # if not (actual_company_col and actual_address_col):
    #     st.warning(f"注意: CSVに会社名({', '.join(company_name_cols)})または住所({', '.join(address_cols)})列が見つからないため、Yahoo検索は実行されません。")
    # ▲▲▲ ここまで修正 ▲▲▲

    sleep_times = {"visit": (1.5, 2.5), "decoy": (1, 2), "loop": (1, 2)}
    driver = None

    target_indices = df[
        (df[phone_column_name].isnull() | (df[phone_column_name] == ''))
    ].index

    total_jobs = len(target_indices)
    if total_jobs == 0:
        st.warning(f"処理対象（'{phone_column_name}'が空の行）が0件です。")
        yield 1.0, "処理対象なし", df
        return

    sorted_area_codes = sorted(area_codes_set, key=len, reverse=True)
    df_copy = df.copy()
    processed_count = 0

    try:
        status_container.info("ブラウザを起動しています...");
        options = Options()
        options.add_argument(f'user-agent={random.choice(USER_AGENTS)}')
        options.add_argument('--blink-settings=imagesEnabled=false')
        if not disable_headless: options.add_argument('--headless=new')
        options.add_argument(f'--window-size=1920,1980')
        options.add_argument('--disable-gpu'); options.add_argument('--lang=ja-JP,ja;q=0.9')
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"]); options.add_experimental_option('useAutomationExtension', False)

        proxy_values = {k: v for k, v in proxy_settings.items() if v}
        if all(k in proxy_values for k in ['proxy_host', 'proxy_port', 'proxy_user', 'proxy_pass']):
            try:
                options.add_extension(io.BytesIO(create_proxy_extension(**proxy_values)))
                status_container.info("プロキシ設定を適用しました。")
            except Exception as e:
                st.error(f"プロキシ設定エラー: {e}")

        try: 
             service = Service(ChromeDriverManager().install());
             driver = webdriver.Chrome(service=service, options=options)
             driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {'source': 'Object.defineProperty(navigator, "webdriver", {get: () => undefined})'})
             driver.set_page_load_timeout(30)
             status_container.success("ブラウザの起動が完了しました。")
        except Exception as e_setup:
             st.error(f"WebDriverの起動に失敗しました: {e_setup}")
             yield 1.0, "ブラウザ起動エラー", df
             return


        for index in target_indices:
            processed_count += 1
            progress_rate = processed_count / total_jobs
            yield progress_rate, f"{processed_count}/{total_jobs}件目 処理中", None

            row = df_copy.loc[index]
            company_hp_url = str(row.get(hp_column_name, '')).strip()
            
            # ▼▼▼ 修正: nan 文字列が入らないようにする & bool チェック追加 ▼▼▼
            company_name_raw = row.get(actual_company_col) if actual_company_col else None
            address_raw = row.get(actual_address_col) if actual_address_col else None
            
            # pd.notna で None や np.nan をチェックし、かつ空文字でないことを確認
            company_name = str(company_name_raw).strip() if pd.notna(company_name_raw) and str(company_name_raw).strip() else ""
            address = str(address_raw).strip() if pd.notna(address_raw) and str(address_raw).strip() else ""
            
            # この行の検索にYahoo検索が可能か判定
            yahoo_search_possible_for_this_row = bool(company_name) and bool(address) 
            if not yahoo_search_possible_for_this_row:
                 status_container.info(f" -> 屋号/住所が空欄または無効なため、Yahoo検索はスキップします。 (屋号: '{company_name}', 住所: '{address}')")
            # ▲▲▲ ここまで修正 ▲▲▲

            found_phone = None
            current_search_step = ""

            try:
                # --- HP URLがある場合のみサイト訪問 ---
                if company_hp_url and company_hp_url.startswith('http'):
                    current_search_step = "HP"
                    status_container.info(f"アクセス中: {company_hp_url}")
                    try:
                        driver.set_page_load_timeout(30)
                        driver.get(company_hp_url)
                        time.sleep(3)
                        soup = BeautifulSoup(driver.page_source, 'html.parser')
                        found_phone = extract_phone_number(soup, area_codes_set, sorted_area_codes)
                        if found_phone: status_container.success(f"HPトップで番号抽出成功: {found_phone}")
                    except (TimeoutException, WebDriverException) as e:
                        status_container.warning(f"ページロードエラー({current_search_step})。下層ページ検索へ移行: {e}")
                        found_phone = None
                    except InvalidSessionIdException as e_sid:
                        status_container.error(f"HPアクセス中にセッション無効: {e_sid}"); raise

                    # --- 概要ページ1 ---
                    if not found_phone:
                        status_container.info("トップページに番号なし。概要ページを探します...")
                        overview_url_l1 = None
                        base_url = driver.current_url if driver.current_url else company_hp_url

                        try:
                            wait = WebDriverWait(driver, 7)
                            link_element = wait.until(EC.presence_of_element_located((By.XPATH, f"({COMPANY_LINK_XPATH})[1]")))
                            link_href = link_element.get_attribute('href')
                            if link_href and not link_href.startswith(('javascript:', 'tel:', 'mailto:')) and '#' not in link_href.split('/')[-1]:
                                overview_url_l1 = urljoin(base_url, link_href)
                                base_domain_match = re.search(r"https://?([^/]+)", base_url)
                                if base_domain_match:
                                    base_domain = base_domain_match.group(1)
                                    if base_domain not in overview_url_l1: overview_url_l1 = None
                                    else: status_container.success(f"概要ページを発見！ -> {overview_url_l1}")
                                else: overview_url_l1 = None
                            else: overview_url_l1 = None
                        except Exception: pass

                        if overview_url_l1:
                            current_search_step = "概要1"
                            status_container.info(f"アクセス中: {overview_url_l1}")
                            try:
                                driver.set_page_load_timeout(30)
                                driver.get(overview_url_l1)
                                time.sleep(3)
                                soup_l1 = BeautifulSoup(driver.page_source, 'html.parser')
                                found_phone = extract_phone_number(soup_l1, area_codes_set, sorted_area_codes)
                                if found_phone: status_container.success(f"概要1で番号抽出成功: {found_phone}")
                            except (TimeoutException, WebDriverException) as e:
                                status_container.warning(f"ページロードエラー({current_search_step})。下層ページ検索へ移行: {e}")
                                found_phone = None
                            except InvalidSessionIdException as e_sid:
                                status_container.error(f"概要1アクセス中にセッション無効: {e_sid}"); raise

                            # --- 概要ページ2 ---
                            if not found_phone:
                                status_container.info("概要1に番号なし。さらに詳細ページを探します...")
                                overview_url_l2 = None
                                base_url_l1 = driver.current_url if driver.current_url else overview_url_l1
                                current_url_no_hash = overview_url_l1.split('#')[0] if overview_url_l1 else ""

                                try:
                                    wait = WebDriverWait(driver, 3)
                                    link_element = wait.until(EC.presence_of_element_located((By.XPATH, f"({SUB_COMPANY_LINK_XPATH})[1]")))
                                    link_href = link_element.get_attribute('href')
                                    if link_href and not link_href.startswith(('javascript:', 'tel:', 'mailto:')) and '#' not in link_href.split('/')[-1]:
                                        overview_url_l2_candidate = urljoin(base_url_l1, link_href)
                                        if overview_url_l2_candidate.split('#')[0] != current_url_no_hash:
                                            overview_url_l2 = overview_url_l2_candidate
                                            base_domain_match_l1 = re.search(r"https://?([^/]+)", base_url)
                                            if base_domain_match_l1:
                                                base_domain = base_domain_match_l1.group(1)
                                                if base_domain not in overview_url_l2: overview_url_l2 = None
                                                else: status_container.success(f"詳細ページを発見！ -> {overview_url_l2}")
                                            else: overview_url_l2 = None
                                        else: overview_url_l2 = None
                                    else: overview_url_l2 = None
                                except Exception: pass

                                if overview_url_l2:
                                    current_search_step = "概要2"
                                    status_container.info(f"アクセス中: {overview_url_l2}")
                                    try:
                                        driver.set_page_load_timeout(30)
                                        driver.get(overview_url_l2)
                                        time.sleep(3)
                                        soup_l2 = BeautifulSoup(driver.page_source, 'html.parser')
                                        found_phone = extract_phone_number(soup_l2, area_codes_set, sorted_area_codes)
                                        if found_phone: status_container.success(f"概要2で番号抽出成功: {found_phone}")
                                    except (TimeoutException, WebDriverException) as e:
                                        status_container.warning(f"ページロードエラー({current_search_step})。Yahoo検索へ移行: {e}")
                                        found_phone = None
                                    except InvalidSessionIdException as e_sid:
                                         status_container.error(f"概要2アクセス中にセッション無効: {e_sid}"); raise
                else:
                    status_container.info("「HP」のURLが無効または空です。Yahoo検索を試みます。")

                # --- ▼▼▼ 修正: yahoo_search_possible_for_this_row でチェック ▼▼▼ ---
                if not found_phone:
                    if yahoo_search_possible_for_this_row:
                        status_container.info("企業HPから番号が見つからなかったか「HP」がありません。Yahoo検索(ダイレクト)で補完します...")
                        found_phone_direct = search_yahoo_search_phone(driver, company_name, address, status_container)
                        if found_phone_direct and found_phone_direct != 'N/A':
                            found_phone = found_phone_direct
                            status_container.success(f"Yahoo検索(ダイレクト)で番号抽出成功: {found_phone}")
                        else:
                            found_phone = None
                    else:
                         status_container.warning("会社名(屋号)/住所が無効なため、Yahoo検索(ダイレクト)はスキップします。")

                if not found_phone:
                    if yahoo_search_possible_for_this_row:
                        status_container.info("Yahoo検索(ダイレクト)でも見つかりません。(予備)Yahoo検索(一覧)で補完します...")
                        search_company_name = re.sub(r'[（\(][株有合][）\)]', '', company_name).strip()
                        address_match = re.match(r'(東京都|北海道|(?:京都|大阪)府|.{2,3}県)([^市]+市|[^区]+区|[^郡]+郡[^町]+町|[^郡]+郡[^村]+村|[^町]+町|[^村]+村)', address)
                        search_address = address_match.group(0) if address_match else address
                        query = f'"{search_company_name}" "{search_address}" 電話番号'
                        found_phone_list = search_yahoo_for_phone(query, driver, area_codes_set, sorted_area_codes, status_container)
                        if found_phone_list:
                            found_phone = found_phone_list
                            status_container.success(f"(予備)Yahoo検索(一覧)で電話番号を抽出: {found_phone}")
                        else:
                            status_container.warning("(予備)Yahoo検索(一覧)でも電話番号は見つかりませんでした。")
                    else:
                         status_container.warning("会社名(屋号)/住所が無効なため、(予備)Yahoo検索(一覧)はスキップします。")
                # --- ▲▲▲ ここまで修正 ▲▲▲ ---

                # --- 抽出結果の記録 ---
                if found_phone:
                    df_copy.loc[index, phone_column_name] = found_phone
                else:
                    df_copy.loc[index, phone_column_name] = '見つかりません'

            except InvalidSessionIdException as e_sid:
                st.error(f"処理中にセッションが無効になりました: {e_sid}")
                st.warning("処理を中断します。再度実行してください。")
                break
            except Exception as e:
                st.error(f"URL処理({current_search_step})中に予期せぬエラー ({company_hp_url}): {e}")
                df_copy.loc[index, phone_column_name] = f'エラー({current_search_step})'

            # --- (デコイ処理) ---
            if processed_count % 5 == 0 and processed_count > 0:
                try:
                    decoy_url = random.choice(DECOY_URLS)
                    status_container.info(f"パターン偽装のため、無関係なサイトにアクセスします: {decoy_url}")
                    try:
                        driver.set_page_load_timeout(15)
                        driver.get(decoy_url)
                        time.sleep(random.uniform(*sleep_times["decoy"]))
                    except (TimeoutException, WebDriverException) as e:
                        status_container.warning(f"デコイアクセスでエラー（タイムアウト等）: {e}")
                    except InvalidSessionIdException as e_sid:
                         status_container.error(f"デコイアクセス中にセッション無効: {e_sid}"); raise
                    except Exception as e_decoy:
                        status_container.warning(f"デコイアクセスで予期せぬエラー: {e_decoy}")
                except Exception as e_outer:
                     status_container.warning(f"デコイ処理全体でエラー: {e_outer}")


            time.sleep(random.uniform(*sleep_times["loop"]))

    except InvalidSessionIdException as e_sid:
       st.error(f"処理の途中でブラウザセッションが無効になりました: {e_sid}")
       st.warning("途中までの結果を出力します。")
       driver = None
    finally:
       if driver: driver.quit()

    yield 1.0, "完了！", df_copy

# --- ▼▼▼ Streamlit UI部分 ▼▼▼ ---
st.set_page_config(page_title="電話番号 補完アプリ", layout="centered")
st.title('🤖 電話番号 自動補完アプリ')
st.markdown("CSVまたはExcelの「HP」「屋号」「住所/所在地」を元に、空欄の「電話番号」列を自動で補完します。") # ラベルを変更
st.sidebar.title("⚙️ 動作設定")
disable_headless = st.sidebar.checkbox("ヘッドレスモードを無効化（デバッグ用）")
with st.sidebar.expander("プロキシ設定（上級者向け）", expanded=False):
    proxy_settings = {
        "proxy_host": st.text_input("ホスト"),
        "proxy_port": st.text_input("ポート"),
        "proxy_user": st.text_input("ユーザー名"),
        "proxy_pass": st.text_input("パスワード", type="password")
    }
progress_text, p_bar, time_info = st.empty(), st.empty(), st.empty()
results_placeholder, download_placeholder = st.empty(), st.empty()

AREA_CODE_CSV_PATH = "市外局番リスト.csv"

# ▼▼▼ 修正: type=["csv", "xlsx", "xls"] に変更 ▼▼▼
if uploaded_file := st.file_uploader("処理対象ファイル (電話番号, [HP], [屋号], [住所/所在地] 列を含む) をアップロード", type=["csv", "xlsx", "xls"]):
# ▲▲▲ ここまで修正 ▲▲▲

    if st.button('処理開始'):

        try:
            try:
                area_codes_df = pd.read_csv(AREA_CODE_CSV_PATH, dtype=str, encoding="utf-8-sig")
            except UnicodeDecodeError:
                try:
                    area_codes_df = pd.read_csv(AREA_CODE_CSV_PATH, dtype=str, encoding="cp932")
                    st.info("市外局番リストを cp932 (Shift-JIS) で読み込みました。")
                except Exception:
                    area_codes_df = pd.read_csv(AREA_CODE_CSV_PATH, dtype=str, encoding="utf-8")
                    st.info("市外局番リストを utf-8 で読み込みました。")

            area_codes_df.columns = area_codes_df.columns.str.strip()

            if '市外局番' not in area_codes_df.columns:
                st.error(f"エラー: '{AREA_CODE_CSV_PATH}' に '市外局番' という列が見つかりません。")
                st.stop()

            area_codes_df['市外局番'] = area_codes_df['市外局番'].str.strip()
            area_codes_set = set(area_codes_df['市外局番'].astype(str).str.zfill(2))
            st.info(f"✅ 市外局番リスト ({AREA_CODE_CSV_PATH}) を読み込みました。 (件数: {len(area_codes_set)})")

        except FileNotFoundError:
            st.error(f"エラー: '{AREA_CODE_CSV_PATH}' が見つかりません。スクリプトと同じ場所に配置してください。")
            st.stop()
        except Exception as e:
            st.error(f"市外局番リストの読み込み中に致命的なエラーが発生しました: {e}")
            st.stop()

        # ▼▼▼ 修正: ファイル読み込み部分を拡張子で分岐 ▼▼▼
        try:
            original_filename = uploaded_file.name
            
            if original_filename.lower().endswith('.csv'):
                st.info(f"CSVファイル ({original_filename}) を読み込んでいます...")
                try:
                    df = pd.read_csv(uploaded_file, dtype=str, encoding="utf-8-sig")
                except UnicodeDecodeError:
                    try:
                        uploaded_file.seek(0)
                        df = pd.read_csv(uploaded_file, dtype=str, encoding="utf-8")
                        st.info("CSVを utf-8 で読み込みました。")
                    except UnicodeDecodeError:
                        uploaded_file.seek(0)
                        df = pd.read_csv(uploaded_file, dtype=str, encoding="cp932")
                        st.info("CSVを cp932 (Shift-JIS) で読み込みました。")
            
            elif original_filename.lower().endswith(('.xlsx', '.xls')):
                st.info(f"Excelファイル ({original_filename}) を読み込んでいます...")
                df = pd.read_excel(uploaded_file, dtype=str)
            
            else:
                st.error("サポートされていないファイル形式です。CSV, XLSX, XLS ファイルをアップロードしてください。")
                st.stop()

            df.columns = df.columns.str.strip()
        
        except Exception as e:
            st.error(f"ファイル ({original_filename}) の読み込みに失敗しました: {e}")
            st.stop()
        # ▲▲▲ ここまで修正 ▲▲▲

        p_bar.progress(0); status_container = st.expander("詳細ログ", expanded=True)
        start_time = time.time()
        final_df = None

        phone_col = '電話番号'
        hp_col = 'HP'
        company_name_cols = ['屋号']
        address_cols = ['住所', '所在地']
        actual_company_col = next((col for col in company_name_cols if col in df.columns), None)
        actual_address_col = next((col for col in address_cols if col in df.columns), None)
        
        # ▼▼▼ 修正: 列存在の警告をここに移動 ▼▼▼
        if not (actual_company_col and actual_address_col):
            st.warning(f"注意: CSVに会社名({', '.join(company_name_cols)})または住所({', '.join(address_cols)})列が見つからないため、Yahoo検索は実行されません。")
        # ▲▲▲ ここまで修正 ▲▲▲

        total_jobs_for_eta = 0
        if phone_col in df.columns:
            total_jobs_for_eta = len(df[
                (df[phone_col].isnull() | (df[phone_col] == ''))
            ])


        processed_count_for_eta = 0
        for prog, msg, df_result in run_scraping_process(df, status_container, proxy_settings, disable_headless, area_codes_set):
            p_bar.progress(prog); progress_text.text(msg); status_container.info(msg)

            if df_result is None and total_jobs_for_eta > 0:
                processed_count_for_eta += 1
                elapsed = time.time() - start_time
                if processed_count_for_eta > 1:
                    eta_total = elapsed / (processed_count_for_eta / total_jobs_for_eta)
                    eta_finish_time = start_time + eta_total
                    time_info.info(f"予想処理時間: 約{int(eta_total//60)}分 (完了予定: {time.strftime('%H:%M頃', time.localtime(eta_finish_time))})")

            if df_result is not None:
                final_df = df_result

        if msg.startswith("完了") or msg.startswith("列名エラー") or msg.startswith("処理対象なし") or msg.startswith("ブラウザ起動エラー"):
            st.success(f"🎉 {msg}");
        else:
             st.warning(f"処理が完了前に終了しました: {msg}")

        results_placeholder.dataframe(final_df)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
             final_df.to_excel(writer, index=False, sheet_name='Sheet1')
        excel_data = output.getvalue()

        base_filename = original_filename.rsplit('.', 1)[0] if '.' in original_filename else original_filename
        download_filename = f"{base_filename}_番号抽出完了.xlsx"
        download_placeholder.download_button("結果をExcelダウンロード", excel_data, download_filename, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')