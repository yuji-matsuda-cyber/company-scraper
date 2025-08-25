# app.py という名前で保存してください

import streamlit as st
import pandas as pd
from googlesearch import search
import requests
from bs4 import BeautifulSoup
import time
import random
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

# --- ▼▼▼ 基本設定 ▼▼▼ ---
USER_AGENTS = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
]
EXCLUDED_DOMAINS = ['ipros', 'hotfrog', 'baseconnect', 'musubu', 'appletech', 'kensetumap', 'ja.wikipedia.org']
EXCLUDED_URL_PATHS = ['/contact', '/inquiry', '/form', '/privacy', '/policy']

# --- ★★★ 抽出専門の関数（ハイブリッドエンジン） ★★★ ---
def extract_data_from_soup(soup):
    data = {'会社名': '', '代表者名': '', '住所': '', '資本金': '', '従業員数': ''}
    
    def get_full_text(tag):
        if not tag: return ""
        return ' '.join(tag.find_all(string=True, recursive=True)).strip()

    # エンジン1：構造化データ（テーブル、リスト）から最優先で抽出
    for label_tag in soup.find_all(['th', 'dt']):
        key_text = get_full_text(label_tag)
        value_tag = label_tag.find_next_sibling(['td', 'dd'])
        if value_tag:
            value_text = get_full_text(value_tag)
            if not data.get('会社名') and any(k in key_text for k in ['会社名', '商号']): data['会社名'] = value_text
            if not data.get('代表者名') and any(k in key_text for k in ['代表者', '代表取締役']): data['代表者名'] = value_text
            if not data.get('住所') and any(k in key_text for k in ['所在地', '本社']): data['住所'] = value_text
            if not data.get('資本金') and '資本金' in key_text: data['資本金'] = value_text
            if not data.get('従業員数') and '従業員' in key_text: data['従業員数'] = value_text

    # エンジン2：キーワード近傍検索で、未取得の項目を補完
    def get_value_by_keyword_proximity(target_soup, keywords):
        for keyword in keywords:
            found_element = target_soup.find(string=re.compile(re.escape(keyword), re.IGNORECASE))
            if not found_element: continue
            for i in range(3):
                container = found_element.find_parent() if i == 0 else container.find_parent()
                if not container: break
                container_text = ' '.join(container.get_text(strip=True).split())
                value_candidate = re.sub(re.escape(keyword), '', container_text, flags=re.IGNORECASE).strip()
                if value_candidate and 1 < len(value_candidate) < 100: return value_candidate
        return ""

    if any(not data.get(key) for key in ['会社名', '代表者名', '住所']):
        if not data['会社名']: data['会社名'] = get_value_by_keyword_proximity(soup, ['会社名', '商号', '社名'])
        if not data['代表者名']: data['代表者名'] = get_value_by_keyword_proximity(soup, ['代表取締役社長', '代表取締役', '代表者'])
        if not data['住所']: data['住所'] = get_value_by_keyword_proximity(soup, ['所在地', '本社所在地', '住所'])
        if not data['資本金']: data['資本金'] = get_value_by_keyword_proximity(soup, ['資本金'])
        if not data['従業員数']: data['従業員数'] = get_value_by_keyword_proximity(soup, ['従業員数', '従業員'])

    # 最終クリーニング
    titles_rep = ['代表取締役社長', '代表取締役', '代表社員', '代表', '社長', '：', ':']
    titles_other = ['取締役', '監査役', '執行役員']
    if data.get('代表者名'):
        for title in titles_other:
            if title in data['代表者名']: data['代表者名'] = data['代表者名'].split(title)[0]
        for title in titles_rep:
            data['代表者名'] = data['代表者名'].replace(title, '')
        data['代表者名'] = data['代表者名'].strip()
    for key, value in data.items():
        if isinstance(value, str):
            cleaned_value = re.sub(r'TEL.*|FAX.*|URL.*|E-mail.*|→.*|地図.*|ダウンロード.*|〒\d{3}-\d{4}', '', value, flags=re.IGNORECASE)
            data[key] = ' '.join(cleaned_value.split())
    return data

def validate_phone_in_html(html_content, phone_number):
    soup = BeautifulSoup(html_content, 'html.parser')
    body_tag = soup.find('body')
    if not body_tag: return False, None
    body_text = body_tag.get_text()
    translation_table = str.maketrans('０１２３４５６７８９（）－‐　', '0123456789()-- ')
    normalized_body_text = body_text.translate(translation_table)
    normalized_body_text = ''.join(filter(str.isdigit, normalized_body_text))
    if phone_number in normalized_body_text: return True, soup
    return False, None

def find_valid_url(query, phone_number, status_container):
    status_container.info(f"検索中: {query}")
    try:
        search_results = search(query, num_results=5, lang='ja', sleep_interval=3)
        for temp_url in search_results:
            if any(path in temp_url for path in EXCLUDED_URL_PATHS) or any(domain in temp_url for domain in EXCLUDED_DOMAINS):
                status_container.warning(f"除外対象のためスキップ: {temp_url}")
                continue
            status_container.info(f"URL候補発見: {temp_url}")
            headers = {'User-Agent': random.choice(USER_AGENTS)}
            response = requests.get(temp_url, headers=headers, timeout=15)
            response.raise_for_status()
            is_validated, validated_soup = validate_phone_in_html(response.text, phone_number)
            if is_validated:
                status_container.success("検証成功！このURLを採用します。")
                return temp_url, validated_soup
            else:
                status_container.warning("検証失敗：電話番号が一致しません。")
                time.sleep(random.uniform(1, 2.5))
    except Exception as e:
        if "429" in str(e): raise e
        else: st.error(f"検索中に予期せぬエラー: {e}")
    return None, None

def run_scraping_process(uploaded_file, status_container):
    df = pd.read_csv(uploaded_file, dtype=str, encoding="utf-8").fillna('')
    phone_column_name = '電話番号' if '電話番号' in df.columns else '発信先電話番号'
    if phone_column_name not in df.columns:
        st.error(f"エラー: CSVに '{phone_column_name}' 列が見つかりません。")
        return

    results = []
    driver = None
    total_rows = len(df)
    google_blocked = False
    
    try:
        for index, row in df.iterrows():
            progress = (index + 1) / total_rows
            yield progress, f"[{index + 1}/{total_rows}] 処理中: {row.get(phone_column_name, '')}", None

            phone_number = row.get(phone_column_name, '')
            if not phone_number:
                results.append({'URL': '電話番号なし', '会社名': '', '代表者名': '', '住所': '', '資本金': '', '従業員数': ''})
                continue

            target_url, final_soup = None, None
            phone_formats = []
            if len(phone_number) >= 10:
                if len(phone_number) == 11 and phone_number.startswith(('070', '080', '090')): phone_formats.append(f'"{phone_number[:3]}-{phone_number[3:7]}-{phone_number[7:]}"')
                elif len(phone_number) == 10:
                    phone_formats.append(f'"{phone_number[:3]}-{phone_number[3:6]}-{phone_number[6:]}"')
                    phone_formats.append(f'"{phone_number[:4]}-{phone_number[4:6]}-{phone_number[6:]}"')
            phone_formats.append(f'"{phone_number}"')
            phone_search_group = f"({' OR '.join(phone_formats)})"
            
            query1_intitle = f'{phone_search_group} (intitle:"会社概要" OR intitle:"会社案内" OR intitle:"企業情報" OR intitle:"会社情報")'
            query2_inurl = f'{phone_search_group} (inurl:company OR inurl:profile OR inurl:about OR inurl:corporate)'
            query3_broad = phone_search_group
            
            try:
                status_container.info("[第1段階] タイトル検索を実行...")
                target_url, final_soup = find_valid_url(query1_intitle, phone_number, status_container)
                
                if not target_url:
                    wait_time = random.uniform(2, 4)
                    status_container.warning(f"第1段階で見つからず、{wait_time:.2f}秒待機してURL検索を実行します。")
                    time.sleep(wait_time)
                    target_url, final_soup = find_valid_url(query2_inurl, phone_number, status_container)

                if not target_url:
                    wait_time = random.uniform(2, 4)
                    status_container.warning(f"第2段階で見つからず、{wait_time:.2f}秒待機して広域検索を実行します。")
                    time.sleep(wait_time)
                    target_url, final_soup = find_valid_url(query3_broad, phone_number, status_container)
            except Exception as e:
                if "429" in str(e):
                    st.error("Googleからブロックされました。処理を中断します。")
                    google_blocked = True
                else: st.error(f"予期せぬエラー: {e}")
            
            if google_blocked:
                results.append({'URL': 'ブロックにより中断', '会社名': '', '代表者名': '', '住所': '', '資本金': '', '従業員数': ''})
                break

            extracted_info = {}
            if target_url and final_soup:
                extracted_info = extract_data_from_soup(final_soup)
                
                if not all(extracted_info.get(k) for k in ['会社名', '代表者名', '住所']):
                    status_container.warning("抽出不十分。JavaScript対応のためSeleniumで再試行します。")
                    if driver is None:
                        try:
                            options = Options(); options.add_argument('--headless'); options.add_argument('--disable-gpu')
                            options.add_argument("user-agent=" + random.choice(USER_AGENTS))
                            service = Service(ChromeDriverManager().install())
                            driver = webdriver.Chrome(service=service, options=options)
                        except Exception as e: st.error(f"WebDriverの起動に失敗: {e}")
                    if driver:
                        try:
                            driver.get(target_url); time.sleep(5)
                            selenium_soup = BeautifulSoup(driver.page_source, 'html.parser')
                            extracted_info = extract_data_from_soup(selenium_soup)
                        except Exception as e: st.error(f"Selenium処理中にエラー: {e}")
            
            result_data = {'URL': target_url or '見つかりません', **extracted_info}
            results.append(result_data)
            time.sleep(random.uniform(5, 10))
    finally:
        if driver: driver.quit()

    result_df = pd.DataFrame(results)
    df_processed = df.head(len(results))
    df_original = df_processed.drop(columns=[col for col in ['URL', '会社名', '代表者名', '住所', '資本金', '従業員数'] if col in df.columns])
    output_df = pd.concat([df_original.reset_index(drop=True), result_df.reset_index(drop=True)], axis=1)
    yield 1.0, "完了！", output_df

# --- ▼▼▼ Streamlit UI部分 ▼▼▼ ---
st.set_page_config(page_title="企業情報スクレイピングアプリ", layout="wide")
st.title('🤖 企業情報 自動取得アプリ')
st.markdown("CSVファイルに含まれる電話番号を元に、企業の公式ウェブサイトを検索し、会社情報を自動で取得します。")

results_placeholder = st.empty()
download_placeholder = st.empty()

uploaded_file = st.file_uploader("電話番号を含むCSVファイルをアップロードしてください", type="csv")

if uploaded_file is not None:
    if st.button('処理開始'):
        results_placeholder.empty()
        download_placeholder.empty()
        
        progress_bar = st.progress(0)
        status_container = st.expander("詳細ログ", expanded=True)
        status_container.info("処理を開始します。完了までお待ちください...")
        final_df = None

        for progress, message, df_result in run_scraping_process(uploaded_file, status_container):
            progress_bar.progress(progress)
            status_container.info(message)
            if df_result is not None:
                final_df = df_result

        st.success("🎉 全ての処理が完了しました！")
        
        results_placeholder.dataframe(final_df)
        csv_data = final_df.to_csv(index=False, encoding='utf_8_sig').encode('utf_8_sig')
        download_placeholder.download_button(
            label="結果をCSVとしてダウンロード",
            data=csv_data,
            file_name='最終結果_リスト.csv',
            mime='text/csv',
        )