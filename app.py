from pathlib import Path
import re
import unicodedata

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Explorer One v0.1", page_icon="🚀", layout="wide")
st.title("Explorer One v0.1 🚀")
st.caption("Japan Defense Contract Explorer｜ITC 防衛省契約情報データベース（プロトタイプ版）")

DATA_DIR = Path("data")

COMPANY_ALIASES = {
    "NEC": ["NEC", "日本電気", "日本電気株式会社"],
    "MHI": ["MHI", "三菱重工", "三菱重工業", "三菱重工業株式会社"],
    "MELCO": ["MELCO", "三菱電機", "三菱電機株式会社"],
    "KHI": ["KHI", "川重", "川崎重工", "川崎重工業", "川崎重工業株式会社"],
    "IHI": ["IHI", "ＩＨＩ", "株式会社IHI", "株式会社ＩＨＩ"],
}


def normalize_search_text(value):
    if pd.isna(value):
        return ""
    return unicodedata.normalize("NFKC", str(value)).casefold()


def normalize_column_name(column_name):
    return (
        str(column_name)
        .replace("\n", "")
        .replace("\r", "")
        .replace(" ", "")
        .replace("　", "")
    )


def find_column(columns, keyword_patterns):
    for pattern in keyword_patterns:
        for column in columns:
            normalized = normalize_column_name(column)
            if all(keyword in normalized for keyword in pattern):
                return column
    return None


def clean_company_name(value):
    if pd.isna(value):
        return ""

    text = str(value).strip()
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    lines = [line.strip() for line in text.split("\n") if line.strip()]
    if lines:
        text = lines[0]

    text = re.split(r"〒?\s*\d{3}[-－ー]\d{4}", text, maxsplit=1)[0]

    prefectures = (
        "北海道|東京都|京都府|大阪府|"
        "青森県|岩手県|宮城県|秋田県|山形県|福島県|"
        "茨城県|栃木県|群馬県|埼玉県|千葉県|神奈川県|"
        "新潟県|富山県|石川県|福井県|山梨県|長野県|"
        "岐阜県|静岡県|愛知県|三重県|滋賀県|兵庫県|"
        "奈良県|和歌山県|鳥取県|島根県|岡山県|広島県|"
        "山口県|徳島県|香川県|愛媛県|高知県|福岡県|"
        "佐賀県|長崎県|熊本県|大分県|宮崎県|鹿児島県|沖縄県"
    )

    text = re.split(rf"\s*(?:{prefectures})", text, maxsplit=1)[0]
    return text.strip(" 　,、/／")


def expand_search_terms(keyword):
    normalized_keyword = normalize_search_text(keyword).strip()
    if not normalized_keyword:
        return []

    terms = [normalized_keyword]
    for company_names in COMPANY_ALIASES.values():
        normalized_names = [normalize_search_text(name) for name in company_names]
        if normalized_keyword in normalized_names:
            terms.extend(normalized_names)

    return list(dict.fromkeys(terms))


@st.cache_data
def load_excels():
    files = sorted(DATA_DIR.rglob("*.xlsx"))
    dataframes = []
    failed_files = []
    skipped_files = []

    for file in files:
        try:
            df = pd.read_excel(file, header=1, engine="openpyxl")

            date_column = find_column(df.columns, [["契約締結日"], ["契約年月日"], ["契約日"]])
            item_column = find_column(df.columns, [["物品役務等", "名称"], ["物品", "名称"], ["件名"]])
            company_column = find_column(df.columns, [["契約相手方", "商号"], ["契約相手方", "名称"], ["契約相手方"]])
            amount_column = find_column(df.columns, [["契約金額"], ["落札価格"], ["金額"]])

            required_columns = [date_column, item_column, company_column, amount_column]
            if any(column is None for column in required_columns):
                skipped_files.append(file.name)
                continue

            selected = df[[date_column, item_column, company_column, amount_column]].copy()
            selected.columns = ["契約日", "物品・役務名", "会社名", "契約金額（円）"]
            selected["会社名（原文）"] = selected["会社名"]
            selected["会社名"] = selected["会社名"].apply(clean_company_name)
            selected["元ファイル"] = file.name
            dataframes.append(selected)

        except Exception as error:
            failed_files.append(f"{file.name}: {error}")

    if not dataframes:
        return pd.DataFrame(), len(files), skipped_files, failed_files

    result = pd.concat(dataframes, ignore_index=True)
    result = result.dropna(how="all", subset=["契約日", "物品・役務名", "会社名（原文）", "契約金額（円）"])
    result = result[result["物品・役務名"].notna() | result["会社名（原文）"].notna()].copy()
    result["契約日"] = pd.to_datetime(result["契約日"], errors="coerce")
    result["契約金額（円）"] = pd.to_numeric(
        result["契約金額（円）"].astype(str).str.replace(",", "", regex=False).str.replace("円", "", regex=False).str.strip(),
        errors="coerce",
    )
    result = result.sort_values(by="契約日", ascending=False, na_position="last").reset_index(drop=True)

    return result, len(files), skipped_files, failed_files


df, file_count, skipped_files, failed_files = load_excels()

st.sidebar.header("検索・絞り込み")
keyword = st.sidebar.text_input("キーワード検索", placeholder="物品名・会社名を入力")
st.sidebar.caption("物品・役務名と会社名を横断して検索します。")
st.sidebar.caption("NEC、MHI、MELCO、KHI（川重）、IHIは略称にも対応しています。その他は正式社名またはその一部で検索してください。")

filtered_df = df.copy()

if keyword and not filtered_df.empty:
    searchable_text = (
        filtered_df["物品・役務名"].fillna("").astype(str)
        + " "
        + filtered_df["会社名"].fillna("").astype(str)
    ).apply(normalize_search_text)

    search_terms = expand_search_terms(keyword)
    mask = pd.Series(False, index=filtered_df.index)

    for term in search_terms:
        mask = mask | searchable_text.str.contains(term, na=False, regex=False)

    filtered_df = filtered_df[mask]

if not df.empty:
    valid_dates = df["契約日"].dropna()
else:
    valid_dates = pd.Series(dtype="datetime64[ns]")

if not valid_dates.empty:
    oldest_date = valid_dates.min().strftime("%Y-%m-%d")
    newest_date = valid_dates.max().strftime("%Y-%m-%d")
    search_period = f"{oldest_date} ～ {newest_date}"
else:
    search_period = "契約日データなし"

col1, col2, col3 = st.columns(3)
with col1:
    st.metric("読み込みファイル数", f"{file_count:,}")
with col2:
    st.metric("全契約件数", f"{len(df):,}")
with col3:
    st.metric("表示件数", f"{len(filtered_df):,}")

st.info(f"検索対象期間：{search_period}")

if skipped_files:
    with st.expander(f"列形式が異なるため読み飛ばしたファイル：{len(skipped_files)}件"):
        for file_name in skipped_files:
            st.write(file_name)

if failed_files:
    with st.expander(f"読み込みエラー：{len(failed_files)}件"):
        for error_message in failed_files:
            st.write(error_message)

if df.empty:
    st.error("データを読み込めませんでした。dataフォルダにExcelファイルが入っているか確認してください。")
elif filtered_df.empty:
    st.warning("検索条件に該当するデータがありません。")
else:
    display_df = filtered_df[["契約日", "物品・役務名", "会社名", "契約金額（円）"]].copy()

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        height=620,
        column_config={
            "契約日": st.column_config.DateColumn("契約日", format="YYYY-MM-DD", width="small"),
            "物品・役務名": st.column_config.TextColumn("物品・役務名", width="large"),
            "会社名": st.column_config.TextColumn("会社名", width="large"),
            "契約金額（円）": st.column_config.NumberColumn("契約金額（円）", format="localized", width="medium"),
        },
    )
