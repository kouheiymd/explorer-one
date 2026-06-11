from pathlib import Path
import re
import unicodedata

import pandas as pd
import streamlit as st


# =========================================================
# ページ設定
# =========================================================
st.set_page_config(
    page_title="Explorer One - Version 0.1",
    page_icon="🚀",
    layout="wide",
)

st.title("Explorer One - Version 0.1 🚀")

st.caption(
    "Japan Defense Contract Explorer｜"
    "防衛省契約情報データベース（プロトタイプ版）"
)

st.info(
    "※ v0.1では防衛装備庁が公表する契約情報のみを収録しています。"
    "今後、他の契約機関の情報にも順次対応予定です。"
)


# =========================================================
# 定数
# =========================================================
DATA_DIR = Path("data")

COMPANY_ALIASES = {
    "NEC": [
        "NEC",
        "日本電気",
        "日本電気株式会社",
    ],
    "MHI": [
        "MHI",
        "三菱重工",
        "三菱重工業",
        "三菱重工業株式会社",
    ],
    "MELCO": [
        "MELCO",
        "三菱電機",
        "三菱電機株式会社",
    ],
    "KHI": [
        "KHI",
        "川重",
        "川崎重工",
        "川崎重工業",
        "川崎重工業株式会社",
    ],
    "IHI": [
        "IHI",
        "ＩＨＩ",
        "株式会社IHI",
        "株式会社ＩＨＩ",
    ],
}


# =========================================================
# 文字列処理
# =========================================================
def normalize_text(value):
    if pd.isna(value):
        return ""

    text = unicodedata.normalize("NFKC", str(value))
    text = text.casefold()
    text = re.sub(r"\s+", "", text)
    return text


def normalize_column_name(value):
    text = unicodedata.normalize("NFKC", str(value))
    text = text.replace("\n", "")
    text = text.replace("\r", "")
    text = text.replace(" ", "")
    text = text.replace("　", "")
    return text


def clean_company_name(value):
    """
    契約相手方欄の1行目だけを会社名として使用する。
    住所部分は検索・表示から除外する。
    """
    if pd.isna(value):
        return ""

    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    lines = [
        line.strip()
        for line in text.split("\n")
        if line.strip()
    ]

    if not lines:
        return ""

    return lines[0]


def expand_search_terms(keyword):
    normalized_keyword = normalize_text(keyword)

    if not normalized_keyword:
        return []

    terms = {normalized_keyword}

    for aliases in COMPANY_ALIASES.values():
        normalized_aliases = {
            normalize_text(alias)
            for alias in aliases
        }

        if normalized_keyword in normalized_aliases:
            terms.update(normalized_aliases)

    return list(terms)


# =========================================================
# 防衛装備庁ファイル専用読み込み
# =========================================================
@st.cache_data(show_spinner=False)
def load_data():
    """
    防衛装備庁の公表Excelを読み込む。

    前提：
    ・見出しはExcelの2行目
    ・契約件名は「物品役務等の名称」
    ・契約日は「契約締結日」
    ・契約相手方は「契約相手方の商号又は名称及び住所」
    ・契約金額は「契約金額（円）」
    """
    if not DATA_DIR.exists():
        return pd.DataFrame(), []

    excel_files = sorted(DATA_DIR.rglob("*.xlsx"))
    frames = []
    skipped_files = []

    for file_path in excel_files:
        if file_path.name.startswith("~$"):
            continue

        try:
            source = pd.read_excel(
                file_path,
                header=1,
                engine="openpyxl",
            )

            column_map = {
                normalize_column_name(column): column
                for column in source.columns
            }

            item_column = column_map.get(
                normalize_column_name("物品役務等の名称")
            )
            date_column = column_map.get(
                normalize_column_name("契約締結日")
            )
            company_column = column_map.get(
                normalize_column_name(
                    "契約相手方の商号又は名称及び住所"
                )
            )
            amount_column = column_map.get(
                normalize_column_name("契約金額（円）")
            )

            required_columns = [
                item_column,
                date_column,
                company_column,
                amount_column,
            ]

            if any(column is None for column in required_columns):
                skipped_files.append(file_path.name)
                continue

            frame = source[
                [
                    date_column,
                    item_column,
                    company_column,
                    amount_column,
                ]
            ].copy()

            frame.columns = [
                "契約日",
                "契約件名",
                "契約相手方",
                "契約金額（円）",
            ]

            frame["契約相手方"] = frame["契約相手方"].apply(
                clean_company_name
            )

            frame = frame.dropna(
                how="all",
                subset=[
                    "契約日",
                    "契約件名",
                    "契約相手方",
                    "契約金額（円）",
                ],
            )

            # 注記行や空行を除外
            frame = frame[
                frame["契約件名"].notna()
                | frame["契約相手方"].astype(str).str.strip().ne("")
            ].copy()

            frame["契約日"] = pd.to_datetime(
                frame["契約日"],
                errors="coerce",
            )

            amount_text = (
                frame["契約金額（円）"]
                .astype(str)
                .str.replace(",", "", regex=False)
                .str.replace("円", "", regex=False)
                .str.replace("¥", "", regex=False)
                .str.replace("￥", "", regex=False)
                .str.strip()
            )

            frame["契約金額（円）"] = pd.to_numeric(
                amount_text,
                errors="coerce",
            )

            frames.append(frame)

        except Exception:
            skipped_files.append(file_path.name)

    if not frames:
        return pd.DataFrame(), skipped_files

    data = pd.concat(
        frames,
        ignore_index=True,
    )

    data = data[
        data["契約件名"].notna()
        & data["契約相手方"].astype(str).str.strip().ne("")
    ].copy()

    data["契約件名"] = (
        data["契約件名"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    data["契約相手方"] = (
        data["契約相手方"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    data["_検索用契約件名"] = data["契約件名"].apply(
        normalize_text
    )

    data["_検索用契約相手方"] = data["契約相手方"].apply(
        normalize_text
    )

    data = data.sort_values(
        by="契約日",
        ascending=False,
        na_position="last",
    ).reset_index(drop=True)

    return data, skipped_files


# =========================================================
# データ取得
# =========================================================
df, skipped_files = load_data()

if df.empty:
    st.error(
        "防衛装備庁の契約情報を読み込めませんでした。"
        "dataフォルダに対象のExcelファイルが入っているか確認してください。"
    )

    if skipped_files:
        with st.expander("読み込めなかったファイル"):
            for file_name in skipped_files:
                st.write(file_name)

    st.stop()


# =========================================================
# 検索条件
# =========================================================
with st.sidebar:
    st.header("検索条件")

    search_target = st.radio(
        "検索対象",
        [
            "すべて",
            "契約件名",
            "契約相手方",
        ],
    )

    selected_companies = st.multiselect(
        "主要企業",
        options=list(COMPANY_ALIASES.keys()),
        help="略称と正式名称をまとめて検索します。",
    )

    st.divider()

    st.caption(
        "契約相手方は会社名のみを検索し、住所部分は除外しています。"
    )


keyword = st.text_input(
    "キーワード検索",
    placeholder="例：衛星、レーダー、川重、IHI、三菱重工",
)


# =========================================================
# 検索処理
# =========================================================
filtered_df = df.copy()

if selected_companies:
    company_terms = set()

    for company in selected_companies:
        for alias in COMPANY_ALIASES[company]:
            company_terms.add(normalize_text(alias))

    company_mask = filtered_df["_検索用契約相手方"].apply(
        lambda value: any(
            term in value
            for term in company_terms
        )
    )

    filtered_df = filtered_df[company_mask]


if keyword.strip():
    search_terms = expand_search_terms(keyword.strip())

    if search_target == "契約件名":
        mask = filtered_df["_検索用契約件名"].apply(
            lambda value: any(
                term in value
                for term in search_terms
            )
        )

    elif search_target == "契約相手方":
        mask = filtered_df["_検索用契約相手方"].apply(
            lambda value: any(
                term in value
                for term in search_terms
            )
        )

    else:
        mask = filtered_df.apply(
            lambda row: any(
                term in row["_検索用契約件名"]
                or term in row["_検索用契約相手方"]
                for term in search_terms
            ),
            axis=1,
        )

    filtered_df = filtered_df[mask]


# =========================================================
# 集計表示
# =========================================================
valid_dates = filtered_df["契約日"].dropna()
valid_amounts = filtered_df["契約金額（円）"].dropna()

if valid_dates.empty:
    contract_period = "日付情報なし"
else:
    contract_period = (
        f"{valid_dates.min().strftime('%Y-%m-%d')}"
        f" ～ "
        f"{valid_dates.max().strftime('%Y-%m-%d')}"
    )

total_amount = (
    valid_amounts.sum()
    if not valid_amounts.empty
    else 0
)

col1, col2, col3 = st.columns(3)

with col1:
    st.metric(
        "該当件数",
        f"{len(filtered_df):,}件",
    )

with col2:
    st.metric(
        "契約金額合計",
        f"{total_amount:,.0f}円",
    )

with col3:
    st.metric(
        "契約期間",
        contract_period,
    )


# =========================================================
# 検索結果
# =========================================================
st.subheader("検索結果")

if filtered_df.empty:
    st.warning(
        "検索条件に一致する契約情報はありません。"
    )

else:
    display_df = filtered_df[
        [
            "契約日",
            "契約件名",
            "契約相手方",
            "契約金額（円）",
        ]
    ].copy()

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        height=620,
        column_config={
            "契約日": st.column_config.DateColumn(
                "契約日",
                format="YYYY-MM-DD",
                width="small",
            ),
            "契約件名": st.column_config.TextColumn(
                "契約件名",
                width="large",
            ),
            "契約相手方": st.column_config.TextColumn(
                "契約相手方",
                width="large",
            ),
            "契約金額（円）": st.column_config.NumberColumn(
                "契約金額（円）",
                format="localized",
                width="medium",
            ),
        },
    )

    csv_data = display_df.to_csv(
        index=False,
    ).encode("utf-8-sig")

    st.download_button(
        label="検索結果をCSVでダウンロード",
        data=csv_data,
        file_name="explorer_one_search_results.csv",
        mime="text/csv",
    )


# =========================================================
# フッター
# =========================================================
st.divider()
st.caption("Explorer One - Version 0.1｜Prototype")
