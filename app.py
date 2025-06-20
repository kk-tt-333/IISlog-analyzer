# app.py
# StreamlitでIISログZIPを解析し、指定されたAccountのログをCSVまたはExcelで出力するWebアプリ

import streamlit as st
import pandas as pd
import zipfile
import io
import re
import xlsxwriter
import numpy as np

# ----------------------------
# ページ設定
# ----------------------------
st.set_page_config(page_title="IISログ解析ツール", layout="wide")
st.title("📊 IISログ解析ツール")

st.markdown("""
#### ZIP形式のIISログをアップロードし、指定したAccountのアクセスログを抽出してファイル出力します。
""")

# ----------------------------
# ユーザー入力
# ----------------------------
uploaded_file = st.file_uploader("ZIPログファイルをアップロード", type="zip")
if uploaded_file:
    st.markdown(
        f"<span style='color:green; font-size:16px; font-weight:bold;'>✔️ ファイル '{uploaded_file.name}' がアップロードされました</span>",
        unsafe_allow_html=True
    )
target_input = st.text_input("対象のAccountをカンマ区切りで入力（空欄で全件）", placeholder="例: 1234567, 1092722")
export_name = st.text_input("出力ファイル名（拡張子不要）", value="parsed_log")
file_type = st.radio("出力形式を選択", ["Excel (.xlsx)", "CSV (.csv)"], captions=["Account指定している場合こちら", "全件出力の場合はこちら"])

# ----------------------------
# ログ解析関数
# ----------------------------
def parse_iis_log(log_text, source_file):
    lines = log_text.strip().split('\n')
    field_lines = [line for line in lines if line.startswith("#Fields:")]
    if not field_lines:
        return pd.DataFrame()

    field_line = field_lines[0]
    fields = field_line.replace("#Fields: ", "").split()
    data_lines = [line for line in lines if not line.startswith("#")]
    if not data_lines:
        return pd.DataFrame()

    try:
        df = pd.read_csv(io.StringIO("\n".join(data_lines)), sep=' ', header=None, names=fields, engine='python')
    except:
        return pd.DataFrame()

    required_columns = ["date", "time", "s-computername", "cs-method", "cs-uri-stem", "cs(User-Agent)", "cs(Referer)",
                        "cs-host", "sc-status", "time-taken", "_RequestID", "True-Client-IP", "_X-SessionID"]
    if not all(col in df.columns for col in required_columns):
        return pd.DataFrame()

    try:
        df_result = pd.DataFrame({
            "datetime": df["date"] + " " + df["time"],
            "Account": df["_RequestID"].str.extract(r"@(.+)$")[0],
            "cs-method": df["cs-method"],
            "cs-uri-stem": df["cs-uri-stem"],
            "time-taken": df["time-taken"],
            "cs(User-Agent)": df["cs(User-Agent)"],
            "cs(Referer)": df["cs(Referer)"],
            "cs-host": df["cs-host"],
            "sc-status": df["sc-status"],
            "_RequestID": df["_RequestID"],
            "True-Client-IP": df["True-Client-IP"],
            "_X-SessionID": df["_X-SessionID"],
            "s-computername": df["s-computername"],
            "logfile": source_file
        })
    except Exception as e:
        st.error(f"データ整形中にエラーが発生しました: {e}")
        return pd.DataFrame()

    return df_result

# ----------------------------
# 実行処理
# ----------------------------
if 'df_all' not in st.session_state:
    st.session_state['df_all'] = None

if 'is_processing' not in st.session_state:
    st.session_state['is_processing'] = False

parse_trigger = st.button("▶ 解析実行", type="primary", disabled=st.session_state['is_processing'])

if uploaded_file and parse_trigger:
    st.session_state['is_processing'] = True
    with st.spinner("解析中..."):
        all_dfs = []
        with zipfile.ZipFile(uploaded_file, "r") as zip_ref:
            for file_name in zip_ref.namelist():
                if file_name.endswith('.log') or file_name.endswith('.txt'):
                    with zip_ref.open(file_name) as f:
                        text = f.read().decode("utf-8", errors="ignore")
                        df = parse_iis_log(text, file_name)
                        if not df.empty:
                            all_dfs.append(df)

        df_all = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()

        if not df_all.empty:
            accounts = [a.strip() for a in target_input.split(',') if a.strip()] if target_input else []
            if accounts:
                df_all = df_all[df_all["Account"].isin(accounts)]

            st.session_state['df_all'] = df_all
            st.success(f"{len(df_all)} 件のレコードが見つかりました。")
        else:
            st.session_state['df_all'] = None
            st.warning("解析結果が空です。ログ構造または対象Accountをご確認ください。")

if st.session_state['df_all'] is not None:
    st.dataframe(st.session_state['df_all'].head(5), use_container_width=True)
    df_all = st.session_state['df_all']

    if file_type == "CSV (.csv)":
        csv_output = io.BytesIO()
        df_all.to_csv(csv_output, index=False, encoding="utf-8-sig")
        st.download_button("⬇ CSVファイルをダウンロード", data=csv_output.getvalue(), file_name=f"{export_name}.csv")
        st.session_state['is_processing'] = False

    else:
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'constant_memory': True, 'nan_inf_to_errors': True})
        worksheet = workbook.add_worksheet("IISログ解析結果")

        for col_num, value in enumerate(df_all.columns):
            worksheet.write(0, col_num, str(value))

        for row_num, row in enumerate(df_all.itertuples(index=False), start=1):
            for col_num, cell in enumerate(row):
                val = "" if pd.isnull(cell) or isinstance(cell, float) and (np.isnan(cell) or np.isinf(cell)) else str(cell)
                worksheet.write(row_num, col_num, val)

        worksheet.autofilter(0, 0, len(df_all), len(df_all.columns) - 1)
        time_taken_col = df_all.columns.get_loc("time-taken")
        cell_format = workbook.add_format({"bold": True, "border": 2})
        worksheet.set_column(time_taken_col, time_taken_col, None, cell_format)
        workbook.close()

        st.download_button("⬇ Excelファイルをダウンロード", data=output.getvalue(), file_name=f"{export_name}.xlsx")
        st.session_state['is_processing'] = False

