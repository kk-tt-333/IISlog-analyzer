# IISログアップロード→解析→Excel出力までを行うWebアプリ（Streamlitベース）

import streamlit as st
import pandas as pd
import re
import io
import zipfile
from datetime import datetime
from pandas import ExcelWriter


def parse_iis_log(log_text):
    lines = log_text.strip().split('\n')
    field_lines = [line for line in lines if line.startswith("#Fields:")]
    if not field_lines:
        st.error("ログファイルに '#Fields:' 行が見つかりません。IISログ形式を確認してください。")
        return pd.DataFrame()

    field_line = field_lines[0]
    fields = field_line.replace("#Fields: ", "").split()
    data_lines = [line for line in lines if not line.startswith("#")]

    if not data_lines:
        st.warning("データ行が見つかりませんでした。ログ内容をご確認ください。")
        return pd.DataFrame()

    try:
        df = pd.read_csv(io.StringIO("\n".join(data_lines)), sep=' ', header=None, names=fields, engine='python')
    except Exception as e:
        st.error(f"データフレーム作成時にエラーが発生しました: {e}")
        return pd.DataFrame()

    if df.empty:
        st.warning("読み込んだログデータが空です。")
        return pd.DataFrame()

    required_columns = ["date", "time", "s-computername", "cs-method", "cs(User-Agent)", "cs(Referer)", "cs-host", "sc-status", "time-taken", "_RequestID", "True-Client-IP", "_X-SessionID"]
    for col in required_columns:
        if col not in df.columns:
            st.error(f"必要なカラムが不足しています: {col}")
            return pd.DataFrame()

    df_result = pd.DataFrame({
        "datetime": df["date"] + " " + df["time"],
        "s-computername": df["s-computername"],
        "cs-method": df["cs-method"],
        "cs(User-Agent)": df["cs(User-Agent)"],
        "cs(Referer)": df["cs(Referer)"],
        "cs-host": df["cs-host"],
        "sc-status": df["sc-status"],
        "time-taken": df["time-taken"],
        "_RequestID": df["_RequestID"],
        "True-Client-IP": df["True-Client-IP"],
        "_X-SessionID": df["_X-SessionID"]
    })
    df_result["Account"] = df_result["_RequestID"].str.extract(r"@(.+)$")
    return df_result


st.set_page_config(page_title="IISログ解析ツール", layout="wide")
st.title("📊 IISログ解析ツール")
st.markdown("<span style='color:blue; font-size:16px; font-weight:bold;'>ZIP形式のIISログファイルをアップロードし、主要項目を抽出してExcel出力します</span>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("IISログZIPファイルをアップロード（.zipのみ）", type=["zip"])

if uploaded_file:
    st.markdown(f"<span style='color:green; font-size:18px; font-weight:bold;'>✔️ ファイル '{uploaded_file.name}' がアップロードされました</span>", unsafe_allow_html=True)

    content = ""
    with zipfile.ZipFile(uploaded_file, "r") as zip_ref:
        all_text = []
        for file_name in zip_ref.namelist():
            if file_name.endswith('.log') or file_name.endswith('.txt'):
                with zip_ref.open(file_name) as f:
                    text = f.read().decode("utf-8", errors="ignore")
                    all_text.append(text)
        content = "\n".join(all_text)

    with st.spinner("🔄 ログ解析中..."):
        df_output = parse_iis_log(content)
        st.session_state["df"] = df_output

if "df" in st.session_state:
    df = st.session_state["df"]
    if not df.empty:
        st.success(f"{len(df)} 行のログを解析しました。下記に内容を表示します。")
        st.dataframe(df, use_container_width=True)

        output = io.BytesIO()
        with ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='ParsedLog', index=False)
            worksheet = writer.sheets['ParsedLog']
            worksheet.auto_filter.ref = worksheet.dimensions
        st.download_button("⬇ Excelファイルをダウンロード", data=output.getvalue(), file_name="iis_log_output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("解析結果がありません。ログファイルの内容を確認してください。")
