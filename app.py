# app.py
# StreamlitでIISログZIPを解析し、指定されたAccountのログをExcel出力するWebアプリ

import streamlit as st
import pandas as pd
import zipfile
import io
import re

# ----------------------------
# ページ設定
# ----------------------------
st.set_page_config(page_title="IISログ解析ツール", layout="wide")
st.title("📊 IISログ解析ツール")

st.markdown("""
#### ZIP形式のIISログをアップロードし、指定したAccountのアクセスログを抽出してExcel出力します。
""")

# ----------------------------
# ユーザー入力
# ----------------------------
uploaded_file = st.file_uploader("ZIPログファイルをアップロード", type="zip")
target_input = st.text_input("対象のAccountをカンマ区切りで入力（空欄で全件）", placeholder="例: 1234567, 1092722")
export_name = st.text_input("Excel出力ファイル名（.xlsxは不要）", value="parsed_log")

# ----------------------------
# ログ解析関数
# ----------------------------
def parse_iis_log(log_text):
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

    df_result = pd.DataFrame()
    try:
        df_result = pd.DataFrame({
            "datetime": df["date"] + " " + df["time"],
            "Account": df["_RequestID"].str.extract(r"@(.+)$")[0],
            "cs-method": df["cs-method"],
            "cs-uri-stem": df["cs-uri-stem"],
            "time-taken": df["time-taken"],
            "cs(User-Agent)": df["cs(User-Agent)"],
            "cs(Referer)": df["cs(Referer)"],
            "s-computername": df["s-computername"],
            "cs-host": df["cs-host"],
            "sc-status": df["sc-status"],
            "_RequestID": df["_RequestID"],
            "True-Client-IP": df["True-Client-IP"],
            "_X-SessionID": df["_X-SessionID"]
        })
    except Exception as e:
        st.error(f"データ整形中にエラーが発生しました: {e}")
        return pd.DataFrame()

    return df_result

# ----------------------------
# 実行処理
# ----------------------------
if uploaded_file and st.button("▶ 解析実行"):
    with st.spinner("ZIP解析中..."):
        all_dfs = []
        with zipfile.ZipFile(uploaded_file, "r") as zip_ref:
            for file_name in zip_ref.namelist():
                if file_name.endswith('.log') or file_name.endswith('.txt'):
                    with zip_ref.open(file_name) as f:
                        text = f.read().decode("utf-8", errors="ignore")
                        df = parse_iis_log(text)
                        if not df.empty:
                            all_dfs.append(df)

        df_all = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()

        if not df_all.empty:
            # Accountフィルタ処理
            accounts = [a.strip() for a in target_input.split(',') if a.strip()] if target_input else []
            if accounts:
                df_all = df_all[df_all["Account"].isin(accounts)]

            st.success(f"{len(df_all)} 件のレコードが見つかりました。")

            # Excel出力
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_all.to_excel(writer, index=False, sheet_name="ParsedLog")
            st.download_button("⬇ Excelファイルをダウンロード", data=output.getvalue(), file_name=f"{export_name}.xlsx")
        else:
            st.warning("解析結果が空です。ログ構造または対象Accountをご確認ください。")
