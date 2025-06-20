import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

def parse_iis_log(log_text):
    lines = log_text.strip().split('\n')
    field_line = [line for line in lines if line.startswith("#Fields:")][0]
    fields = field_line.replace("#Fields: ", "").split()
    data_lines = [line for line in lines if not line.startswith("#")]

    df = pd.read_csv(io.StringIO("\n".join(data_lines)), sep=' ', header=None, names=fields, engine='python')

    # 必要な列を抽出し、整形
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

# Streamlit UI
st.title("IISログ解析ツール")
st.write("IISログファイル（テキスト形式）をアップロードしてください。必要なフィールドを抽出してExcelに出力します。")

uploaded_file = st.file_uploader("ログファイルを選択", type=["log", "txt"]) 

if uploaded_file is not None:
    log_text = uploaded_file.read().decode("utf-8")
    df_output = parse_iis_log(log_text)

    st.success("ログを解析しました。下記に内容を表示します。")
    st.dataframe(df_output)

    # Excel出力
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_output.to_excel(writer, sheet_name='ParsedLog', index=False)
    st.download_button("Excelファイルをダウンロード", data=output.getvalue(), file_name="parsed_iis_log.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
