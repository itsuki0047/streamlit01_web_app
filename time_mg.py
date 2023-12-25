import streamlit as st
import pandas as pd
from datetime import datetime


st.set_page_config(page_title="勤怠管理")
st.title("勤怠管理")

df = pd.read_excel("time.xlsx",sheet_name="Sheet1",index_col = 0)

def add_row():
    current_datetime = datetime.now()
    dt = current_datetime.replace(second=0,microsecond=0)
    
    with st.form(key = "form"):
        name = st.text_input("名前の入力")
        dep_sel = st.selectbox("部署の選択",(" ","営業","エンジニア", "スタッフ"))
        start_btn = st.form_submit_button("出勤")
        finish_btn = st.form_submit_button("退勤")

        if start_btn:
            print("aaa")
            df = pd.read_excel("time.xlsx",sheet_name = "Sheet1",index_col=0)
            ck = df.iloc[:,3][df.loc[:,"名前"]== name][df.loc[:,"部署"]== dep_sel]
            if  name != "" and dep_sel != " ":
                if ck.empty or ck.iloc[-1] == "退勤":
                    new_row = {"名前":name,"部署":dep_sel,"時間":dt, "出退勤":"出勤"}
                    df.loc[len(df)] = new_row
                    df.to_excel("time.xlsx")
                    st.table(df)
                else:
                    st.error("前回退勤されていません")
            else:
                st.error("空白があります")
        if finish_btn:
            df = pd.read_excel("time.xlsx",sheet_name = "Sheet1",index_col=0)
            ck = df.iloc[:,3][df.loc[:,"名前"]== name][df.loc[:,"部署"]== dep_sel]
            if  name != "" and dep_sel != " ":
                if ck.empty:
                    st.error("出勤されていません")
                elif  ck.iloc[-1] == "出勤":
                    wt0 = df.iloc[:,2][df.loc[:,"名前"]== name][df.loc[:,"部署"]== dep_sel]
                    wt0.number_format = "[h]:mm"
                    wt = dt - wt0.iloc[-1] 
                    wt.number_format = "[h]:mm"
                    new_row = {"名前":name,"部署":dep_sel,"時間":dt, "出退勤":"退勤","労働時間":wt}    
                    df.loc[len(df)] = new_row
                    df.to_excel("time.xlsx")
                    st.table(df)
                    
                else:
                    st.error("出勤されていません")
            else:
                st.error("空白があります")
                
def drop_row():
    drop_btn = st.button("削除")
    if drop_btn:
        global df
        for i in range(len(df)):
            df = df.drop(i)
        df = df.reset_index(drop=True)
        df.to_excel("time.xlsx")

def check():
    csv = df.to_csv()
    st.download_button(label = "Download", data = csv, file_name = "attendance.csv")
    drop_row()
    with st.expander("出勤記録",expanded = False):
        st.table(df)


def main():
    apps = {
        "出勤管理" : add_row,
        "確認" : check
    }
    
    selected_app_name = st.sidebar.selectbox(label="項目の選択",options = list(apps.keys()))
    
    if selected_app_name == "-":
        st.info("行追加")
        st.stop()
    
    render_func = apps[selected_app_name]
    render_func()


if __name__ == "__main__":
    main()
