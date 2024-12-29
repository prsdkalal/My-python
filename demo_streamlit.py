# import module
import streamlit as st
import pandas as pd
from mailmerge import MailMerge

# Title
st.title("IIIT Dharwad welcome!!!")
uploaded_temp = st.file_uploader("choose template file",type=["docx"])
if uploaded_temp is not None:
    document = MailMerge(uploaded_temp)
    st.write(document.get_merge_fields())


uploaded_file1 = st.file_uploader("Choose a file",type=['csv'])
if uploaded_file1 is not None:
    df = pd.read_csv(uploaded_file1)
    df.columns = df.columns.str.replace(".","")
    df.columns = df.columns.str.replace(" ","_")
    st.write(df.columns)

    dict_list = []
    for d in df['Date'].unique():
        df1 = df[df['Date'] == d]
        for i in df1['Course_Code'].unique():
            df2 = df1[df1['Course_Code'] == i]
            for j in df2['Room'].unique():
                dict = {}
                df3 = df2[df2['Room'] == j]
                df3 = df3.reset_index()
                df3 = df3.rename(columns = {"index":"SNo"})
                df3["SNo"] = df3.index + 1      

                dict['Course_Code']= i
                dict['Course_Name'] = df3.iloc[0,df3.columns.get_loc("Course_Name")]
                dict['Room'] = j
                dict['Date'] = d
                dict['SNo'] = df3[["SNo","Roll_No","Student_Name"]].to_dict(orient = "records")
                dict_list.append(dict)
                #print(len(dict_list))
    document.merge_templates(dict_list, separator='page_break')
    document.write('outputAN.docx')
    with open('outputAN.docx','rb') as f:
        st.download_button(label="attendance sheet document",data=f,file_name="outputAN.docx",mime="docx")











































































