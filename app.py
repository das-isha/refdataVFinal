import streamlit as st
import pandas as pd
import plotly.express as px
import base64
from io import StringIO, BytesIO

def generate_excel_download_link(df, filename="data_download.xlsx"):
    towrite = BytesIO()
    df.to_excel(towrite, encoding="utf-8", index=False, header=True)
    towrite.seek(0)
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Download Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

def generate_html_download_link(fig):
    towrite = StringIO()
    fig.write_html(towrite, include_plotlyjs="cdn")
    towrite = BytesIO(towrite.getvalue().encode())
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:text/html;charset=utf-8;base64, {b64}" download="plot.html">Download Plot</a>'
    return st.markdown(href, unsafe_allow_html=True)

st.set_page_config(page_title='Excel Plotter')
st.title('Excel Plotter 📈')
st.subheader('Feed me with your Excel file')

uploaded_file = st.file_uploader('Choose a XLSX file', type='xlsx')
if uploaded_file:
    st.markdown('---')
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    
    non_datetime_columns = [col for col in df.columns if not pd.api.types.is_datetime64_any_dtype(df[col])]
    df_display = df[non_datetime_columns]
    
    st.dataframe(df_display)
    
    groupby_column = st.selectbox(
        'What would you like to analyse?',
        non_datetime_columns
    )

    if groupby_column in non_datetime_columns:
        # -- GROUP DATAFRAME
        df_grouped = df.groupby(by=[groupby_column], as_index=False).sum()

        # -- PLOT DATAFRAME
        fig = px.bar(
            df_grouped,
            x=groupby_column,
            template='plotly_white',
            title=f'<b>{groupby_column} Analysis</b>'
        )
        st.plotly_chart(fig)

        # -- DOWNLOAD SECTION
        st.subheader('Downloads:')
        # Button for downloading grouped Excel file
        if st.button("Download Grouped Excel File"):
            generate_excel_download_link(df_grouped)
        # Button for downloading entire Excel file
        if st.button("Download Entire Excel File"):
            generate_excel_download_link(df)
        generate_html_download_link(fig)
    else:
        st.warning(f'The selected column "{groupby_column}" does not exist in the DataFrame.')
