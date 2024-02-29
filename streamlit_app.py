import pandas as pd
import random
import streamlit as st
from io import BytesIO

def load_data(file_path):
    return pd.read_excel(file_path)

def random_categories(df, n=1, level_number='1', exclude_categories=None):
    level_mapping = {'1': '대분류', '2': '중분류', '3': '소분류', '4': '세분류'}
    level = level_mapping[level_number]
    levels = ['대분류', '중분류', '소분류', '세분류']
    level_index = levels.index(level)
    included_columns = levels[:level_index + 1]

    if exclude_categories and level in exclude_categories:
        df = df[~df[level].isin(exclude_categories[level])]

    if n <= df[level].nunique():
        available_categories = df[level].drop_duplicates()
        if level in exclude_categories:
            available_categories = available_categories[~available_categories.isin(exclude_categories[level])]
        unique_categories = available_categories.sample(n=n)
        filtered_df = df[df[level].isin(unique_categories)]
        filtered_df = filtered_df[included_columns].drop_duplicates()
    else:
        filtered_df = pd.DataFrame(columns=included_columns)

    return filtered_df

def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

def main():
    st.markdown("""
    <style>
    .title {
        font-family: 'Times New Roman', Times, serif;
        font-size: 24px;
    }
    </style>
    <div class="title">Naver DataLab Random Shopping Category Extractor</div>
    """, unsafe_allow_html=True)

    file_path = '/content/drive/MyDrive/Python/etc/네이버카테고리_modified.xls'
    df = load_data(file_path)

    if 'exclude_categories' not in st.session_state:
        st.session_state.exclude_categories = {}

    col1, col2 = st.columns(2)
    with col1:
        n = st.number_input("Enter the number of random categories you want:", min_value=1, value=1)
    with col2:
        level_number = st.selectbox("Select the category level:", options=['1', '2', '3', '4'], format_func=lambda x: {'1': '대분류', '2': '중분류', '3': '소분류', '4': '세분류'}[x])

    col_submit, col_reset = st.columns([1, 1])
    with col_submit:
        submit = st.button("Submit")
    with col_reset:
        reset = st.button("Reset")

    if submit:
        random_categories_output = random_categories(df, n, level_number, st.session_state.exclude_categories)
        if not random_categories_output.empty:
            st.table(random_categories_output)
            excel = convert_df_to_excel(random_categories_output)
            st.download_button(
                label="Download results as Excel",
                data=excel,
                file_name='random_categories.xlsx',
                mime='application/vnd.ms-excel'
            )
            level = {'1': '대분류', '2': '중분류', '3': '소분류', '4': '세분류'}[level_number]
            if level not in st.session_state.exclude_categories:
                st.session_state.exclude_categories[level] = []
            st.session_state.exclude_categories[level].extend(random_categories_output[level].tolist())
        else:
            st.write("No categories available for selection. Please reset.")

    if reset:
        st.session_state.exclude_categories.clear()

if __name__ == "__main__":
    main()

from pyngrok import ngrok

# Terminate open tunnels if exist
ngrok.kill()

# Setup a tunnel to the streamlit port 8501
public_url = ngrok.connect(8501)
print('Public URL:', public_url)
