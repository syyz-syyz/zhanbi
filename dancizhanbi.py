import os
import pandas as pd
from collections import defaultdict
import re
import streamlit as st
from io import BytesIO


def process_chinese(data_set, total_rows, min_length):
    chinese_word_count = defaultdict(int)
    chinese_pattern = re.compile(r'[\u4e00-\u9fa5]+')
    for line in data_set:
        line = str(line)
        chinese_matches = chinese_pattern.findall(line)
        for match in chinese_matches:
            for start in range(len(match)):
                for end in range(start + 1, min(start + 11, len(match) + 1)):
                    word = match[start:end]
                    if len(word) >= min_length:
                        chinese_word_count[word] += 1
    all_chinese_data = []
    for word, count in chinese_word_count.items():
        word_num = len(word)
        # 计算占比并转换为百分比形式
        proportion = f"{(count / total_rows) * 100:.2f}%"
        all_chinese_data.append([word_num, word, count, proportion, '中文'])
    return all_chinese_data


def process_english(data_set, total_rows, min_word_count):
    english_word_count = defaultdict(int)
    english_pattern = r'\([a-zA-Z\s]+\)|[a-zA-Z]+'
    for line in data_set:
        line = str(line)
        english_matches = re.findall(english_pattern, line)
        for match in english_matches:
            if match.startswith('(') and match.endswith(')'):
                words = match[1:-1].strip().split()
                if len(words) >= min_word_count:
                    word = " ".join(words)
                    english_word_count[word] += 1
            else:
                words = match.split()
                if len(words) >= min_word_count:
                    word = " ".join(words)
                    english_word_count[word] += 1
    all_english_data = []
    for word, count in english_word_count.items():
        word_num = len(word.split())
        # 计算占比并转换为百分比形式
        proportion = f"{(count / total_rows) * 100:.2f}%"
        all_english_data.append([word_num, word, count, proportion, '英文'])
    return all_english_data


def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.close()
    processed_data = output.getvalue()
    return processed_data


def main():
    st.title('中英文分词结果展示')
    uploaded_file = st.file_uploader("请上传 Excel 文件", type=["xlsx"])
    min_chinese_length = st.number_input("请输入中文分词的最小长度", min_value=1, value=1, step=1)
    min_english_word_count = st.number_input("请输入英文分词的最小单词个数", min_value=1, value=1, step=1)

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"读取文件时出现错误: {e}")
            return
        # 获取所有列名
        column_names = df.columns.tolist()
        # 让用户选择列
        selected_column = st.selectbox("请选择要处理的列", column_names)
        data_set = df[selected_column].tolist()
        total_rows = len(data_set)

        all_chinese_data = process_chinese(data_set, total_rows, min_chinese_length)
        all_english_data = process_english(data_set, total_rows, min_english_word_count)

        all_data = all_chinese_data + all_english_data
        columns = ['单词个数', '词', '个数', '占比', '语言类型']
        result_df = pd.DataFrame(all_data, columns=columns)
        result_df = result_df.sort_values(by='占比', ascending=False)

        option = st.radio('选择显示结果类型', ('全部', '中文', '英文'))
        if option == '全部':
            filtered_df = result_df
        elif option == '中文':
            filtered_df = result_df[result_df['语言类型'] == '中文']
        else:
            filtered_df = result_df[result_df['语言类型'] == '英文']

        st.dataframe(filtered_df)

        if st.button('导出结果'):
            df_xlsx = to_excel(filtered_df)
            st.download_button(
                label='点击下载 Excel 文件',
                data=df_xlsx,
                file_name='output.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )


if __name__ == "__main__":
    main()
    
