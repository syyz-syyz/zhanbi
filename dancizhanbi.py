import os
import pandas as pd
from collections import defaultdict
import re
import streamlit as st
from io import BytesIO


def process_chinese(data_set, total_rows, min_length):
    chinese_word_count = defaultdict(lambda: [0, set()])
    chinese_pattern = re.compile(r'[\u4e00-\u9fa5]+')
    for line in data_set:
        line = str(line)
        chinese_matches = chinese_pattern.findall(line)
        for match in chinese_matches:
            for start in range(len(match)):
                for end in range(start + 1, min(start + 11, len(match) + 1)):
                    word = match[start:end]
                    if len(word) >= min_length:
                        chinese_word_count[word][0] += 1
                        chinese_word_count[word][1].add(match)
    all_chinese_data = []
    for word, (count, original_words) in chinese_word_count.items():
        word_num = len(word)
        # 计算占比并转换为百分比形式
        proportion = f"{(count / total_rows) * 100:.2f}%"
        original_words_str = ", ".join(original_words)
        all_chinese_data.append([word_num, word, count, proportion, '中文', original_words_str])
    return all_chinese_data


def process_english(data_set, total_rows, min_word_count):
    english_word_count = defaultdict(lambda: [0, set()])
    english_pattern = r'[a-zA-Z]+'
    for line in data_set:
        line = str(line)
        words = re.findall(english_pattern, line)
        for i in range(len(words)):
            for j in range(i + min_word_count, len(words) + 1):
                word = " ".join(words[i:j])
                english_word_count[word][0] += 1
                english_word_count[word][1].add(" ".join(words))
    all_english_data = []
    for word, (count, original_words) in english_word_count.items():
        word_num = len(word.split())
        # 计算占比并转换为百分比形式
        proportion = f"{(count / total_rows) * 100:.2f}%"
        original_words_str = ", ".join(original_words)
        all_english_data.append([word_num, word, count, proportion, '英文', original_words_str])
    return all_english_data


def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.close()
    processed_data = output.getvalue()
    return processed_data


def main():
    st.title('token词频统计小工具')
    st.write('此工具可用于统计 Excel 文件中中英文 token 的词频，并支持按不同条件筛选和导出结果。具体来说，它会对指定列中的文本数据进行处理，分别统计中文和英文 token 的出现次数和占比。')
    st.write('### 使用示例')
    st.write('#### 中文示例')
    st.write('假设 Excel 文件某列中有文本 “我爱学习”，设置中文分词最小长度为 3。工具会将其拆分为 “我爱学”、“爱学习”、“我爱学习”，并统计每个 token 的出现次数和占比。')
    st.write('#### 英文示例')
    st.write('假设 Excel 文件某列中有文本 “I love learning”，设置英文分词最小单词个数为 1。工具会将其拆分为 “I”、“love”、“learning”、“I love”、“love learning”、“I love learning”，并统计每个 token 的出现次数和占比。')
    st.write('#### 中英混合示例')
    st.write('假设 Excel 文件某列中有文本 “我爱学习 I love learning”，设置中文分词最小长度为 3，设置英文分词最小单词个数为 3。工具会将其拆分为 “我爱学”、“爱学习”、“我爱学习”、“I love learning”，并统计每个 token 的出现次数和占比。')

    uploaded_file = st.file_uploader("请上传 Excel 文件", type=["xlsx"])
    # 将中文分词的最小长度默认值设置为 3
    min_chinese_length = st.number_input("请输入中文分词的最小长度", min_value=1, value=3, step=1)
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
        columns = ['token length', 'token', '词频', '占比', '语言类型', '原词']
        result_df = pd.DataFrame(all_data, columns=columns)
        # 按“占比”降序排序，当“占比”相同时，按“token length”降序排序
        result_df = result_df.sort_values(by=['占比', 'token length'], ascending=[False, False])

        option = st.radio('选择输入和输出文字类型', ('全部', '中文', '英文'))
        if option == '全部':
            filtered_df = result_df
        elif option == '中文':
            filtered_df = result_df[result_df['语言类型'] == '中文']
        else:
            filtered_df = result_df[result_df['语言类型'] == '英文']

        st.dataframe(filtered_df.head(20), use_container_width=True, hide_index=True)

        if st.button('导出文件'):
            df_xlsx = to_excel(filtered_df)
            st.download_button(
                label='下载结果为 Excel 文件',
                data=df_xlsx,
                file_name='output.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )


if __name__ == "__main__":
    main()
