import os
import pandas as pd
from collections import defaultdict
import re
import streamlit as st
from io import BytesIO
import jieba


def process_chinese(data_set, weight_list, total_rows, min_length, use_jieba=False):
    chinese_word_count = defaultdict(lambda: [0, set()])
    chinese_pattern = re.compile(r'[\u4e00-\u9fa5]+')
    for i, line in enumerate(data_set):
        line = str(line)
        weight = weight_list[i]
        if use_jieba:
            # 使用结巴分词
            words = jieba.lcut(line)
            for word in words:
                if len(word) >= min_length:
                    chinese_word_count[word][0] += weight
                    chinese_word_count[word][1].add(line)
        else:
            chinese_matches = chinese_pattern.findall(line)
            for match in chinese_matches:
                for start in range(len(match)):
                    for end in range(start + 1, min(start + 11, len(match) + 1)):
                        word = match[start:end]
                        if len(word) >= min_length:
                            chinese_word_count[word][0] += weight
                            chinese_word_count[word][1].add(match)
    all_chinese_data = []
    for word, (count, original_words) in chinese_word_count.items():
        word_num = len(word)
        # 计算占比并转换为百分比形式
        proportion = f"{(count / total_rows) * 100:.2f}%"
        original_words_str = ", ".join(original_words)
        all_chinese_data.append([word_num, word, count, proportion, '中文', original_words_str])
    return all_chinese_data


def process_english(data_set, weight_list, total_rows, min_word_count):
    english_word_count = defaultdict(lambda: [0, set()])
    english_pattern = r'[a-zA-Z]+'
    for i, line in enumerate(data_set):
        line = str(line)
        weight = weight_list[i]
        words = re.findall(english_pattern, line)
        for i in range(len(words)):
            for j in range(i + min_word_count, len(words) + 1):
                word = " ".join(words[i:j])
                english_word_count[word][0] += weight
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
    # 设置页面标题和布局
    st.set_page_config(page_title="Token词频统计小工具", layout="centered")
    st.title('Token词频统计小工具')
    st.write('此工具可用于统计Excel文件中中英文Token的词频（次数和占比），并支持按不同条件筛选，如中英文、指定列。')
    with st.expander("使用示例", expanded=False):
        st.write('#### 中文示例')
        st.write('假设Excel文件某列中有文本“我爱学习”，设置中文分词最小长度为3。工具会将其拆分为“我爱学”、“爱学习”、“我爱学习”，并统计每个Token的出现次数和占比。')
        st.write('#### 英文示例')
        st.write('假设Excel文件某列中有文本“I love learning”，设置英文分词最小单词个数为1。工具会将其拆分为“I”、“love”、“learning”、“I love”、“love learning”、“I love learning”，并统计每个Token的出现次数和占比。')
        st.write('#### 中英混合示例')
        st.write('假设Excel文件某列中有文本“我爱学习I love learning”，设置中文分词最小长度为3，设置英文分词最小单词个数为3。工具会将其拆分为“我爱学”、“爱学习”、“我爱学习”、“I love learning”，并统计每个Token的出现次数和占比。')
        st.write('#### 权重功能示例')
        st.write('假设Excel文件中除了文本列外，还有一列权重值（如0.5, 1, 2等）。勾选"使用权重"后，选择该权重列。此时，某行文本生成的所有Token的词频都会乘以该行对应的权重值。例如，权重为2的行生成的Token，其词频将是普通情况的2倍。')

    uploaded_file = st.file_uploader("请上传Excel文件", type=["xlsx"])
    # 将中文分词的最小长度默认值设置为3
    min_chinese_length = st.number_input("中文Token的最小长度（可调整）", min_value=1, value=3, step=1)
    min_english_word_count = st.number_input("英文Token的最小单词个数（可调整）", min_value=1, value=1, step=1)
    # 添加是否使用结巴分词的选项
    use_jieba = st.checkbox("使用结巴分词进行中文分词")
    # 添加是否使用权重的选项
    use_weight = st.checkbox("使用权重")

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"读取文件时出现错误: {e}")
            return
        
        # 获取所有列名
        column_names = df.columns.tolist()
        
        # 找出所有数值类型的列，用于权重选择
        numeric_columns = []
        for col in column_names:
            try:
                # 检查该列是否可以转换为数值类型，且不全为NaN
                if not df[col].dropna().empty and pd.to_numeric(df[col], errors='coerce').notna().all():
                    numeric_columns.append(col)
            except:
                continue
        
        # 让用户选择列
        selected_column = st.selectbox("请选择要处理的列", column_names)
        data_set = df[selected_column].tolist()
        
        # 处理权重列选择
        weight_list = [1] * len(data_set)  # 默认权重为1
        if use_weight:
            if not numeric_columns:
                st.error("未找到合适的权重列。权重列应为全数字列（不包括标题）。")
                return
            weight_column = st.selectbox("请选择权重列", numeric_columns)
            try:
                weight_list = pd.to_numeric(df[weight_column], errors='coerce').fillna(1).tolist()
            except:
                st.error("无法将选中的列转换为权重值，请确保该列只包含数字。")
                return
        
        total_rows = sum(weight_list)  # 总权重作为总行数

        all_chinese_data = process_chinese(data_set, weight_list, total_rows, min_chinese_length, use_jieba)
        all_english_data = process_english(data_set, weight_list, total_rows, min_english_word_count)

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
                label='下载结果为Excel文件',
                data=df_xlsx,
                file_name='output.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )


if __name__ == "__main__":
    main()    
