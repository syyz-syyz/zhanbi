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
        weight = weight_list[i]  # 直接获取处理后的权重值
        if use_jieba:
            words = jieba.lcut(line)
            for word in words:
                if len(word) >= min_length:
                    chinese_word_count[word][0] += weight  # 词频乘以权重
                    chinese_word_count[word][1].add(line)
        else:
            chinese_matches = chinese_pattern.findall(line)
            for match in chinese_matches:
                for start in range(len(match)):
                    for end in range(start + 1, min(start + 11, len(match) + 1)):
                        word = match[start:end]
                        if len(word) >= min_length:
                            chinese_word_count[word][0] += weight  # 词频乘以权重
                            chinese_word_count[word][1].add(match)
    all_chinese_data = []
    for word, (count, original_words) in chinese_word_count.items():
        word_num = len(word)
        # 占比计算基于总权重（total_rows已包含权重总和）
        proportion = f"{(count / total_rows) * 100:.2f}%"
        original_words_str = ", ".join(original_words)
        all_chinese_data.append([word_num, word, count, proportion, '中文', original_words_str])
    return all_chinese_data


def process_english(data_set, weight_list, total_rows, min_word_count):
    english_word_count = defaultdict(lambda: [0, set()])
    english_pattern = r'[a-zA-Z]+'
    for i, line in enumerate(data_set):
        line = str(line)
        weight = weight_list[i]  # 直接获取处理后的权重值
        words = re.findall(english_pattern, line)
        for i in range(len(words)):
            for j in range(i + min_word_count, len(words) + 1):
                word = " ".join(words[i:j])
                english_word_count[word][0] += weight  # 词频乘以权重
                english_word_count[word][1].add(" ".join(words))
    all_english_data = []
    for word, (count, original_words) in english_word_count.items():
        word_num = len(word.split())
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
    st.set_page_config(page_title="Token词频统计小工具", layout="centered")
    st.title('Token词频统计小工具')
    st.write('支持中英文Token词频统计，可配置分词长度、权重列（含空值自动转换）。')
    
    with st.expander("权重功能说明", expanded=False):
        st.write("- 权重列支持数字、空值、文本混合，空值/非数字会自动转为 `1`")
        st.write("- 词频 = 原始出现次数 × 权重值，占比 = (词频总和 / 总权重) × 100%")
    
    uploaded_file = st.file_uploader("请上传Excel文件", type=["xlsx"])
    min_chinese_length = st.number_input("中文Token最小长度", min_value=1, value=3, step=1)
    min_english_word_count = st.number_input("英文Token最小单词数", min_value=1, value=1, step=1)
    use_jieba = st.checkbox("使用结巴分词")
    use_weight = st.checkbox("启用权重列", help="支持含空值或非数字的列，自动转为1")

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file, keep_default_na=False)  # 空值转为空字符串
        except Exception as e:
            st.error(f"文件读取失败: {e}")
            return
        
        column_names = df.columns.tolist()
        selected_column = st.selectbox("选择文本列", column_names)
        data_set = df[selected_column].astype(str).tolist()  # 统一转为字符串处理
        
        weight_list = [1] * len(data_set)  # 默认权重为1
        if use_weight:
            weight_col = st.selectbox("选择权重列", column_names)
            # 强制转换权重列（空值/非数字转为1）
            try:
                weight_series = pd.to_numeric(df[weight_col], errors='coerce')  # 非数字转为NaN
                weight_list = weight_series.fillna(1).tolist()  # NaN转为1
            except:
                st.warning(f"权重列 '{weight_col}' 转换失败，自动使用默认权重1")
        
        total_rows = sum(weight_list)  # 总权重和作为分母
        
        # 处理中英文分词
        chinese_data = process_chinese(data_set, weight_list, total_rows, min_chinese_length, use_jieba)
        english_data = process_english(data_set, weight_list, total_rows, min_english_word_count)
        result_df = pd.DataFrame(chinese_data + english_data, 
                               columns=['token长度', 'token', '词频', '占比', '语言类型', '原词'])
        
        # 排序：占比降序 → token长度降序
        result_df = result_df.sort_values(by=['占比', 'token长度'], ascending=[False, False])
        # 筛选语言类型
        option = st.radio("筛选语言", ('全部', '中文', '英文'))
        filtered_df = result_df if option == '全部' else result_df[result_df['语言类型'] == option]
        
        st.dataframe(filtered_df.head(20), use_container_width=True, hide_index=True)
        
        if st.button('导出结果'):
            st.download_button(
                "下载Excel文件",
                to_excel(filtered_df),
                "词频统计结果.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


if __name__ == "__main__":
    main()
