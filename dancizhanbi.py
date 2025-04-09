import pandas as pd
from collections import defaultdict
import re
import streamlit as st
from io import BytesIO
import jieba


def perform_chinese_analysis(data, min_length, use_jieba):
    word_count = defaultdict(lambda: [0, set()])
    chinese_pattern = re.compile(r'[\u4e00-\u9fa5]+')
    for text in data:
        text = str(text)
        if use_jieba:
            words = jieba.lcut(text)
            for word in words:
                if len(word) >= min_length:
                    word_count[word][0] += 1
                    word_count[word][1].add(text)
        else:
            matches = chinese_pattern.findall(text)
            for match in matches:
                for start in range(len(match)):
                    for end in range(start + 1, min(start + 11, len(match) + 1)):
                        word = match[start:end]
                        if len(word) >= min_length:
                            word_count[word][0] += 1
                            word_count[word][1].add(match)

    result = []
    for word, (count, original_words) in word_count.items():
        word_length = len(word)
        proportion = f"{(count / len(data)) * 100:.2f}%"
        original_words_str = ", ".join(original_words)
        result.append([word_length, word, count, proportion, '中文', original_words_str])
    return result


def perform_english_analysis(data, min_word_count):
    word_count = defaultdict(lambda: [0, set()])
    english_pattern = r'[a-zA-Z]+'
    for text in data:
        text = str(text)
        words = re.findall(english_pattern, text)
        for i in range(len(words)):
            for j in range(i + min_word_count, len(words) + 1):
                word = " ".join(words[i:j])
                word_count[word][0] += 1
                word_count[word][1].add(" ".join(words))

    result = []
    for word, (count, original_words) in word_count.items():
        word_length = len(word.split())
        proportion = f"{(count / len(data)) * 100:.2f}%"
        original_words_str = ", ".join(original_words)
        result.append([word_length, word, count, proportion, '英文', original_words_str])
    return result


def convert_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
    return output.getvalue()


def main():
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

    file = st.file_uploader("请上传Excel文件", type=["xlsx"])
    min_chinese_length = st.number_input("中文Token的最小长度（可调整）", min_value=1, value=3, step=1)
    min_english_word_count = st.number_input("英文Token的最小单词个数（可调整）", min_value=1, value=1, step=1)
    # 使用 st.radio 替换 st.checkbox
    jieba_option = st.radio("是否使用结巴分词进行中文分词", ["是", "否"])
    use_jieba = jieba_option == "是"

    if file:
        try:
            df = pd.read_excel(file)
        except Exception as e:
            st.error(f"读取文件时出现错误: {e}")
            return

        columns = df.columns.tolist()
        selected_column = st.selectbox("请选择要处理的列", columns)
        data = df[selected_column].tolist()

        chinese_result = perform_chinese_analysis(data, min_chinese_length, use_jieba)
        english_result = perform_english_analysis(data, min_english_word_count)

        all_results = chinese_result + english_result
        columns = ['token length', 'token', '词频', '占比', '语言类型', '原词']
        result_df = pd.DataFrame(all_results, columns=columns)
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
            excel_data = convert_to_excel(filtered_df)
            st.download_button(
                label='下载结果为Excel文件',
                data=excel_data,
                file_name='output.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )


if __name__ == "__main__":
    main()
    
