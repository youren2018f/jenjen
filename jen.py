#清洗資料
import re
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl.writer.excel import save_virtual_workbook

st.title('''
插播統計表處理
''')
st.header('上傳區')

uploaded_file = st.file_uploader("請上傳插播統計表xlsx檔", type = ".xlsx")

source_file = uploaded_file

if uploaded_file is not None:

    df = pd.read_excel(source_file, header=0)
    for i in range(len(df)):       
        str = df["插播名稱"][i]
        df["插播名稱"][i] = re.sub('\d{7}', '', str)


    def generate_vectors(vec_a, vec_b):
        check_dict = dict.fromkeys(set(vec_a) | set(vec_b), 0) #產生2個向量的交集的字典，且設定value為0
        #比對check_dict，如果有存在則+1
        import copy
        c1_copy = copy.deepcopy(check_dict)
        for a in vec_a:
            for key in c1_copy.keys():
                if a == key:
                    c1_copy[key] += 1
        consine_list_1 = [value for value in c1_copy.values()] #轉換為cosine_list
        c2_copy = copy.deepcopy(check_dict)
        for b in vec_b:
            for key in c2_copy.keys():
                if b == key:
                    c2_copy[key] += 1
        consine_list_2 = [value for value in c2_copy.values()] #轉換為cosine_list
        return (consine_list_1, consine_list_2)

    #計算Cosine Similarity
    def cosine_similarity(vec_ab):
        vec_a = vec_ab[0]
        vec_b = vec_ab[-1]
        # Dot and norm
        dot = sum(a*b for a, b in zip(vec_a, vec_b))
        norm_a = sum(a*a for a in vec_a) ** 0.5
        norm_b = sum(b*b for b in vec_b) ** 0.5

        # Cosine similarity
        cos_sim = dot / (norm_a*norm_b)
        return cos_sim

    def rename(inds):
        shortest_word = min(inds, key = len)
        re_name_list = []
        for word in shortest_word:
            matrix = []
            for ind in inds:
                if word in ind:
                    matrix.append(1)
                else:
                    matrix.append(0)
            if all(matrix):
                re_name_list.append(word)
        name = "".join(re_name_list)
        my_str = f"改為名稱:{name}"
        st.write(my_str)
        name_dict = {}
        name_dict[name] = inds
        for key,values in name_dict.items():
            for value in values:
                try:
                    df["插播名稱"] = df["插播名稱"].replace(value, key)
                except:
                    pass

    import jieba
    duplicate = [] #為了移除（1,3,4)之後的(3,4)
    for index_x in range(len(df)):
        same=[]
        if index_x not in set(duplicate):
            for index_y in range(index_x + 1, len(df)):  #避免出現(3,2)(2,3)及(3,3)的狀況
                #分詞並去除符號
                filter_elems = [" ", "(", ")", "_", "-", "、", "，", "《", "》"]
                vec_a = jieba.lcut(df['插播名稱'][index_x])
                for filter_elem in filter_elems:                
                    vec_a = [x for x in vec_a if x!= filter_elem]
                vec_b = jieba.lcut(df['插播名稱'][index_y])
                for filter_elem in filter_elems:                
                    vec_b = [x for x in vec_b if x!= filter_elem]

                #計算餘弦相似度
                vec_ab = generate_vectors(vec_a, vec_b)

                similarity = cosine_similarity(vec_ab)

                if similarity >= 0.67:
                    #print(df['插播名稱'][index_x], df['插播名稱'][index_y])
                    same.append(df['插播名稱'][index_x])
                    same.append(df['插播名稱'][index_y])
                    duplicate.append(index_x)
                    duplicate.append(index_y)

            if same:
                if len(set(same)) > 1:
                    st.write(set(same))
                    #st.write(set(same))#移除掉重覆的
            #print(same)
                

    # seg_list = jieba.lcut(sentence)


    my_dict = {}
    for index, row in df.iterrows():
        if row['插播名稱'] not in my_dict.keys():
            my_dict[row['插播名稱']] = row['播放次數']
        else:        my_dict[row['插播名稱']] += row['播放次數']


    df2 = pd.DataFrame(list(my_dict.items()),columns = ['插播名稱','播放次數']) 
    st.write("檔案下載預覽")
    st.table(df2)

    import openpyxl
    wb = openpyxl.load_workbook(r"blank.xlsx")
    #指定那一個worksheet
    ws =wb['統計結果']
    #開始疊代
    i = 2 #從第幾列開始
    for ind in df2.index:
        ws.cell(row=i, column=1).value = df2.loc[ind, "插播名稱"]
        ws.cell(row=i, column=2).value = df2.loc[ind, "播放次數"]
        i = i + 1
    data = BytesIO(save_virtual_workbook(wb))
    st.header('下載區')
    st.download_button("下載檔案",
        data=data,
        mime='xlsx',
        file_name="插播統計表_修改過.xlsx")

