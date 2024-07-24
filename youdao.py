import pandas as pd
import requests
import json
import re


df = pd.read_excel('CET4 RAW.xlsx')
for index, row in df.iterrows():

    cell_value = row[0]
    
    if isinstance(cell_value, str) and len(cell_value) >= 2:
        query_word = cell_value

        url = f"https://dict.youdao.com/result?word=lj%3A{query_word}&lang=en"
        response = requests.get(url)
        html_content = response.text

        pattern = re.compile(r'sentence-pair\":\[\{sentence:\"(.*?)\",\"sentence-eng\":\"(.*?)\",\"sentence-translation\":\"(.*?)\"')
        match = pattern.search(html_content)

        if match:
            eng_sent = match.group(1).replace("\\", "")
            chn_sent = match.group(3).replace("\\", "")
            df.at[index, '中文例句'] = chn_sent
            df.at[index, '英文例句'] = eng_sent
        else:
            print(f"无法找到单词 {query_word} 的例句")
            

df.to_excel('CET4 RAW_modified.xlsx', index=False)
