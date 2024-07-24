import pandas as pd
import requests
import json
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

url='https://dict.youdao.com/jsonapi_s'

session = requests.Session()
retry = Retry(connect=3, backoff_factor=0.5)
adapter = HTTPAdapter(max_retries=retry)
session.mount('http://', adapter)
session.mount('https://', adapter)

session.get(url)


df = pd.read_excel('CET4 RAW.xlsx')
for index, row in df.iterrows():

    cell_value = row[0]
    
    if isinstance(cell_value, str) and len(cell_value) >= 2:
        query_word = cell_value
        
        res = requests.post(url, data={
            'q': query_word,
            'le': 'en'
        })
        
        res_json = json.loads(res.text)
        
        try:
            chn_sent = res_json['collins']['collins_entries'][0]['entries']['entry'][0]['tran_entry'][0]['exam_sents']['sent'][0]['chn_sent']
            eng_sent = res_json['collins']['collins_entries'][0]['entries']['entry'][0]['tran_entry'][0]['exam_sents']['sent'][0]['eng_sent']
            
            #print(chn_sent)
            #print(eng_sent)
            
            df.at[index, '中文例句'] = chn_sent
            df.at[index, '英文例句'] = eng_sent
        except (KeyError, IndexError):
            print(f"无法找到单词 {query_word} 的例句")
            

df.to_excel('CET4 RAW_modified.xlsx', index=False)
