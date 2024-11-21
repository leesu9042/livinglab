from donghwpcreate.DongHwpCreate import DongTableProcessor

import json
import pandas as pd
import os

current_directory = os.getcwd()
print("현재 작업 디렉터리:", current_directory)

with open(r'../../testjsondata/dongData.json', 'r', encoding='utf-8') as f:

    json_data = json.load(f)


# print(json.dumps(json_data, ensure_ascii=False, indent=4))
#dumps 함수 json을 문자열로 변환해서 출력
#ensure_ascii = False -> 한글 안깨지게


df = pd.DataFrame(json_data['data'])
processor = DongTableProcessor(r"../../template/dongrealtemplate2.hwp", df)
processor.process()