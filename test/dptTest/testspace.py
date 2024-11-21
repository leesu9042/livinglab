import json
import pandas as pd

from dephwpcreate.DepHwpCreate import EventTableProcessor

with open(r'../../testjsondata/package.json', 'r', encoding='utf-8') as f:

    json_data = json.load(f)

#dumps 함수 json을 문자열로 변환해서 출력
#ensure_ascii = False -> 한글 안깨지게


#json파일 pandas dataframe으로 받아오기
df = pd.DataFrame(json_data['data'])

processor = EventTableProcessor("../../template/templateEx.hwp", df)

# 모든 작업을 한 번에 처리
processor.process()
