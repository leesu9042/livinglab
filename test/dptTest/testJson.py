import json
import pandas as pd

with open(r'../../testjsondata/package.json', 'r', encoding='utf-8') as f:

    json_data = json.load(f)

print(json.dumps(json_data, ensure_ascii=False, indent=4))
#dumps 함수 json을 문자열로 변환해서 출력
#ensure_ascii = False -> 한글 안깨지게


df = pd.DataFrame(json_data)

print(type(json_data))
# JSON을 읽어들이면 dict형태로 바뀐다.


print(df)





















# data = {
#     'name': 'nero',
#     'age': 2,
#     'color': 'black',
#     'like_food': ['banana', 'apple', 'chewrrr'],
#     'is_cat': True,
#     'is_dog': False,
#     'friends': {
#     	'poki': 'cat',
#         'bowow': 'dog',
#         'yatong': 'cat',
#     }
# }
#
# json_ex = json.dumps(data, ensure_ascii=False, indent=4)
# print(json_ex)