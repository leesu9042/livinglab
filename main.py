# This project includes code from pyhwpx Library (https://pypi.org/project/pyhwpx/)
# Licensed under the MIT License
import base64
import io
from typing import Optional, Dict
import pandas as pd
from fastapi.responses import StreamingResponse
from fastapi import FastAPI
from dephwpcreate.DepHwpCreate import EventTableProcessor
from dept_data_processor.dateprocess import DeptDataProcessor
from dong_data_processor.dong_data_processor import DongDataProcessor
from donghwpcreate.DongHwpCreate import DongTableProcessor

'''
해야할일 - 한글파일 형식 다듬기 + 
서버 여는 코드 짜야됨

'''

# hwp 불러올떄 open이아니라 각각객체에 표를 복사해주고 필드값 초기화 함수 넣어주면 될 것 같은디.
# uvicorn main:app --reload --host 0.0.0.0 --port=8000
# http://localhost:8000/test/text/bytes/


app = FastAPI()





@app.get("/")
async def root():
    return {"message": "Hello World"}





@app.post("/depthwpcreate")
async def return_depthwp(data: Dict):
    # 'data'는 JSON 전체가 dict 형태로 들어옵니다
    # items = data['data']  # 'data' 키에 있는 값이 우리가 원하는 배열입니다

    # df_1 = pd.DataFrame(data['data'])
    # print("df_1 모양입니다")
    # print(df_1)
    # print("df_1 모양입니다")

    print(data)


    dept_data_processor = DeptDataProcessor(data)
    df = dept_data_processor.process_data()
    start_date = data["startDate"]
    end_date = data["endDate"]
    # pandas DataFrame으로 변환
    # df = pd.DataFrame(items)  # 'items'는 리스트이므로 바로 DataFrame으로 변환 가능
    # print(df)
    # print("end :" + end_date , "start" + start_date)

    processor = EventTableProcessor("template/templateEx.hwp", df,start_date,end_date)
    try:
        processor.process()
        # GetTextFile을 사용하여 HWP 파일을 BASE64로 인코딩된 문자열로 가져옴
        base64_content = processor.hwp.GetTextFile("HWP", "")

        # BASE64 문자열을 디코딩하여 바이트로 변환
        file_content = base64.b64decode(base64_content)

        # 메모리 스트림에 파일 내용을 저장
        memory_file = io.BytesIO(file_content)

    finally:
        # 리소스를 정리할 수 있는 부분이 있다면 여기에 작성

        processor.close()  # HwpObject를 종료
        del processor  # 메모리에서 객체 해제

        pass

    headers = {
        "Content-Disposition": "attachment; filename=download_test.hwp",
    }

    # StreamingResponse로 메모리 스트림 전송
    return StreamingResponse(memory_file, headers=headers, media_type="application/x-hwp")





@app.post("/donghwpcreate")
async def return_donghwp(data: Dict):
    # 'data'는 JSON 전체가 dict 형태로 들어옵니다
    # items = data['data']  # 'data' 키에 있는 값이 우리가 원하는 배열입니다
    #
    # # pandas DataFrame으로 변환
    # df = pd.DataFrame(items)  # 'items'는 리스트이므로 바로 DataFrame으로 변환 가능
    # print(df)

    d = DongDataProcessor(data)
    dong_df = d.process_data()
    print("처리 마침 ㅇㅇ")
    print(d)


    processor = DongTableProcessor("template/dongrealtemplate2.hwp", dong_df)
    try:
        processor.process()
        # GetTextFile을 사용하여 HWP 파일을 BASE64로 인코딩된 문자열로 가져옴
        base64_content = processor.hwp.GetTextFile("HWP", "")

        # BASE64 문자열을 디코딩하여 바이트로 변환
        file_content = base64.b64decode(base64_content)

        # 메모리 스트림에 파일 내용을 저장
        memory_file = io.BytesIO(file_content)

    finally:
        # 리소스를 정리할 수 있는 부분이 있다면 여기에 작성
        processor.close()  # HwpObject를 종료
        del processor  # 메모리에서 객체 해제
        pass

    headers = {
        "Content-Disposition": "attachment; filename=download_test.hwp",
    }

    # StreamingResponse로 메모리 스트림 전송
    return StreamingResponse(memory_file, headers=headers, media_type="application/x-hwp")