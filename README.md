# hwp-create-api

**한글파일 생성 API**  
JSON 데이터를 기반으로 한글(HWP) 파일을 생성하는 API 서버입니다. 이 프로젝트는 FastAPI와 `pyhwp` 라이브러리를 사용합니다.

---

## 사용된 라이브러리

- **pyhwp**: [pyhwp PyPI 페이지](https://pypi.org/project/pyhwp/)  
  Licensed under the MIT License.  

---

## 프로젝트 소개

이 API 서버는 JSON 데이터를 입력받아, 지정된 형식에 따라 HWP(한글파일)를 생성합니다. 제공된 데이터는 이벤트 이름, 날짜, 시간, 장소, 인원수, 부서, 비고 등의 정보를 포함합니다.

---

## json data  예시
```json

{
    "data": [
        {
            "eventName": "테스트2",
            "date": "2000.11.13 (월)",
            "time": "17:25",
            "place": "대마고",
            "personnel": 1,
            "department": "토지정보과",
            "note": "23"
        },
        {
            "eventName": "방가방가",
            "date": "2000.11.13 (월)",
            "time": "09:33",
            "place": "구ㅜ구구구",
            "personnel": 12,
            "department": "건설과",
            "note": "방가방가"
        },
        {
            "eventName": "테스트",
            "date": "2000.11.17 (금)",
            "time": "17:25",
            "place": "대마고",
            "personnel": 1,
            "department": "문화관광체육과",
            "note": ""
        },
        {
            "eventName": "테스트3",
            "date": "2000.11.19 (일)",
            "time": "17:25",
            "place": "대마고등학교",
            "personnel": 9,
            "department": "일자리정책과",
            "note": ""
        }
    ],
    "startDate": "2000.11.13 (월)",
    "endDate": "2000.11.19 (일)"
}

```

---

## 요구사항

- Python 3.8 이상
- FastAPI
- pyhwp 라이브러리
- Uvicorn (API 서버 실행용)

