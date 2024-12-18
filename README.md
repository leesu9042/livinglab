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

## 요구사항

- Python 3.8 이상
- FastAPI
- pyhwp 라이브러리
- Uvicorn (API 서버 실행용)

