from datetime import datetime, timedelta

import pandas as pd


class DeptDataProcessor:
    def __init__(self, data):
        self.data_df = pd.DataFrame(data['data'])
        self.start_date = data['startDate']
        self.end_date = data['endDate']


    def process_data(self):
        # personnel을 문자열로 변환 후 소수점 제거
        self.data_df['personnel'] = (
            pd.to_numeric(self.data_df['personnel'], errors='coerce')
            .fillna(0)
            .astype(int)
            .astype(str)
        )
        # 'personnel' 열에서 숫자 문자열에 '명' 추가
        self.data_df['personnel'] = self.data_df['personnel'].apply(lambda x: f"{x}명" if x.strip().isdigit() else x)



        # 시작일과 종료일 파싱
        start_date_clean = datetime.strptime(self.start_date.split(" ")[0], "%Y.%m.%d")
        end_date_clean = datetime.strptime(self.end_date.split(" ")[0], "%Y.%m.%d")

        # 날짜 범위 생성
        full_date_range = [start_date_clean + timedelta(days=x) for x in range((end_date_clean - start_date_clean).days + 1)]

        # 기존 이벤트 날짜 추출
        existing_dates = {datetime.strptime(row["date"].split(" ")[0], "%Y.%m.%d") for _, row in self.data_df.iterrows()}

        # 누락된 이벤트 생성
        missing_dates = [date for date in full_date_range if date not in existing_dates]
        null_events = pd.DataFrame([{
            "eventName": None,
            "date": date.strftime("%Y.%m.%d (%a)"),
            "time": None,
            "place": None,
            "personnel": None,
            "department": None,
            "note": None
        } for date in missing_dates])

        # 데이터 병합 및 정렬
        full_data = pd.concat([self.data_df, null_events], ignore_index=True)
        full_data.sort_values(by="date", key=lambda x: pd.to_datetime(x.str.split(" ").str[0], format="%Y.%m.%d"), inplace=True)

        full_data["date"] = full_data["date"].apply(self.convert_day_to_korean)



        return full_data

    def convert_day_to_korean(self, date_str):
        date_part, day_part = date_str.split(" ")
        day_code = day_part[1:-1]  # 괄호 제거
        day_mapping = {"Mon": "월", "Tue": "화", "Wed": "수", "Thu": "목",
                       "Fri": "금", "Sat": "토", "Sun": "일"}

        # 한글 요일로 변환
        korean_day = day_mapping.get(day_code, day_code)
        print(self.data_df.columns)
        return f"{date_part} ({korean_day})"
