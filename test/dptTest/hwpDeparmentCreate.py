from pyhwpx import Hwp
import pandas as pd
import os
import re


class EventTableProcessor:
    def __init__(self, template_path, data_path):
        self.hwp = Hwp()
        self.template_path = template_path
        self.data_path = data_path
        self.hwp.Open(template_path,arg="versionwarning:false")
        self.df = None
        self.year = None
        self.date_counts = None

        self.first_date_components = None
        self.last_date_components = None



    def reset_time_field(self, date_count):
        """테이블에서 time 필드를 초기화하는 메소드"""
        self.hwp.set_pos(15, 0, 0)
        self.hwp.TableCellBlock()
        self.hwp.TableCellBlockExtend()
        self.hwp.TableColPageDown()
        self.hwp.set_cur_field_name("")

        for i in range(date_count):
            self.hwp.set_pos(15 + i * 13, 0, 0)
            self.hwp.set_cur_field_name("time")

    def load_and_preprocess_data(self):
        """엑셀 파일을 읽어와 전처리하는 메소드"""
        #대충 빈값 처리
        self.df = pd.read_excel(self.data_path)
        # self.df = self.df.replace(r'\n', '', regex=True).fillna(" ")
        self.df = self.df.fillna(" ")  # NaN 값을 " "으로 채우기

        # # 가장 마지막애 년도 뽑기
        # date_df = self.df['date']
        # start_date = date_df.iloc[-1]
        # # 첫 번째 점(.) 앞의 숫자만 추출
        # self.year = "20" + start_date.split(".")[0]
        #
        # # 날짜 컬럼 처리 년도삭제 +  date_counts컬럼 생성
        # self.df['date'] = self.df['date'].str.split('.', n=1).str[1]
        # self.date_counts = self.df['date'].value_counts().reset_index()
        # self.date_counts.columns = ['date', 'count']
        #
        # # 날짜 순으로 정렬된 df["date"] -> date_column_df으로
        # self.date_counts[['month', 'day']] = self.date_counts['date'].apply(
        #     lambda x: pd.Series(self._extract_month_day(x)))
        # self.date_counts = self.date_counts.sort_values(by=['month', 'day']).drop(columns=['month', 'day']).reset_index(
        #     drop=True)
        #
        # print("self.date_counts")
        if self.start_year == self.end_year: self.end_year = ""
        # 만약 시작 date와 마지막 date의 년도가 같다면 end_year은 빈칸으로

        # date별로 카운트와 datetime 컬럼 추가하여 새로운 DataFrame 생성
        self.date_counts = self.df.groupby('date').size().reset_index(name='count')

        # datetime 컬럼 추가 및 정렬
        self.date_counts['datetime'] = pd.to_datetime(self.date_counts['date'].str.extract(r'(\d{4}\.\d{2}\.\d{2})')[0],
                                                      format='%Y.%m.%d')
        self.date_counts = self.date_counts.sort_values(by='datetime').reset_index(drop=True)

        self.date_counts['date'] = self.date_counts['date'].str.split('.', n=1).str[1]  # delete year
        self.date_counts['date'] = self.date_counts['date'].apply(self.format_date)  # 날짜 형식 조정

    def format_date(self, date_str):
        """날짜 문자열 형식을 '월. 일.'에서 '월.  일.'로 변경합니다."""
        return re.sub(r"(\d+)\.(\d+)\.", r"\1.  \2.", date_str)

    def simplify_date(self , date_str):
        # 연도와 요일 부분을 제거합니다.
        parts = date_str.split('.')  # 먼저 점을 기준으로 문자열을 분리합니다.
        month_day = parts[1:3]  # 월과 일 부분만 추출합니다.
        clean_date = '.'.join(month_day).replace(' ', '')  # 공백을 제거하고 문자열을 다시 조합합니다.
        clean_date = clean_date[:5] + " " + clean_date[5:]  # 월과 일 사이에 공간을 추가합니다.
        return clean_date

    def get_first_last_date_components(self):
        """데이터프레임에서 첫 번째와 마지막 날짜의 월, 일, 요일을 추출하는 메소드"""
        date_df = self.df["date"]
        first_date = date_df.iloc[0]
        last_date = date_df.iloc[-1]

        # 각각의 날짜를 '월. 일.' 형식으로 변환
        first_date_components = self.split_date_components(first_date)
        last_date_components = self.split_date_components(last_date)

        return first_date_components, last_date_components

    @staticmethod
    def split_date_components(date_str):
        """날짜 문자열을 '월. 일.' 형식으로 변환"""
        result = re.sub(r"(\d+)\.(\d+)\s\([^)]+\)", r"\1. \2.", date_str)
        return result
    @staticmethod
    def _extract_month_day(date_str):
        match = re.search(r'(\d+)\.\s*(\d+)', date_str)
        if match:
            return int(match.group(1)), int(match.group(2))
        return None

    def adjust_table_rows(self):
        """날짜 별 데이터 개수에 맞춰 테이블 행을 조정하는 메소드"""
        for index, row in self.date_counts.iterrows():
            self.hwp.move_to_field(f'time{{{{{index}}}}}')
            count = row['count']

            if count > 2:
                for _ in range(count - 2):
                    self.hwp.TableAppendRow()
            elif count < 2:
                pset = self.hwp.HParameterSet.HTableDeleteLine
                self.hwp.TableLowerCell()
                self.hwp.HAction.GetDefault("TableDeleteRow", pset.HSet)
                self.hwp.HAction.Execute("TableDeleteRow", pset.HSet)


    def insert_dates(self):
        """날짜 데이터를 각 셀에 삽입하는 메소드"""
        for index, date in self.date_counts["date"].items():
            self.hwp.move_to_field(f'date{{{{{index}}}}}')
            self.hwp.insert_text(date)
    def insert_event_data(self):
        """date 컬럼을 제외한 데이터 삽입"""
        df_no_date = self.df.drop(columns=['date'])

        self.hwp.set_pos(15, 0, 0)
        self.hwp.TableCellBlock()
        self.hwp.TableCellBlockExtend()
        self.hwp.TableColEnd()
        self.hwp.TableColPageDown()

        # 선택된 셀의 필드 이름을 imsi로 변경 후 데이터 삽입
        self.hwp.set_cur_field_name("imsi")
        self.hwp.put_field_text("imsi", df_no_date.values.flatten().tolist())




    def save_file(self, output_name="output.hwp"):
        """수정된 파일을 다운로드 폴더에 저장하는 메소드"""
        download_path = os.path.join(os.path.expanduser("~"), 'downloads', output_name)
        self.hwp.SaveAs(download_path)
        print(f"File saved as: {download_path}")

#------------------------------------------------#
#period 삽입
    def split_date_components(self, date_str):
        """정규식을 사용해 월과 일만 분리"""
        match = re.match(r"(\d+)\.(\d+)", date_str)
        if match:
            month = match.group(1)  # 월
            day = match.group(2)  # 일
            return month, day
        else:
            print("날짜 형식이 올바르지 않습니다.")


    def get_first_last_date_components(self):
        """첫 번째와 마지막 날짜에서 월, 일, 요일을 분리"""
        # df의 첫 번째와 마지막 인덱스에서 date 값 추출
        first_date = self.df.iloc[0]["date"]  # 첫 번째 인덱스의 date 값
        last_date = self.df.iloc[-1]["date"]  # 마지막 인덱스의 date 값

        # 첫 번째와 마지막 날짜의 월, 일 분리
        self.first_date_components = self.split_date_components(first_date)
        self.last_date_components = self.split_date_components(last_date)

        return self.first_date_components, self.last_date_components

    def extract_year(self, date_str):
        # 년도 부분만 추출합니다.
        parts = date_str.split('.')  # 점을 기준으로 문자열을 분리합니다.
        year = parts[0].strip()  # 공백 제거
        return year

    def change_period(self):
        """처음 날짜와 마지막 날짜를 넣어주면 날짜를 변환"""
        if not self.first_date_components or not self.last_date_components:
            self.get_first_last_date_components()

        self.hwp.set_pos(5, 0, 0)
        self.hwp.SelectAll()

        # 월, 일 추출
        firstMonth = self.first_date_components[0] + ". "
        firstDay = self.first_date_components[1] + ". "
        lastMonth = self.last_date_components[0] + ". "
        lastDay = self.last_date_components[1] + "."

        # 기간 계산
        # period = firstMonth + firstDay + "~ " + lastMonth + lastDay
        # print(period)

        firstDate = self.simplify_date(self.first_date_components)
        endDate = self.simplify_date(self.last_date_components)
        firstYear = self.extract_year(self.first_date_components)
        endYear = self.extract_year(self.last_date_components)

        print(firstDate)

        if firstYear == endYear :
            self.hwp.insert_text(f"[{firstYear}. {firstDate} ~ {endDate}]")
        else :
            # 한글 문서에 기간 삽입
            self.hwp.insert_text(f"[{firstYear}. {firstDate} ~ {endYear}. {endDate}]")





    def process(self):
        self.reset_time_field(7)          # 초기화
        self.load_and_preprocess_data()    # 데이터 불러오기 및 전처리
        self.adjust_table_rows()           # 테이블 행 조정
        self.insert_dates()                # 날짜 삽입
        self.insert_event_data()           # 행사 데이터 삽입
        # 추출한 날짜로 기간 설정
        self.change_period()

        self.save_file()                   # 파일 저장


# run_processor.py

# def process(self):
#     # 전체 처리 흐름을 한번에 실행하는 메서드
#     self.reset_time_field(7)
#     self.load_and_preprocess_data()  # 데이터 불러오기 및 전처리
#     self.adjust_table_rows()  # 테이블 행 조정
#     self.insert_dates()  # 날짜 삽입
#     self.insert_event_data()  # 행사 데이터 삽입
#     # df에서 첫 번째와 마지막 날짜의 구성 요소 추출
#
#     first_date_components, last_date_components = self.get_first_last_date_components()
#
#     self.get_year_from_last_date()
#     # 연도 추출
#
#     # 추출한 날짜로 기간 설정
#     self.changePeriod( first_date_components, last_date_components)
#
#     # 파일 저장
#     self.save_file()

