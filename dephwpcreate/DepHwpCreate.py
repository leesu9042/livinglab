import numpy as np
from pyhwpx import Hwp
import pandas as pd
import os
import re

''' 
인원에 명 붙이기
06 -> 6 이런식으로 바꾸기
'''
class EventTableProcessor:
    def __init__(self, template_path, dataframe,first_date_components,last_date_components):
        self.hwp = Hwp()
        self.template_path = template_path
        self.hwp.Open(template_path,arg="versionwarning:false")
        self.df = dataframe
        self.year = None
        self.date_counts = None
        self.first_date_components = first_date_components
        self.last_date_components = last_date_components
        self.start_year = ""
        self.end_year = ""



    # def reset_time_field(self, date_count):
    #     """테이블에서 time 필드를 초기화하는 메소드"""
    #     self.hwp.set_pos(15, 0, 0)
    #     self.hwp.TableCellBlock()
    #     self.hwp.TableCellBlockExtend()
    #     self.hwp.TableColPageDown()
    #     self.hwp.set_cur_field_name(" ")
    #
    #     for i in range(date_count):
    #         self.hwp.set_pos(15 + i * 13, 0, 0)
    #         self.hwp.set_cur_field_name("time")


    # 시간 column 전부다 선택하는 기능



    def time_field_initial(self):
        for i in range(7):
            self.hwp.set_pos(15 + i * 13, 0, 0)
            self.hwp.set_cur_field_name("time")
    def date_field_initial(self):
        self.hwp.set_pos(14, 0, 0)
        self.hwp.TableCellBlock()
        self.hwp.TableCellBlockExtend()
        self.hwp.TableColPageDown()
        self.hwp.set_cur_field_name("date")

    def table_clear(self):
        self.hwp.set_pos(14, 0, 0)
        self.hwp.TableCellBlock()
        self.hwp.TableCellBlockExtend()
        self.hwp.TableColEnd()
        self.hwp.TableColPageDown()
        self.hwp.TableDeleteCell(remain_cell=True)  # 셀 내용 전체 삭제
        self.hwp.delete_all_fields()  # 모든 필드 삭제

    def field_setting(self):
        self.table_clear()  # hwp field clear
        self.date_field_initial()  # date field initialize
        self.time_field_initial()  # time field initialize

    def load_and_preprocess_data(self):
        """데이터 프레임을 전처리하는 메소드"""
        # 컬럼 순서 정립
        # self.df = self.df[['date', 'time', 'eventName', 'place', 'personnel', 'department', 'note']]
        # self.df['personnel'] = self.df['personnel'].replace('', np.nan).fillna(0).astype(int).astype(str)  # 정수형 변환 후 문자열 처리
        # self.df['personnel'] = self.df['personnel'].replace(0, np.nan)
        self.df = self.df[['date', 'time', 'eventName', 'place', 'personnel', 'department', 'note']]

        # self.df['personnel'] = self.df['personnel'].replace('', np.nan).fillna(0).astype(int).astype(
        #     str)  # 정수형 변환 후 문자열 처리
        # self.df['personnel'] = self.df['personnel'].replace(0, np.nan)


        # NaN 값을 " "으로 채우기
        self.df = self.df.fillna(" ")

        # 데이터의 날짜 년도 추출
        date_df = self.df['date']
        start_date = date_df.iloc[0]
        end_date = date_df.iloc[-1]
        self.start_year = start_date.split(".")[0]
        self.end_year = end_date.split(".")[0]

        if self.start_year == self.end_year: self.end_year = ""
        # 만약 시작 date와 마지막 date의 년도가 같다면 end_year은 빈칸으로

        # date별로 카운트와 datetime 컬럼 추가하여 새로운 DataFrame 생성
        self.date_counts = self.df.groupby('date').size().reset_index(name='count')

        # datetime 컬럼 추가 및 정렬
        self.date_counts['datetime'] = pd.to_datetime(self.date_counts['date'].str.extract(r'(\d{4}\.\d{2}\.\d{2})')[0],
                                                      format='%Y.%m.%d')
        self.date_counts = self.date_counts.sort_values(by='datetime').reset_index(drop=True)

        self.date_counts['date'] = self.date_counts['date'].str.split('.', n=1).str[1]  # delete year
        self.date_counts['date'] = self.date_counts['date'].apply(
            self.format_date)  # date column change ex) 12.23 (월) > 12.23.(월)

        self.df['personnel'] = self.df['personnel'].replace('0', ' ')

        print("처음 전처리 \n")
        print(self.df)



    ...

    ...

    def extract_month_day(self,date_str):
        # 요일 부분을 제거하고, 년도 부분도 제거합니다.
        date_part = date_str.split(' ')[0]  # 공백을 기준으로 분리하고 날짜 부분만 추출합니다.
        month_day = date_part.split('.')[1:]  # 점을 기준으로 분리하고, 월과 일 부분만 추출합니다.
        # 월과 일 사이에 공간을 추가합니다.
        formatted_date = month_day[0] + ". " + month_day[1]
        return formatted_date

    @staticmethod
    def format_date(date_str):
        """
        날짜 형식을 변경하는 메서드.
        입력 형식: "12.23 (월)"
        출력 형식: "12. 23. (월)"
        """
        import re
        # 정규식으로 월과 일을 분리하고 점(.)과 공백을 추가
        match = re.match(r"(\d+)\.(\d+)\s*\((.+)\)", date_str)
        if match:
            month = match.group(1)  # "12"
            day = match.group(2)    # "23"
            day_of_week = match.group(3)  # "월"
            return f"{month}. {day}. ({day_of_week})"
        return date_str  # 잘못된 형식은 그대로 반환


    @staticmethod
    def _extract_month_day(date_str):
        match = re.search(r'(\d+)\.\s*(\d+)', date_str)
        if match:
            return int(match.group(1)), int(match.group(2))
        return None

    def get_first_last_date_components(self):
        """데이터프레임에서 첫 번째와 마지막 날짜의 월, 일, 요일을 추출하는 메소드"""
        first_date = self.df.iloc[0]["date"]
        last_date = self.df.iloc[-1]["date"]

        self.first_date_components = self.split_date_components(first_date)
        self.last_date_components = self.remove_leading_zeros(self.split_date_components(last_date))



        #print(self.first_date_components)

        return self.first_date_components, self.last_date_components

    @staticmethod
    def split_date_components(date_str):
        """정규식을 사용해 월과 일만 분리"""
        match = re.match(r"(\d+)\.(\d+)", date_str)
        if match:
            return match.group(1), match.group(2)
        print("날짜 형식이 올바르지 않습니다.")
        return None, None

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

    # 앞자리 0 제거 함수
    @staticmethod
    def remove_leading_zeros(date_str):
        try:
            # 정규식을 사용해 월과 일의 앞자리 0 제거
            cleaned_date = re.sub(r'\b0(\d)', r'\1', date_str)
            return cleaned_date
        except Exception as e:
            print(f"Error processing '{date_str}': {e}")
            return date_str

    # 날짜 형식을 변환하는 함수

    # 이벤트 이름 앞에 '·' 추가 함수
    def add_bullet_to_event(self, event_str):
        try:
            # 이벤트 이름이 있는 경우 앞에 '·' 추가
            if event_str.strip():
                return f"· {event_str.strip()}"
            return event_str
        except Exception as e:
            print(f"Error processing '{event_str}': {e}")
            return event_str

    # 'eventName' 열 앞에 '·' 추가

    def insert_dates(self):
        """날짜 데이터를 각 셀에 삽입하는 메소드"""
        print("data type: ")
        print(self.date_counts['date'].dtype)

        # # 강제 문자열 변환
        # self.date_counts['date'] = self.date_counts['date'].astype(str)

        # 날짜 변환 적용
        self.date_counts['date'] = self.date_counts['date'].apply(self.remove_leading_zeros)


        print("01 -> 1로 바꾼다")
        print(self.date_counts)

        for index, date in self.date_counts["date"].items():
            self.hwp.move_to_field(f'date{{{{{index}}}}}')
            self.hwp.insert_text(date)

    def insert_event_data(self):
        """date 컬럼을 제외한 데이터 삽입"""
        df_no_date = self.df.drop(columns=['date'])
        df_no_date['eventName'] = df_no_date['eventName'].apply(self.add_bullet_to_event)

        self.hwp.set_pos(15, 0, 0)
        self.hwp.TableCellBlock()
        self.hwp.TableCellBlockExtend()
        self.hwp.TableColEnd()
        self.hwp.TableColPageDown()

        # 선택된 셀의 필드 이름을 imsi로 변경 후 데이터 삽입
        self.hwp.set_cur_field_name("imsi")

        self.hwp.put_field_text("imsi", df_no_date.values.flatten().tolist())



    # def change_period(self):
    #     """처음 날짜와 마지막 날짜를 넣어주면 날짜를 변환"""
    #     if not self.first_date_components or not self.last_date_components:
    #         self.get_first_last_date_components()
    #
    #     self.hwp.set_pos(5, 0, 0)
    #     self.hwp.SelectAll()
    #     print(self.first_date_components)
    #     print(self.last_date_components)
    #
    #     # # 월, 일 추출
    #     # firstMonth = self.first_date_components[0] + ". "
    #     # firstDay = self.first_date_components[1] + ". "
    #     # lastMonth = self.last_date_components[0] + ". "
    #     # lastDay = self.last_date_components[1] + "."
    #
    #     first_date = self.first_date_components[0] + ". " + self.first_date_components[1] + ". "
    #
    #     last_date = self.last_date_components[0] + ". " + self.last_date_components[1] + "."
    #
    #     # 기간 계산
    #     # period = firstMonth + firstDay + "~ " + lastMonth + lastDay
    #
    #     # 한글 문서에 기간 삽입
    #
    #     if self.end_year == "": # end_year가 "" 라면 end_year = start_year
    #         self.hwp.insert_text(f"【{self.start_year}. {first_date}~ {last_date}】")
    #
    #     else:
    #         self.hwp.insert_text(f"【{self.start_year}. {first_date}~ {self.end_year}. {last_date}】")
    # def save_file(self, output_name="output.hwp"):
    #     """수정된 파일을 다운로드 폴더에 저장하는 메소드"""
    #     download_path = os.path.join(os.path.expanduser("~"), 'downloads', output_name)
    #     self.hwp.SaveAs(download_path)
    #     print(f"File saved as: {download_path}")
    #
    #     #이건 클라 관점에서 로컬디스크에 직접 저장하는 방식임

    def extract_year(self, date_str):
        # 년도 부분만 추출합니다.
        parts = date_str.split('.')  # 점을 기준으로 문자열을 분리합니다.
        year = parts[0].strip()  # 공백 제거
        return year



    def simplify_date(self, date_str):

        date_part = date_str.split(' ')[0]  # 공백을 기준으로 분리하고 날짜 부분만 추출합니다.
        month_day = date_part.split('.')[1:]  # 점을 기준으로 분리하고, 월과 일 부분만 추출합니다.
        # 월과 일 사이에 공간을 추가합니다.
        formatted_date = month_day[0] + ". " + month_day[1]
        return formatted_date


    def change_period(self):
        """처음 날짜와 마지막 날짜를 넣어주면 날짜를 변환"""
        if not self.first_date_components or not self.last_date_components:
            self.get_first_last_date_components()

        self.hwp.set_pos(5, 0, 0)
        self.hwp.SelectAll()


        firstDate = self.remove_leading_zeros(self.extract_month_day(self.first_date_components))
        endDate = self.remove_leading_zeros(self.extract_month_day(self.last_date_components))
        firstYear = self.remove_leading_zeros(self.extract_year(self.first_date_components))
        endYear = self.remove_leading_zeros(self.extract_year(self.last_date_components))

        print("firstData 전처리 재낀거")
        print(firstDate)



        print("firstdate ="+ firstDate)

        if firstYear == endYear :
            self.hwp.insert_text(f"【{firstYear}. {firstDate} ~ {endDate}】")
        else :
            # 한글 문서에 기간 삽입
            self.hwp.insert_text(f"【{firstYear}. {firstDate} ~ {endYear}. {endDate}】")


    def close(self):
        self.hwp.clear()
        self.hwp.Quit()  # HwpObject를 종료

    def process(self):
        """전체 처리 흐름을 한번에 실행하는 메서드"""
        # self.reset_time_field(7)
        self.field_setting()
        self.load_and_preprocess_data()
        self.adjust_table_rows()
        self.insert_dates()
        self.insert_event_data()
        self.change_period()
        #self.close()

        # self.save_file()
