import numpy as np
from pyhwpx import Hwp
import pandas as pd
import re


# 추가 알고리즘 바꿔야함

class DongTableProcessor:
    def __init__(self, template_path, dataframe):
        self.hwp = Hwp()
        self.template_path = template_path
        self.df = dataframe
        self.hwp.Open(template_path, arg="versionwarning:false")
        self.date_counts = None  # 이걸로 표의 date컬럼 넣고 , period 교체 + 표 생성
        self.start_year = None
        self.end_year = None
        self.first_date_components = None
        self.last_date_components = None

    def get_hwp(self):
        """Hwp 객체를 리턴하는 메서드"""
        return self.hwp

    def table_clear(self):
        self.hwp.set_pos(11, 0, 0)
        self.hwp.TableCellBlock()
        self.hwp.TableCellBlockExtend()
        self.hwp.TableColEnd()
        self.hwp.TableColPageDown()
        self.hwp.TableDeleteCell(remain_cell=True)  # 셀 내용 전체 삭제
        self.hwp.delete_all_fields()  # 모든 필드 삭제

    def time_field_initial(self):
        self.hwp.set_pos(12, 0, 0)
        self.hwp.set_cur_field_name("time")

    def date_field_initial(self):
        for i in range(7):
            self.hwp.set_pos(11 + i * 14, 0, 0)
            self.hwp.set_cur_field_name("date")

    def field_setting(self):
        self.table_clear()  # hwp field clear
        self.date_field_initial()  # date field initialize
        self.time_field_initial()  # time field initialize

    #   field initialize

    @staticmethod
    def format_date(date_str):
        # 공백을 제거하고 괄호 앞에 '.' 추가
        return re.sub(r'\s*\((\S+)\)', r'.(\1)', date_str)

    def load_and_preprocess_data(self):
        # 컬럼 순서 정렬
        self.df = self.df[['date', 'time', 'dong', 'eventName', 'place', 'personnel', 'phone']]
        # self.df['personnel'] = self.df['personnel'].replace('', np.nan).fillna(0).astype(int).astype(
        #     str)  # 정수형 변환 후 문자열 처리
        # self.df['personnel'] = self.df['personnel'].replace(0, np.nan)

        # self.df = pd.read_excel(self.df)
        # NaN 값을 " "으로 채우기
        self.df = self.df.fillna(" ")
        self.df = self.df.replace(r'\n', '', regex=True).fillna(" ")

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

        # 결과 출력
        print("dong df")
        print(self.df)

    def _extract_month_day(self, date_str):
        match = re.search(r'(\d+)\.\s*(\d+)', date_str)
        if match:
            return int(match.group(1)), int(match.group(2))
        return None

    @staticmethod
    def remove_leading_zeros(date_str):
        try:
            # 정규식을 사용해 월과 일의 앞자리 0 제거
            cleaned_date = re.sub(r'\b0(\d)', r'\1', date_str)
            return cleaned_date
        except Exception as e:
            print(f"Error processing '{date_str}': {e}")
            return date_str

    def get_first_last_date_components(self):
        """데이터프레임에서 첫 번째와 마지막 날짜의 월, 일, 요일을 추출하는 메소드"""
        first_date = self.date_counts.iloc[0]["date"]
        last_date = self.date_counts.iloc[-1]["date"]

        self.first_date_components = self.split_date_components(first_date)
        self.last_date_components = self.split_date_components(last_date)

        # print(self.first_date_components, self.first_date_components)

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
        """ 표 데이터 갯수에 맞게 생성"""
        for index, row in self.date_counts.iterrows():
            pset = self.hwp.HParameterSet.HTableDeleteLine
            self.hwp.move_to_field(f'date{{{{{index}}}}}')
            count = row['count']

            if count > 2:
                for _ in range(count - 2):
                    self.hwp.TableAppendRow()
            elif count < 2:
                self.hwp.TableLowerCell()
                self.hwp.HAction.GetDefault("TableDeleteRow", pset.HSet)
                self.hwp.HAction.Execute("TableDeleteRow", pset.HSet)

    def change_period(self):
        """처음 날짜와 마지막 날짜를 넣어주면 날짜를 변환"""
        if not self.first_date_components or not self.last_date_components:
            self.get_first_last_date_components()

        self.hwp.set_pos(5, 0, 0)
        self.hwp.SelectAll()

        # 월, 일 추출
        first_date = self.first_date_components[0] + ". " + self.first_date_components[1] + ". "

        last_date = self.last_date_components[0] + ". " + self.last_date_components[1] + "."

        # 기간 계산
        # period = self.start_year + ". " + firstMonth + firstDay + "~ " + next_year + lastMonth + lastDay
        # print(period)
        self.insert_period(self.remove_leading_zeros(first_date), self.remove_leading_zeros(last_date))

    # 지정된 위치에 period 넣는 함수
    def insert_period(self, first_date, last_date):
        self.hwp.set_pos(0, 1, 0)
        while self.hwp.MoveSelRight():
            text = self.hwp.get_selected_text()  # 파이썬으로 가져와서

            if '】' in text: break  # 그냥 대괄호랑은 다름

        '''
        moveselRight는 shift -> 세번 시키는것인데  이걸 안하면 맨 우측의 (유성구)가 다음줄로밀려남
        '''
        self.hwp.MoveSelRight()
        self.hwp.MoveSelRight()
        self.hwp.MoveSelRight()

        if self.end_year == "": # end_year가 "" 라면 end_year = start_year
            self.hwp.insert_text(f"【{self.start_year}. {first_date}~ {last_date}】")

        else:
            self.hwp.MoveSelRight()
            self.hwp.MoveSelRight()
            self.hwp.MoveSelRight()
            self.hwp.MoveSelRight()
            self.hwp.MoveSelRight()
            self.hwp.MoveSelRight()
            self.hwp.MoveSelRight()

            self.hwp.insert_text(f"【{self.start_year}. {first_date}~ {self.end_year}. {last_date}】")



    def insert_data(self):
        result_list = [day for day in self.date_counts['date']]

        # result_list 값마다 remove_leading_zeros 적용
        result_list = [self.remove_leading_zeros(day) for day in self.date_counts['date']]
        self.hwp.put_field_text("date", result_list)



        df_no_date = self.df.drop(columns=['date'])
        self.hwp.set_pos(12, 0, 0)
        self.hwp.TableCellBlock()
        self.hwp.TableCellBlockExtend()
        self.hwp.TableColEnd()
        self.hwp.TableColPageDown()
        self.hwp.set_cur_field_name("imsi")
        self.hwp.put_field_text("imsi", df_no_date.values.flatten().tolist())

    def close(self):
        self.hwp.clear()
        self.hwp.Quit()  # HwpObject를 종료




    def process(self):
        self.field_setting()
        self.load_and_preprocess_data()
        self.adjust_table_rows()
        self.change_period()
        self.insert_data()

