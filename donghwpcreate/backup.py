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
        self.date_counts = None
        self.start_year = None
        self.end_year = None
        self.first_date_components = None
        self.last_date_components = None

    def table_clear(self):
        self.hwp.set_pos(11, 0, 0)
        self.hwp.TableCellBlock()
        self.hwp.TableCellBlockExtend()
        self.hwp.TableColEnd()
        self.hwp.TableColPageDown()
        self.hwp.TableDeleteCell(remain_cell=True)  # 셀 내용 전체 삭제
        self.hwp.delete_all_fields()  # 모든 필드 삭제

        # self.hwp.set_cur_field_name(" ")
        # self.hwp.insert_text("")

    def time_field_initial(self):
        self.hwp.set_pos(12, 0, 0)
        # self.hwp.TableCellBlock()
        # self.hwp.TableCellBlockExtend()
        # self.hwp.TableColPageDown()
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

        # self.df = pd.read_excel(self.df)
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
        self.date_counts['date'] = self.date_counts['date'].apply(self.format_date)  # date column change ex) 12.23 (월) > 12.23.(월)


        # #date_count 생성 date마다 데이터 갯수 보여주는 df
        # self.date_counts = self.df['date'].astype(str).str.split(' ').str[0]
        # self.date_counts = self.date_counts.str.replace('.', '-')  # .을 -로 교체
        # self.date_counts = pd.to_datetime(self.date_counts, format='%Y-%m-%d')  # 날짜 형식에 맞게 변환
        #
        # # 날짜별 빈도 계산
        # self.date_counts = pd.Series(self.date_counts).value_counts().reset_index()
        # self.date_counts.columns = ['date', 'count']
        #
        # # 날짜에서 연도, 월, 일 추출
        # self.date_counts[['year', 'month', 'day']] = self.date_counts['date'].apply(
        #     lambda x: pd.Series([x.year, x.month, x.day]))
        #
        # # 연도, 월, 일로 정렬
        # self.date_counts = self.date_counts.sort_values(by=['year', 'month', 'day']).reset_index(drop=True)

        # 결과 출력
        print(self.date_counts)

        # date 컬럼 전처리
        self.df['date'] = self.df['date'].str.split('.', n=1).str[1]  # delete year
        self.df['date'] = self.df['date'].apply(self.format_date)  # date column change ex) 12.23 (월) > 12.23.(월)
        #

        # self.date_counts = self.df['date'].value_counts().reset_index()
        # self.date_counts.columns = ['date', 'count']
        # self.date_counts[['month', 'day']] = self.date_counts['date'].apply(
        #     lambda x: pd.Series(self._extract_month_day(x)))
        # self.date_counts = self.date_counts.sort_values(by=['month', 'day']).drop(columns=['month', 'day']).reset_index(
        #     drop=True)
        #
        # print(self.date_counts)

    def _extract_month_day(self, date_str):
        match = re.search(r'(\d+)\.\s*(\d+)', date_str)
        if match:
            return int(match.group(1)), int(match.group(2))
        return None

    def get_first_last_date_components(self):
        """데이터프레임에서 첫 번째와 마지막 날짜의 월, 일, 요일을 추출하는 메소드"""
        first_date = self.df.iloc[0]["date"]
        last_date = self.df.iloc[-1]["date"]

        self.first_date_components = self.split_date_components(first_date)
        self.last_date_components = self.split_date_components(last_date)

        print(self.first_date_components, self.first_date_components)

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
        firstMonth = self.first_date_components[0] + ". "
        firstDay = self.first_date_components[1] + ". "
        lastMonth = self.last_date_components[0] + ". "
        lastDay = self.last_date_components[1] + "."

        # 기간 계산
        period = self.start_year + ". " + firstMonth + firstDay + "~ " + self.end_year + ". " + lastMonth + lastDay
        print(period)
        self.insert_period(period)

    # 지정된 위치에 period 넣는 함수
    def insert_period(self, period):
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

        if self.start_year == self.end_year:
            self.hwp.insert_text(f"【{period}】")

        else:
            self.hwp.MoveSelRight()
            self.hwp.MoveSelRight()
            self.hwp.MoveSelRight()
            self.hwp.MoveSelRight()
            self.hwp.MoveSelRight()
            self.hwp.MoveSelRight()
            self.hwp.MoveSelRight()

            self.hwp.insert_text(f"【{period}】")

    def insert_data(self):
        result_list = [day for day in self.date_counts['date']]
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
        print(self.df)
        # self.close()
