from pyhwpx import Hwp
import pandas as pd
import os
import re

class EventTableProcessor:
    def __init__(self, template_path, dataframe):
        self.hwp = Hwp()
        self.template_path = template_path
        self.hwp.Open(template_path)
        self.df = dataframe
        self.year = None
        self.date_counts = None
        self.first_date_components = None
        self.last_date_components = None





    # 시간 column 전부다 선택하는 기능
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
        """데이터 프레임을 전처리하는 메소드"""
        # 컬럼 순서 정립
        self.df = self.df[['date', 'time', 'eventName', 'place', 'personnel', 'department', 'note']]

        # NaN 값을 " "으로 채우기
        self.df = self.df.fillna(" ")

        # 가장 마지막 날짜의 년도 추출
        date_df = self.df['date']
        start_date = date_df.iloc[-1]
        self.year = "20" + start_date.split(".")[0]

        # 날짜 컬럼 처리 및 date_counts 생성
        self.df['date'] = self.df['date'].str.split('.', n=1).str[1]
        self.date_counts = self.df['date'].value_counts().reset_index()
        self.date_counts.columns = ['date', 'count']

        # 날짜 순으로 정렬
        self.date_counts[['month', 'day']] = self.date_counts['date'].apply(
            lambda x: pd.Series(self._extract_month_day(x)))
        self.date_counts = self.date_counts.sort_values(by=['month', 'day']).drop(columns=['month', 'day']).reset_index(drop=True)

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
        self.last_date_components = self.split_date_components(last_date)

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
        period = firstMonth + firstDay + "~ " + lastMonth + lastDay
        print(period)

        # 한글 문서에 기간 삽입
        self.hwp.insert_text(f"[{self.year}. {period}]")

    def save_file(self, output_name="output.hwp"):
        """수정된 파일을 다운로드 폴더에 저장하는 메소드"""
        download_path = os.path.join(os.path.expanduser("~"), 'downloads', output_name)
        self.hwp.SaveAs(download_path)
        print(f"File saved as: {download_path}")

        #이건 클라 관점에서 로컬디스크에 직접 저장하는 방식임

    def close(self):
        self.hwp.clear()
        self.hwp.Quit()  # HwpObject를 종료

    def process(self):
        """전체 처리 흐름을 한번에 실행하는 메서드"""
        self.reset_time_field(7)
        self.load_and_preprocess_data()
        self.adjust_table_rows()
        self.insert_dates()
        self.insert_event_data()
        self.change_period()

        # self.save_file()
