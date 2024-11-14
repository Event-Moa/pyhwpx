from pyhwpx import Hwp
import win32com.client as win32
import pandas as pd
import os
import re

class DongTableProcessor:
    def __init__(self, template_path, data_path):
        self.hwp = Hwp()
        self.template_path = template_path
        self.data_path = data_path
        self.hwp.Open(template_path)
        self.df = None
        self.date_counts = None

    def table_clear(self):
        """???? ????? ???"""
        self.hwp.set_pos(11, 0, 0)
        self.hwp.TableCellBlock()
        self.hwp.TableCellBlockExtend()
        self.hwp.TableColEnd()
        self.hwp.TableColPageDown()
        self.hwp.set_cur_field_name(" ")

    def date_column_initial(self):
        """date ??? ????? ???"""
        self.hwp.set_pos(11, 0, 0)
        self.hwp.TableCellBlock()
        self.hwp.TableCellBlockExtend()
        self.hwp.TableColPageDown()
        self.hwp.set_cur_field_name("date")

    def time_column_initial(self):
        """time ??? ????? ???"""
        self.hwp.set_pos(12, 0, 0)
        self.hwp.TableCellBlock()
        self.hwp.TableCellBlockExtend()
        self.hwp.TableColPageDown()
        self.hwp.set_cur_field_name("")
        for i in range(7):
            self.hwp.set_pos(12 + i * 13, 0, 0)
            self.hwp.set_cur_field_name("time")

    def field_setting(self):
        """??? ???? ???"""
        self.table_clear()           # ??? ???
        self.date_column_initial()   # date ?? ???
        self.time_column_initial()   # time ?? ???

    def load_and_preprocess_data(self):
        """?? ??? ??? ????? ???"""
        self.df = pd.read_excel(self.data_path)
        self.df = self.df.replace(r'\n', '', regex=True).fillna(" ")
        self.date_counts = self.df['date'].value_counts().reset_index()
        self.date_counts.columns = ['date', 'count']
        self.date_counts[['month', 'day']] = self.date_counts['date'].apply(
            lambda x: pd.Series(self._extract_month_day(x)))
        self.date_counts = self.date_counts.sort_values(by=['month', 'day']).drop(columns=['month', 'day']).reset_index(drop=True)

    def _extract_month_day(self, date_str):
        """?? ????? ?? ?? ???? ???"""
        match = re.search(r'(\d+)\.\s*(\d+)', date_str)
        if match:
            return int(match.group(1)), int(match.group(2))
        return None

    def adjust_table_rows(self):
        """??? ??? ??? ?? ??? ?? ???? ???"""
        for index, row in self.date_counts.iterrows():
            pset = self.hwp.HParameterSet.HTableDeleteLine
            self.hwp.move_to_field(f'time{{{{{index}}}}}')
            count = row['count']

            if count > 2:
                for _ in range(count - 2):
                    self.hwp.TableAppendRow()
            elif count < 2:
                self.hwp.TableLowerCell()
                self.hwp.HAction.GetDefault("TableDeleteRow", pset.HSet)
                self.hwp.HAction.Execute("TableDeleteRow", pset.HSet)

    def insert_data(self):
        """???? ? ?? ???? ???"""
        # ?? ??? ??
        result_list = [day for day in self.date_counts['date']]
        self.hwp.put_field_text("date", result_list)

        # ??? ??? ??? ??? ??
        df_no_date = self.df.drop(columns=['date'])
        self.hwp.set_pos(12, 0, 0)
        self.hwp.TableCellBlock()
        self.hwp.TableCellBlockExtend()
        self.hwp.TableColEnd()
        self.hwp.TableColPageDown()
        self.hwp.set_cur_field_name("imsi")
        self.hwp.put_field_text("imsi", df_no_date.values.flatten().tolist())

    def save_file(self, output_name="dong_result.hwp"):
        """??? ??? ???? ???"""
        self.hwp.save_as(output_name)
        print(f"File saved as: {output_name}")

    def process(self):
        """?? ????? ???? ???"""
        self.field_setting()         # ?? ?? ???
        self.load_and_preprocess_data()  # ??? ???? ? ???
        self.adjust_table_rows()     # ??? ? ??
        self.insert_data()           # ??? ??
        self.save_file()             # ?? ??


"""# ?? ??
if __name__ == "__main__":
    template_path = r"C:\Users\thdco\OneDrive\Documents\GitHub\pyhwpx\dongtemplate.hwp"
    data_path = r"C:\Users\thdco\OneDrive\Documents\GitHub\pyhwpx\dongdummy.xlsx"
    table_processor = DongTableProcessor(template_path, data_path)
    table_processor.process()"""
    
