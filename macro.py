import pandas as pd
import numpy as np
import os
import sys
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QGridLayout, QFileDialog, QLabel






class QtGUI(QWidget):

    def __init__(self):
        super().__init__()
        self.num = 0
        self.setWindowTitle("엑셀 매크로")
        self.resize(200, 200)
        self.qclist = []
        self.position = 0
        self.Lgrid = QGridLayout()
        self.setLayout(self.Lgrid)
        self.label1 = QLabel('', self)
        # self.label2 = QLabel('', self)
        addbutton1 = QPushButton('Open File', self)
        self.Lgrid.addWidget(self.label1, 1, 1)
        self.Lgrid.addWidget(addbutton1, 2, 1)
        addbutton1.clicked.connect(self.add_open)
        # addbutton2 = QPushButton('Divide Sheet', self)
        # self.Lgrid.addWidget(self.label2, 3, 1)
        # self.Lgrid.addWidget(addbutton2, 4, 1)
        # addbutton2.clicked.connect(self.div_sheet)
        self.show()

    def AutoFitColumnSize(self, worksheet, columns=None, margin=2):
        for i, column_cells in enumerate(worksheet.columns):
            is_ok = False
            if columns == None:
                is_ok = True
            elif isinstance(columns, list) and i in columns:
                is_ok = True
            if is_ok:
                length = max(len(str(cell.value)) for cell in column_cells)
                worksheet.column_dimensions[column_cells[0].column_letter].width = length + margin
        return worksheet

    def add_open(self):
        self.FileOpen = QFileDialog.getOpenFileName(self, 'Open file', './')
        self.label1.setText(self.FileOpen[0])
        self.wb = openpyxl.load_workbook(self.FileOpen[0])
        self.rowDf = pd.read_excel(io=self.FileOpen[0])
        self.rowDf['실결제금액'] = self.rowDf['실결제금액'].fillna(0)
        self.rowDf = self.rowDf.astype({'상호': 'string', '대표자명': 'string',
                                        '주소': 'string', '합계금액': 'int64',
                                        '공급가액': 'int64', '세액': 'int64',
                                        '품목명': 'string','비고(업체정보)': 'string',
                                        '실결제금액': 'int64', '결제수단': 'string',
                                        '차액': 'int64', '현장': 'string'
                                        })
        self.rowDf['작성일자(년)'] = self.rowDf['작성일자'].dt.year
        self.rowDf['작성일자(월)'] = self.rowDf['작성일자'].dt.month
        self.groupbydf = self.rowDf.groupby(['작성일자(년)', '작성일자(월)', '공급자사업자등록번호', '상호', '대표자명']).agg('sum')
        self.df_static = pd.DataFrame(self.groupbydf)
        self.df_static = self.df_static.reset_index()
        years = set(self.df_static['작성일자(년)'].values)
        for year in years:
            static = []
            year_df = self.df_static[self.df_static['작성일자(년)'] == year]
            enterprises_nums = set(year_df['공급자사업자등록번호'].values)
            for enterprise_num in enterprises_nums:
                # 사업자번호
                enterprise_num_df = year_df[year_df['공급자사업자등록번호'] == enterprise_num]
                dataset = np.zeros(shape=(19,), dtype=object)
                index = enterprise_num_df['공급자사업자등록번호'].index[0]
                np.put(dataset, [0], enterprise_num_df['공급자사업자등록번호'][index])
                np.put(dataset, [1], enterprise_num_df['상호'][index])
                np.put(dataset, [2], enterprise_num_df['대표자명'][index])
                real_total = 0
                for idx, row in enumerate(enterprise_num_df.iterrows()):
                    row = row[1]
                    real_total += row['실결제금액']
                    if idx == 0:
                        np.put(dataset, [row['작성일자(월)'] + 3], row['합계금액'])
                        np.put(dataset, [16], row['실결제금액'])
                    else:
                        np.put(dataset, [row['작성일자(월)'] + 3], row['합계금액'])
                        np.put(dataset, [16], real_total)
                static.append(dataset)

            self.df_result = pd.DataFrame(static, columns=['사업자번호', '상호', '대표자', '사기성총액', '1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월', '지불소계', '대비', '비고(업체정보)'])
            self.df_result['사기성총액'] = self.df_result[['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월']].sum(axis=1)
            # self.df_result['사기성총액'] = self.df_result['사기성총액'].apply(lambda x : format(x, ','))
            # self.df_result['1월'] = self.df_result['1월'].apply(lambda x: format(x, ','))
            # self.df_result['2월'] = self.df_result['2월'].apply(lambda x: format(x, ','))
            # self.df_result['3월'] = self.df_result['3월'].apply(lambda x: format(x, ','))
            # self.df_result['4월'] = self.df_result['4월'].apply(lambda x: format(x, ','))
            # self.df_result['5월'] = self.df_result['5월'].apply(lambda x: format(x, ','))
            # self.df_result['6월'] = self.df_result['6월'].apply(lambda x: format(x, ','))
            # self.df_result['7월'] = self.df_result['7월'].apply(lambda x: format(x, ','))
            # self.df_result['8월'] = self.df_result['8월'].apply(lambda x: format(x, ','))
            # self.df_result['9월'] = self.df_result['9월'].apply(lambda x: format(x, ','))
            # self.df_result['10월'] = self.df_result['10월'].apply(lambda x: format(x, ','))
            # self.df_result['11월'] = self.df_result['11월'].apply(lambda x: format(x, ','))
            # self.df_result['12월'] = self.df_result['12월'].apply(lambda x: format(x, ','))
            # self.df_result['지불소계'] = self.df_result['지불소계'].apply(lambda x: format(x, ','))
            # self.df_result['유보금액'] = self.df_result['유보금액'].apply(lambda x: format(x, ','))
            # self.df_result['미불금액'] = self.df_result['미불금액'].apply(lambda x: format(x, ','))
            self.df_result['비고(업체정보)'] = self.df_result[(self.df_result['비고(업체정보)'] == 0)] = '-'
            ws = self.wb.create_sheet(str(year)+'_지불집계표')
            for r in dataframe_to_rows(self.df_result, index=True, header=True):
                print(r)
                ws.append(r)
            ws.delete_rows(2)
            AutoFitColumnSize(ws)
        self.wb.save(self.FileOpen[0].split('.')[0]+"_집계.xlsx")

    def div_sheet(self):
        self.FileOpen = QFileDialog.getOpenFileName(self, 'Divide Sheet', './')
        self.label2.setText(self.FileOpen[0])
        self.wb = openpyxl.load_workbook(self.FileOpen[0])
        self.rowDf_divide = pd.read_excel(io=self.FileOpen[0])
        self.rowDf_divide['실결제금액'] = self.rowDf_divide['실결제금액'].fillna(0)
        self.rowDf_divide = self.rowDf_divide.astype({'상호': 'string', '대표자명': 'string',
                                        '주소': 'string', '합계금액': 'int64',
                                        '공급가액': 'int64', '세액': 'int64',
                                        '품목명': 'string', '품목비고': 'string',
                                        '실결제금액': 'int64', '결제수단': 'string',
                                        '차액': 'int64', '현장': 'string'
                                        })
        print(self.rowDf_divide)



if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = QtGUI()
    app.exec_()
