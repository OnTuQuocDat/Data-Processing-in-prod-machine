# Author: On Tu Quoc Dat - Control System Engineer
# Company : Sonion Viet Nam Co.,Ltd
# Version : 1.0
# Update: 29/09/2022
# Built = Python 3.10.7 

from asyncio.windows_events import NULL
from asyncore import write
import sys
from tracemalloc import start
from PyQt5.QtWidgets import QApplication,QMainWindow
from data_process import create_excel
from interface import Ui_ExcelSupportApp
import os
import pandas as pd
import shutil
from openpyxl.cell import Cell

class MainWindow:
    def __init__(self):
        self.main_win = QMainWindow()
        self.uic = Ui_ExcelSupportApp()
        self.uic.setupUi(self.main_win)
    def show(self):
        self.main_win.show()
    # def show_JCblank(self):
    #     self.JCtext = self.uic.JCname_blank.text()
    #     print(self.JCtext)
    def create_file(self):
        self.JCtext = self.uic.JCname_blank.text()
        self.report_path,self.xlsx_name = create_excel(self.JCtext)
    def delete_database(self):
        #Delete Database and Default Excel file  --------- Before press Start HMI -------------------
        database_pathname= 'E:\\Job_Sonion\\AOI_Machine\\draft\\test.db3'
        open(database_pathname,'w').close()
    def copy_content_excelfile(self):
        #Copy to new file
        start_file = r'E:\\Job_Sonion\\AOI_Machine\\report\\DAQresult.csv'

        end_file = r'E:\\Job_Sonion\\AOI_Machine\\report\\' + self.report_path
        print(end_file)
        shutil.copyfile(start_file,end_file)

        
        #Delete old file --------------- BONUS OPTION ---------------------
        #open(start_file,'w').close()

        #Convert csv to xlsx
        read_file = pd.read_csv(end_file)
        read_file.to_excel(r'E:\\Job_Sonion\\AOI_Machine\\report\\' +self.xlsx_name +'.xlsx', index = None, header=True)
        self.separate_and_empty_same_row

    def separate_and_empty_same_row(self):
        #input_file = 'E:\\Job_Sonion\\AOI_Machine\\report\\' +self.xlsx_name +'.xlsx'
        input_file = 'E:\\Job_Sonion\\AOI_Machine\\report\\daq.xlsx'
        #Inner column 1 2 3 4
        excel_data_df = pd.read_excel(input_file,sheet_name='Sheet1',header=None)
        #print(excel_data_df.iat[1,2]) # Hàng 1 cột 2
        #print(excel_data_df.iloc[1:2])  #Print cả hàng hàng 1 -> hàng 2
        pre_value = 0
        pre_value_outer = 0
        #print(len(excel_data_df))
        
        #Set name columns
        excel_data_df.columns = ['a','b','c','d','e','f','g','h']
        excel_data_df_copy = excel_data_df.copy()


        for j in range (1,len(excel_data_df)):
            present_value = excel_data_df.iat[j,0]
            if present_value == pre_value:
                #print("Trung,        ",j)
                excel_data_df.iat[j,0] = ''
                excel_data_df.iat[j,1] = ''
                excel_data_df.iat[j,2] = ''
                excel_data_df.iat[j,3] = ''
            elif present_value == pre_value + 1:
                pre_value = present_value
            #print(excel_data_df.iat[j,0],excel_data_df.iat[j,1],excel_data_df.iat[j,2],excel_data_df.iat[j,3]) 
        
        # print(excel_data_df.iloc[:])
        # Delete empty rows in xlsx file
        # nan_value = float("NaN")
        # excel_data_df.replace("",nan_value,inplace=True)
        # excel_data_df.dropna(subset=["a"],inplace=True)


        # Save final inner data
        # excel_data_df.to_excel('dattest.xlsx')
        excel_data_df[['a','b','c','d']].to_excel('Inner_Data.xlsx')

        #Outer column 5 6 7 8
        for i in range (1,len(excel_data_df_copy)):
            present_value_outer = excel_data_df_copy.iat[i,4]
            if present_value_outer == pre_value_outer:
                excel_data_df_copy.iat[i,4] = ''
                excel_data_df_copy.iat[i,5] = ''
                excel_data_df_copy.iat[i,6] = ''
                excel_data_df_copy.iat[i,7] = ''
            elif present_value_outer == pre_value_outer + 1:
                pre_value_outer = present_value_outer
            # print(excel_data_df_copy.iat[i,4],excel_data_df_copy.iat[i,5],excel_data_df_copy.iat[i,6],excel_data_df_copy.iat[i,7]) 


        # Delete empty rows in xlsx file
        # nan_value = float("NaN")
        # excel_data_df.replace("",nan_value,inplace=True)
        # excel_data_df.dropna(subset=["e"],inplace=True)

        #Save final outer data
        excel_data_df_copy[['e','f','g','h']].to_excel('Outer_Data.xlsx')

    def delete_empty_row(self):
        #Code doan nay ko hay gi het. thoi ke me no, chay duoc la ok roi
        inner_file = 'Inner_Data.xlsx'
        outer_file = 'Outer_Data.xlsx'
        inner_df = pd.read_excel(inner_file,sheet_name='Sheet1',header=None)
        outer_df = pd.read_excel(outer_file,sheet_name='Sheet1',header=None)
        inner_df.columns = ['temp','a','b','c','d']
        outer_df.columns = ['temp','e','f','g','h']
        nan_value = float("NaN")
        inner_df.replace("",nan_value,inplace=True)
        inner_df.dropna(subset=["a"],inplace=True)   

        outer_df.replace("",nan_value,inplace=True)
        outer_df.dropna(subset=["e"],inplace=True) 

        writer = pd.ExcelWriter('Summary.xlsx')
        inner_df[['a','b','c','d']].to_excel(writer,sheet_name='Sheet1',startrow=0,startcol=0)
        outer_df[['e','f','g','h']].to_excel(writer,sheet_name='Sheet1',startrow=5,startcol=5)
        # final_df.to_excel('3.xlsx')
        writer.save()



if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_win = MainWindow()
    main_win.show()
    
    #Trigger Event
    main_win.uic.createfile_button.clicked.connect(main_win.create_file)

    main_win.uic.deletedatabase_button.clicked.connect(main_win.delete_database)

    main_win.uic.copyfile_button.clicked.connect(main_win.copy_content_excelfile)
    main_win.separate_and_empty_same_row()
    main_win.delete_empty_row()
    sys.exit(app.exec())