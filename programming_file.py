# Author: On Tu Quoc Dat - Control System Engineer
# Company : Sonion Viet Nam Co.,Ltd
# Version : 1.0
# Update: 04/10/2022
# Built = Python 3.10.7 

#Special command python -m PyQt5.uic.pyuic -x inteface.ui -o interface.py

from asyncio.windows_events import NULL
#from asyncore import write
import sys
from tempfile import TemporaryDirectory
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
        self.reject_list = []
        self.notsure_list = []
        self.report_path = None
        self.xlsx_name = None
    def show(self):
        self.main_win.show()

    def create_file(self):
        self.JCtext = self.uic.JCname_blank.text()
        if self.JCtext == '':
            self.uic.Alarm_Info.appendPlainText("You forgot to input into the JCname box. Input and press button again")
        else:
            self.report_path,self.xlsx_name = create_excel(self.JCtext)
            if self.xlsx_name != None:
                self.uic.Alarm_Info.appendPlainText("Create File Successfully.")

    def delete_database(self): 
        #Delete Database and Default Excel file  --------- Before press Start HMI -------------------
        database_pathname= 'C:\\Users\\U-1TL8FV2\\Documents\\Cognex Designer\\Projects\\Add_Vidi_DAQ\\Deploy2\\Data\\DAQdatabase.db3'
        open(database_pathname,'w').close()
        #self.reset_function()

    def copy_content_excelfile(self):
        #Copy to new file
        #Point to default directory of excel Cognex
        start_file = r'C:\\Users\\U-1TL8FV2\\Documents\\Cognex Designer\\Projects\\Add_Vidi_DAQ\\Inner.csv'
        if self.report_path != None:
            end_file = r'C:\\Users\\U-1TL8FV2\\Desktop\\Report Excel\\temp\\' + self.report_path
            # print(end_file)
            if os.stat(start_file).st_size != 0:
                shutil.copyfile(start_file,end_file)

                
                #Delete old file --------------- BONUS OPTION ---------------------
                open(start_file,'w').close()

                #Convert csv to xlsx
                read_file = pd.read_csv(end_file)
                read_file.to_excel(r'C:\\Users\\U-1TL8FV2\\Desktop\\Report Excel\\' +self.xlsx_name +'.xlsx', index = None, header=True)
                
                #Xu ly du lieu tho thanh du lieu tinh
                self.separate_and_empty_same_row()
            else:
                self.uic.Alarm_Info.appendPlainText("Empty input file. CHECK BUTTON EXCEL OUTPUT IN MAIN RUNTIME.Maybe you forgot press it")
        else:
            self.uic.Alarm_Info.appendPlainText("Forgot to press Create File")

    def separate_and_empty_same_row(self):
        input_file = 'C:\\Users\\U-1TL8FV2\\Desktop\\Report Excel\\' +self.xlsx_name +'.xlsx'
        # input_file = 'E:\\Job_Sonion\\AOI_Machine\\report\\daq.xlsx'
        #Inner column 1 2 3 4
        excel_data_df = pd.read_excel(input_file,sheet_name='Sheet1',header=None)
        #print(excel_data_df.iat[1,2]) # Hàng 1 cột 2
        #print(excel_data_df.iloc[1:2])  #Print cả hàng hàng 1 -> hàng 2
        pre_value = 0
        pre_value_outer = 0
        # print(len(excel_data_df))
        
        #Set name columns
        excel_data_df.columns = ['NumberIn','xNumxIn','ResultIn','ScoreIn','NumberOut','xNumxOut','ResultOut','ScoreOut']
        excel_data_df_copy = excel_data_df.copy()

        #Inner column 1 2 3 4
        for j in range (1,len(excel_data_df)):
            present_value = excel_data_df.iat[j,0]
            if present_value == pre_value:
                #print("Trung,        ",j)
                excel_data_df.iat[j,0] = ''
                excel_data_df.iat[j,1] = ''
                excel_data_df.iat[j,2] = ''
                excel_data_df.iat[j-1,3] = excel_data_df.iat[j,3]
                excel_data_df.iat[j,3] = ''
            elif present_value == pre_value + 1:
                pre_value = present_value
            #print(excel_data_df.iat[j,0],excel_data_df.iat[j,1],excel_data_df.iat[j,2],excel_data_df.iat[j,3])
            elif present_value == pre_value + 2:
                pre_value = present_value
            #else:
                #print("Error Inner Number Export.Recheck your data and Cognex???")
                #excel_data_df.iat[j,0] = ''
                #excel_data_df.iat[j,1] = ''
                #excel_data_df.iat[j,2] = ''
                #excel_data_df.iat[j,3] = ''
        
        # print(excel_data_df.iloc[:])

        # Save final inner data
        # excel_data_df.to_excel('dattest.xlsx')
        excel_data_df[['NumberIn','xNumxIn','ResultIn','ScoreIn']].to_excel('Inner_Data.xlsx')

        #Outer column 5 6 7 8
        for i in range (1,len(excel_data_df_copy)):
            present_value_outer = excel_data_df_copy.iat[i,4]
            if present_value_outer == pre_value_outer:
                excel_data_df_copy.iat[i,4] = ''
                excel_data_df_copy.iat[i,5] = ''
                excel_data_df_copy.iat[i,6] = ''
                excel_data_df_copy.iat[i-1,7] = excel_data_df_copy.iat[i,7]
                excel_data_df_copy.iat[i,7] = ''
            elif present_value_outer == pre_value_outer + 1:
                pre_value_outer = present_value_outer
            # print(excel_data_df_copy.iat[i,4],excel_data_df_copy.iat[i,5],excel_data_df_copy.iat[i,6],excel_data_df_copy.iat[i,7])
            elif present_value_outer == pre_value_outer + 2:
                pre_value_outer = present_value_outer
            #else:
                #print("Error Outer Number Export.Recheck your data and Cognex???")
             #   excel_data_df.iat[i,4] = ''
              #  excel_data_df.iat[i,5] = ''
               # excel_data_df.iat[i,6] = ''
                #excel_data_df.iat[i,7] = ''               
            

        #Save final outer data
        excel_data_df_copy[['NumberOut','xNumxOut','ResultOut','ScoreOut']].to_excel('Outer_Data.xlsx')

        self.delete_empty_row()

    def filter_skip_samples(self):
        read_input_file = 'C:\\Users\\DAQ\\Downloads\\10_10_2022_Pythonapp\\AOI_Machine_pythonapp\\draft\\Inner_Data.xlsx'
        transfer_to_pd = pd.read_excel(read_input_file,sheet_name='Sheet1',header=None)

        transfer_to_pd.columns = ['temp','NumberIn','xNumxIn','ResultIn','ScoreIn'] 
        
        temp_inner = 0
        for k in range(2,300):#range(1,len(transfer_to_pd)):
            # print("Number Inner:    ",type(transfer_to_pd.iat[k,1]),transfer_to_pd.iat[k,1])
            # print("Num: ",transfer_to_pd.iat[1,1])
            if pd.isnull(transfer_to_pd.iat[k,1]):
                # print("DAQ empty")
                pass
            else:
                # print("K = ",k)
                if transfer_to_pd.iat[k,1] == temp_inner + 1 :
                    temp_inner = transfer_to_pd.iat[k,1]
                else:
                    row_need_to_add = transfer_to_pd.iat[k,1] - temp_inner - 1
                    row_need_to_add == int(row_need_to_add)
                    print("Row need: ",row_need_to_add, k)
                    df_new = self.add_blank_rows(transfer_to_pd,row_need_to_add)
                    print(df_new)
                    
                # if transfer_to_pd.iat[k,1] == temp_inner + 1 or transfer_to_pd.iat[k,1] == temp_inner:
                #     temp_inner = transfer_to_pd[k,1]
                # elif transfer_to_pd.iat[k,1] == temp_inner:
                #     temp_inner = transfer_to_pd[k,1]
                # else: 
                #     row_need_to_add = transfer_to_pd.iat[k,1] - temp_inner
                #     print("Row need to add =   ",row_need_to_add)
                    # print("K:       ",k)
                    # df_new = self.add_blank_rows(transfer_to_pd,row_need_to_add)
        
        # df_new[['NumberIn','xNumxIn','ResultIn','ScoreIn']].to_excel('Dat_test.xlsx')
                    
    def add_blank_rows(self,df,no_rows):
        df_new = pd.DataFrame(columns=df.columns)
        
        # for idx in range(len(df)):
        #     df_new = df_new.append(df.iloc[idx])
        #     for _ in range(no_rows):
        #         df_new = df_new.append(pd.Series(),ignore_index = True)
        return df_new


    def delete_empty_row(self):
        #Code doan nay ko hay gi het. thoi ke me no, chay duoc la ok roi
        inner_file = 'Inner_Data.xlsx'
        outer_file = 'Outer_Data.xlsx'
        inner_df = pd.read_excel(inner_file,sheet_name='Sheet1',header=None)
        outer_df = pd.read_excel(outer_file,sheet_name='Sheet1',header=None)
        inner_df.columns = ['temp','NumberIn','xNumxIn','ResultIn','ScoreIn']
        outer_df.columns = ['temp','NumberOut','xNumxOut','ResultOut','ScoreOut']
        nan_value = float("NaN")
        inner_df.replace("",nan_value,inplace=True)
        inner_df.dropna(subset=["NumberIn"],inplace=True)   

        outer_df.replace("",nan_value,inplace=True)
        outer_df.dropna(subset=["NumberOut"],inplace=True) 

        writer = pd.ExcelWriter('C:\\Users\\U-1TL8FV2\\Desktop\\Report Excel\\'+self.xlsx_name+'.xlsx')
        inner_df[['NumberIn','xNumxIn','ResultIn','ScoreIn']].to_excel(writer,sheet_name='Sheet1',startrow=0,startcol=0)
        outer_df[['NumberOut','xNumxOut','ResultOut','ScoreOut']].to_excel(writer,sheet_name='Sheet1',startrow=0,startcol=5)
        # final_df.to_excel('3.xlsx')
        writer.save()
        self.uic.Alarm_Info.appendPlainText("Saved Final Result Successfully.")
    


    def reset_function(self):
        self.uic.Alarm_Info.clear()
        self.uic.Alarm_Info.setPlainText("ALARM AREA !!! CALL DEVELOPER IF ANY ERROR HAPPENED")

    def notice_function(self,notice):
        self.uic.Alarm_Info.appendPlainText(notice)
    
    def final_group_decision(self):
        pass

    def failure_mode_display(self):
        pass

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_win = MainWindow()
    main_win.show()
    
    #Trigger Event
    main_win.uic.createfile_button.clicked.connect(main_win.create_file)

    main_win.uic.deletedatabase_button.clicked.connect(main_win.delete_database)

    main_win.uic.copyfile_button.clicked.connect(main_win.copy_content_excelfile)

    main_win.uic.reset_button.clicked.connect(main_win.reset_function)

    main_win.filter_skip_samples()
    # main_win.separate_and_empty_same_row()
    # main_win.delete_empty_row()
    sys.exit(app.exec())
