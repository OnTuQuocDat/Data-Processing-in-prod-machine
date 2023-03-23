from genericpath import isfile
import os



def create_excel(JCtext):
    # if os.path.isfile(os.path.join())
    save_path = 'C:\\Users\\U-1TL8FV2\\Desktop\\Report Excel\\temp'
    # save_path = 'C:\\Users\\DAQ\\Downloads\\10_10_2022_Pythonapp\\AOI_Machine_pythonapp'
    if os.path.isdir(save_path + str(JCtext)):
        pass
    else:
        print("Da tao file JCtext thanh cong")
        # os.mkdir(save_path + str(JCtext))
        with open(save_path + '/' + str(JCtext) + ".csv","a") as log:
            pass
    return str(JCtext) + ".csv",str(JCtext)
    
