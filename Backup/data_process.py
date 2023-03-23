from genericpath import isfile
import os



def create_excel(JCtext):
    # if os.path.isfile(os.path.join())
    save_path = 'E:\\Job_Sonion\\AOI_Machine\\report'
    if os.path.isdir(save_path + str(JCtext)):
        pass
    else:
        print("Da tao file JCtext thanh cong")
        # os.mkdir(save_path + str(JCtext))
        with open(save_path + '/' + str(JCtext) + ".csv","a") as log:
            pass
    return str(JCtext) + ".csv",str(JCtext)
    