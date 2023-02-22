import os
import xlrd
import time
from netmiko import ConnectHandler
import threading
#path
Base_Path=os.path.dirname(__file__)
#configure path
Conf_Path= Base_Path + '\log'

Log_dir = os.path.join(Conf_Path,time.strftime("%Y-%m-%d_%H_%M",time.localtime()))
os.mkdir(Log_dir)
#get info from excle
#open excle
Device_Shell = xlrd.open_workbook(Base_Path+'\Devier.xlsx')

#open device info
Device_info  =  Device_Shell.sheet_by_name("Device")

#open command sheel
Command_info = Device_Shell.sheet_by_name("Command")


# get device info to list


#get device infomation
def dev_info():
    All_info = []
    for i in range(1,Device_info.nrows):

        info  = { 
        "device_type": "cisco_ios" if Device_info.row_values(i)[3] == "ssh" else "cisco_ios_telnet",
        "host": Device_info.row_values(i)[2],
        "username": Device_info.row_values(i)[5],
        "password": Device_info.row_values(i)[6],
        'port' : 22 if Device_info.row_values(i)[3] == 'ssh' else 23,
        "secret": None if Device_info.row_values(i)[7] == '' else Device_info.row_values(i)[7],
        }

        hostname= Device_info.row_values(i)[0]
        manufactor= Device_info.row_values(i)[1]

        All_info.append([hostname,manufactor,info])
    return All_info


#login device,get config
def get_data(hostname,login_info,manufactor):
    try:
        dev_file=os.path.join(Log_dir,hostname)
        os.mkdir(dev_file)
        connect=ConnectHandler(**login_info)
        if manufactor == "cisco":
            connect.enable()
            for num in range(1,Command_info.nrows):             
                output = connect.send_command(Command_info.row_values(num)[0])
                os.chdir(dev_file)
                with open(Command_info.row_values(num)[0] + '.conf', 'wt') as f:
                    f.write(output)
                    f.close()
                    print(hostname+" -"+ Command_info.row_values(num)[0] +" backup successful!")
        else:
            for num in range(1,Command_info.nrows):
                output = connect.send_command(Command_info.row_values(num)[1])
                os.chdir(dev_file)
                with open(Command_info.row_values(num)[1]+ '.log', 'wt') as f:
                    f.write(output)
                    f.close()
                    print(hostname+" -"+Command_info.row_values(num)[1] +" backup successful!")
        connect.disconnect()
    except Exception as e:
        print(hostname +" -"+ " bankup faild!")



if __name__ == '__main__':
    for info in dev_info():
        th = threading.Thread(target=get_data,args=(info[0],info[2],info[1]))
        th.start()

    for i in range(8):
        th.join()
    print("stoped")


