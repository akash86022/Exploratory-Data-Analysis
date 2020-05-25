import pandas as pd
import numpy as np
import time
import win32com.client as win32
import os

def send_mail():
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'akash.jain@cerner.com'
    mail.Subject = '7 Day Communication '
    mail.Body = ''
    attachment  = "C:\\Users\\aj055556\\Desktop\\7_Day_Comm.csv"
    mail.Attachments.Add(attachment)
    mail.Send()

print("Reading File")
temp=pd.read_csv("C:\\Users\\aj055556\\Downloads\\Data Dash - Queue Sizes.csv")
temp=temp.iloc[:,[0,2,4,8,40,41]]
temp1=temp[temp['Last Outbound Comm']>7].sort_values('Last Outbound Comm',ascending=False)
temp1.to_csv('7_Day_Comm.csv',index=False)
#print("Sending Email")
#send_mail()
os.remove("C:\\Users\\aj055556\\Downloads\\Data Dash - Queue Sizes.csv")
print("Operation Completed !! ")
time.sleep(2)
