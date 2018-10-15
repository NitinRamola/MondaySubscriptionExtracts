# coding: utf-8

# In[ ]:

import numpy.core._methods
import numpy.lib.format
import tkinter
from tkinter import * 
import pyhdb
import logging
import time
import pandas as pd

logging.basicConfig(filename='C:\\Users\\Administrator\\Desktop\\Monday_Subscription_Reports\\ResellerKorea_New\\RuntimeLogs.log',level=logging.INFO,
                format='%(asctime)s:%(levelname)s:%(message)s')
def login():
    try:
        connection = pyhdb.connect(
        host="hp1node01.corp.adobe.com",
        port = 30015,
        user=uname.get(),
        password=passwd.get())

    except Exception as  err: 
        logging.info("Authentication Error Occured : ",err)
    
    else: 
        logging.info("Log-in Successfull")
        print("log-in Successfull")
        window.destroy()
        

    try:
        file = open(r'C:\Users\Administrator\Desktop\Monday_Subscription_Reports\ResellerKorea_New\KoreaQuery\Korea.sql',mode ='r')
        txt=file.read()
    except:
        logging.info("Error Occured while picking up query")
    else: 
        logging.info("Picked up the Query")
        cursor = connection.cursor()
        cursor.execute(txt)
        data=pd.DataFrame(cursor.fetchall())
        header=[]
        for i in cursor.description:
            header.append((i[0]))
    
    
        data.columns=header
        logging.info("The data size is : %s ",str(data.shape))
    
        timestr = time.strftime("%Y%m%d")
    
        outputpath = 'C:\\Users\\Administrator\\Desktop\\Monday_Subscription_Reports\\ResellerKorea_New\\OutputFiles-ADASH-4120 Reseller Report\\ADASH-4120 Reseller Report '+timestr+' KR.xlsx'
        writer = pd.ExcelWriter(outputpath)
        data.to_excel(writer,'Sheet1')
        writer.save()
        logging.info("Data successfully written")
        
        #closing window
        cwindow = Tk()
        cwindow.title("Monday Subscription Extract")
        cwindow.iconbitmap(r'Adobe.ico')
        cwindow.geometry('450x40')

        cl = Label(cwindow,text="The process has now completed")
        cl.pack()

#Opening window
window = Tk()
window.title("Monday Subscription Extract")
window.iconbitmap(r'Adobe.ico')
window.geometry('450x60')

uname = tkinter.StringVar()
passwd = tkinter.StringVar()


l1 = Label(window,text="LDAP ID : ")
l2 = Label(window,text="Password : ")

UserIdEntryBox=Entry(window,textvariable=uname)
UserPassEntryBox=Entry(window,show='*',textvariable=passwd)

loginbutton=Button(window,text=' Log-in ', command = login)

l1.pack(side=LEFT)
UserIdEntryBox.pack(side=LEFT)
loginbutton.pack(side=RIGHT)
UserPassEntryBox.pack(side=RIGHT)
l2.pack(side=RIGHT)
window.mainloop()



