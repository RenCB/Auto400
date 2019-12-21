import windnd
import tkinter as tk
import ctypes
from EHLLAPI import Emulator
import openpyxl as opxl

#hapi._wait()
#print(hapi.set_cursor(753))
#print(hapi.get_field(453,20))
#print(hapi.get_cursor())
#print(hapi.search_str("WORD",0))
#print(hapi.lock_kb())
#print(hapi.find_field_length("T ",453))
#hapi.get_field(pos,feild length)
#print(hapi._wait())
# try:
#     hapi.search_str("WORLD",0)
# except:
#  print("CODE 24")
# Load DLL into memory.
hapi = Emulator()
#窗口
root = tk.Tk()
root.geometry("300x200")

#获取数据
dataList = {"Cust":[],"Deli":[],"MPN":[],"PO_num":[],"WH":[],"Deli_date":[],"Po_qty":[]}
endFlag = 0
def loadExcel_Data():
    global endFlag
    lable1['text'] = "开始加载EXCEL数据..."
    #Column (B=2 C=3 E=5 F=6 G=7 K=11 M=13)
    wb = opxl.load_workbook(lable['text'],data_only=True)
    ws = wb.active

    selectlist = [2,3,5,6,7,11,13]
    row_range=ws[2:ws.max_row]
    tempArr=[]
    for row in row_range:
        for cell in row:
            if(cell.column in selectlist):
                tempArr.append(cell.value)
        count = 0
        for key in dataList:
            dataList[key].append(tempArr[count])
            count = count+1
        tempArr=[]

    #处理Deli_date 数据格式
    formatDate_Arr = []
    for item in dataList["Deli_date"]:
        dt = item.date()
        dateYear = str(dt.year)[2:4] #只取年份后面两位
        dateMonth = str(dt.month)
        dateDay = str(dt.day)

        if(len(dateDay) <2 ):
            dateDay = "0" + dateDay
       
        if(len(dateMonth) <2):
            dateMonth = "0" + dateMonth
        #格式化日月年
        formatDate_Arr.append(dateDay+dateMonth+dateYear) 
    dataList["Deli_date"] = formatDate_Arr
    #print(dataList)
    formatDate_Arr = [] 
    lable1['text'] = "数据加载完毕！"
    endFlag = ws.max_row-1


def processFile():
    if(hapi.connect()==0):
        print("Connected!")
        
        cunt = 0
        while cunt < endFlag:

            #screen1
            hapi.copy_str_to_field(str(dataList["Cust"][cunt]),173)
            hapi.send_keys("@T")
            hapi.copy_str_to_field(str(dataList["Deli"][cunt]),253)
            hapi.send_keys("@T")
            hapi.copy_str_to_field(str(dataList["MPN"][cunt]),333)
            hapi.send_keys("@T")
            hapi.copy_str_to_field(str(dataList["PO_num"][cunt]),359)
            hapi.send_keys("@E@E")
            hapi._wait()

            #screen2
            
            hapi.copy_str_to_field(str(dataList["Deli_date"][cunt]),645)
            hapi.set_cursor(651)
            hapi.send_keys("@O@O@T")
            hapi.copy_str_to_field(str(dataList["Po_qty"][cunt]),654)
            hapi.send_keys("@T@E@a")
            hapi._wait()
            #print(dataList["MPN"][cunt])
            cunt = cunt +1

    if(hapi.disconnect()==0):
        print("Disconnected")

lable = tk.Label(root,bg="pink",width=55,height=5,text="文件拖放到这里")
lable1 = tk.Label(root,bg="orange",width=55,height=1,text="未加载文件")
button = tk.Button(root,text="开始录入",command=processFile)

# 对文件路径
def func(ls):
        filepath = str(ls)[3:-2]
        lable['text'] = filepath
        loadExcel_Data()
# windows 挂钩
windnd.hook_dropfiles(lable.winfo_id(),func)
lable.pack()
lable1.pack()
button.pack()







root.mainloop()