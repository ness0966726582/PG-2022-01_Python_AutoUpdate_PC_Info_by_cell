'''
程式開發者:NessHuang
開發日期:2022-09-30
版本:v4
使用平台:window10 / Google Sheet
開發語言:python3

打包:pyinstaller -F main_v4.pyw

主程式功能:
0獲取NAS路徑文檔
1取得本機資訊{IP/AD/MAC}
2透過IP先在GOOGLE SHEET上比對
 └---檢查IP&MAC無差異---->在分頁【Get_ROW_LIST】 取得更新的CELL--->3
 └---檢查IP&MAC有差異---->在分頁【Get_ROW_LIST】 寫入IP/MAC/CELL--->回上一動
3.在分頁【IP_MAC_INFO】更新update_list[get_TIME(),get_hostname(),get_IP(),getMAC()]於指定CELL

'''

#印出設備有的所有連線資訊+排序
def getUsing_TYPE_IP_MAC():
    print("\n\n【呈現連線種類清單】")
    print("-----------------------------------------------------------------------------------------------------------------------------")
    global update_list #取得硬體[time, hostname, adapter, IP,MAC]
    from datetime import datetime
    import pandas as pd
    import uuid, psutil, socket
    
    network_Info=[]         #存取本機電腦的網路連線清單
    time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')          #時間 
    hostname = socket.gethostname()                              #主機名稱  
    
    IP = socket.gethostbyname(socket.gethostname())
    print("目前連線IP:",IP)
    
    #https://www.twblogs.net/a/5d5e4351bd9eee541c324130
    r""" 打印多網卡 mac 和 ip 信息 """
    dic = psutil.net_if_addrs()    
    for adapter in dic:
        snicList = dic[adapter]
        mac = '無 mac 地址'
        ipv4 = '無 ipv4 地址'
        ipv6 = '無 ipv6 地址'
        for snic in snicList:
            if snic.family.name in {'AF_LINK', 'AF_PACKET'}:
                mac = snic.address
            elif snic.family.name == 'AF_INET':
                ipv4 = snic.address
            elif snic.family.name == 'AF_INET6':
                ipv6 = snic.address
        list=[['%s, %s, %s' % (adapter, mac, ipv4)]]        
        #print(list[0])
        network_Info.append( list[0] )
        
        
    df = pd.DataFrame(network_Info)  #取得回傳值
    row_end = df.shape[0]           #總行數統計迴圈使用......印出目前行數
    print("設備連線方式筆數:",row_end)#-2為去頭尾
    
    print("取得設備清單IP+MAC:",str(network_Info[0][0]).split(','))
    print("取得設備清單IP+MAC:",str(network_Info[1][0]).split(','))
    print("取得設備清單IP+MAC:",str(network_Info[2][0]).split(','))
    print("取得設備清單IP+MAC:",str(network_Info[3][0]).split(','))
    
    #list=str(network_Info[0][0]).split(',')
    #print("取得IP:",list[2])
    
    #下方pandas進行排序與判斷
    print("---------------for進行排序與判斷--------------------")
    
    for x in range (0, int(row_end-1)):          #從0開始 逐行進行split
        list=str(network_Info[x][0]).split(',')
        #print("目前連線:",IP)      
        #print("取得ADAPTER:",list[0])
        #print("取得MAC:",list[1])
        #list[2].strip(" ")                   #除去空白字元
        #print("取得IP:",list[2])
        
        if str(list[2].strip(" "))==str(IP):      #除去空白字元比對IP
            print("找到",IP,"對應MCA欄位=",list[1])
            adapter=list[0]
            MAC=list[1]
            update_list = [time, hostname, IP,MAC]
            break

def getCSV():
    print("\n\n【抓取 data.csv 關於管理頁面的相關資訊】")
    print("-----------------------------------------------------------------------------------------------------------------------------")
    global GS_KEY,GS_URL,GS_ID,GSheet_listPage_name
    import csv
    import sys

    nasPath_for_dataCSV = ["\\\\10.231.199.10\\Temp\\.AutoKey-ness\\data.csv"]
    GS_KEY = ["\\\\10.231.199.10\\Temp\\.AutoKey-ness\\creds.json"]
    #with open('./data.csv', newline='',encoding='utf-8') as csvfile:
    with open(nasPath_for_dataCSV[0], newline='',encoding='utf-8') as csvfile:     
        rows = csv.reader(csvfile,delimiter=",")# 讀取 CSV 檔案內容，分隔符號是 Tab "\t"
        csvList = []# 設定一個空陣列      
        for row in rows:# # 以迴圈輸出每一列資料加到 csvlist 陣列裡
            csvList.append(row) 

    #自動化程式的管理列表------->提供程式連入取得程式應用的相關資訊
    print("程式管理頁面URL:",csvList[1][1])
    print("程式管理頁面ID:",csvList[1][2])
    print("程式管理頁面PAGE:",csvList[1][3])
    GS_URL = csvList[1][1]
    GS_ID = csvList[1][2]
    GSheet_listPage_name = csvList[1][3]

'''
#從Google管理頁"比對程式編號"獲取整行資訊
'''
def all_Auto_GSeet_List():
    global read_sheet_API
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    import pandas as pd 
    #獲取授權與連結#
    scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(GS_KEY[0], scope) #權限金鑰
    read_sheet_API= gspread.authorize(creds) .open_by_key(GS_ID).worksheet(GSheet_listPage_name)
    
    result = read_sheet_API.get_all_values()
    return result
'''
抓取主程式使用的必要資訊
啟動功能 / 程式編號 / 程式檔名 / GooglesheetURL / GooglesheetID / PageName1 / PageName2
'''
def getConnect_Info():
    print("\n\n【保留主程式初始變數】")
    print("-----------------------------------------------------------------------------------------------------------------------------")
    global code_Name, GS_URL, GS_ID, GS_update_Page, GS_insert_Page
    import pandas as pd
    df = pd.DataFrame(all_Auto_GSeet_List())  #取得回傳值read_test()-->內文陣列values[DEVICE_IP][USED_CELL]
    
    #取得頁面行數
    row_end = df.shape[0]           
    print("目前清單建立比數:",row_end)#-2為去頭尾
    
    #取得程式編號使用的 [url,id,page1,page2] 
    for x in range(0, int(row_end)):
        
        if df.loc[x][0]=="TRUE" and df.loc[x][1]==program_Number:
            print("檢查程式功能啟動:",df.loc[x][0])
            code_Name = df[2][x]
            GS_URL = df[3][x]
            GS_ID = df[4][x]
            GS_update_Page = df[5][x]
            GS_insert_Page = df[6][x]
            
            print ("程式編號:" , program_Number)
            print ("程式命名:" , code_Name)
            
            print ("獲取的URL:" , GS_URL)
            print ("獲取的ID:" , GS_ID)
            print ("獲取的PAGE1:" , GS_update_Page)
            print ("獲取的PAGE2:" , GS_insert_Page)
            
            break
        elif df.loc[x][0]=="FALSE" and df.loc[x][1]==program_Number:
            code_Name = df[2][x]
            print("code_Name:",program_Number,"code_Name:",code_Name,"is not used !")
            sys.exit(0)






'''
Open_URL(GS_URL)
#自訂開啟瀏覽器

insertData_googleSeet_API_Key(GS_ID,GS_insert_Page)
#自訂GOOGLE SHEET的頁面連線---->插入IP_CEL的比對資訊為目的

updateData_googleSeet_API_Key(GS_ID,GS_update_Page)
#自訂GOOGLE SHEET的頁面連線---->更新TIME/AD/IP/MAC為目的
'''
def Open_URL():
    import webbrowser
    webbrowser.open(GS_URL)

    
def insertData_googleSeet_API_Key():
    global insert_sheet_API
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    import pandas as pd 

    scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(GS_KEY[0], scope) #權限金鑰
    insert_sheet_API = gspread.authorize(creds) .open_by_key(GS_ID).worksheet(GS_insert_Page)# ID + PAGE

def updateData_googleSeet_API_Key():
    global update_sheet_API
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    import pandas as pd

    scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(GS_KEY[0], scope) #權限金鑰
    update_sheet_API = gspread.authorize(creds) .open_by_key(GS_ID).worksheet(GS_update_Page)# ID + PAGE

def read_CheckList_Page():
    global read_sheet_API
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    import pandas as pd 

    scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(GS_KEY[0], scope) #權限金鑰
    read_sheet_API= gspread.authorize(creds) .open_by_key(GS_ID).worksheet(GS_insert_Page)# ID + PAGE
    
    result = read_sheet_API.get_all_values()
    return result

'''
check_valuse()檢查頁面筆數
update_IP_systemInfo_Row()更新
check_Insert_or_not()檢查是否需要插入新的一筆
'''

def check_valuse():
    global row_end,df
    import pandas as pd
    df = pd.DataFrame(read_CheckList_Page())  #取得回傳值read_test()-->內文陣列values[DEVICE_IP][USED_CELL]
    row_end = df.shape[0]           #總行數統計迴圈使用......印出目前行數
    print("目前清單建立比數:",row_end-2)#-2為去頭尾

def update_IP_systemInfo_Row():
    global do_insert_or_not
    check_valuse()

    for x in range(1, int(row_end)): 
        if df.loc[x][0]==check_IP and df.loc[x][1]==check_MAC:
            print("找到",check_IP,"對應欄位=",df.loc[x][2])
            select_Cell=df.loc[x][2]
            print("本機資訊:",update_list)
            update_Row = [ update_list ]
            print("更新資訊於IP_sysInfo:",update_list)
            update_sheet_API.update(select_Cell, update_Row) #更新Array
            do_insert_or_not="N"
            break
        else :
            print(check_IP,df.loc[x][0],"無IP對應欄位")
            do_insert_or_not="Y"

def check_Insert_or_not():
    global do_insert_or_not
    #迴圈結束無建立IP與CELL對照表---->插入IP與新的CELL
    check_valuse()
    next_cell="A"+str(row_end)

    if do_insert_or_not=="Y":
        insert_Row=[check_IP,check_MAC,next_cell]
        select_Row=2
        insert_sheet_API.insert_row(insert_Row, select_Row)#插入List
        print("寫入資訊內容:",insert_Row,"插入行數:",select_Row)
        print("完成插入新IP與CELL",)
    do_insert_or_not="N"
    print (do_insert_or_not)

'''
程式運行
'''

import sys
program_Number= "PG-2022-01"#程式編號
getUsing_TYPE_IP_MAC()
print("\n--->執行程式編號:",program_Number)
print("\n--->取得硬體連線相關資訊:",update_list)
getCSV()#取得CSV文檔
getConnect_Info()#取得所需主程式資訊
print("------------------------------以上為管理頁面的資訊取得------------------------------")
#Open_URL()
read_CheckList_Page()
updateData_googleSeet_API_Key()#更新-主要查詢IP設備資訊使用
insertData_googleSeet_API_Key()#插入-IP & CELL 對照表使用

print("\n\n【主功能運行】")
print("-----------------------------------------------------------------------------------------------------------------------------")

check_IP=update_list[2]#待檢查IP資訊
check_MAC=update_list[3]#待檢查MAC資訊
print("本機IP:",update_list[2],"本機MAC:",update_list[3])

update_IP_systemInfo_Row()
check_Insert_or_not()

update_IP_systemInfo_Row()
sys.exit(0)