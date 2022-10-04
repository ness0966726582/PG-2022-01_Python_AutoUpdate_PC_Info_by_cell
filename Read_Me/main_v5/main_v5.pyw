'''
程式開發者:NessHuang
開發日期:2022-09-30
版本:v5
使用平台:window10 / Google Sheet
開發語言:python3

打包:pyinstaller -F main_v5.pyw

主程式功能:
-檢查內外網狀態
 └---內網---->在NAS儲存IP.txt
 └---外網---->接續下方程序
-獲取NAS路徑文檔
-取得本機資訊{IP/AD/MAC}
-透過IP先在GOOGLE SHEET上比對
 └---檢查IP&MAC無差異---->在分頁【Get_ROW_LIST】 取得更新的CELL--->3
 └---檢查IP&MAC有差異---->在分頁【Get_ROW_LIST】 寫入IP/MAC/CELL--->回上一動
-在分頁【IP_MAC_INFO】更新update_list[get_TIME(),get_hostname(),get_IP(),getMAC()]於指定CELL

'''
import sys
import os
#第一步指派管理頁面相關資訊
code_Number= "PG-2022-01"#程式編號用於比對取得行數內容
GS_KEY = "\\\\10.231.199.10\\Department\\Form\\HO-ITD\\creds.json"#設定授權金鑰的存放於NAS路徑

#修改以下作為管理頁面的資訊
GS_Admin_URL="https://docs.google.com/spreadsheets/d/1w-9j0kvvvbDCaAeUS0cJXLr-W8CpbwCZtv4QuZWjG3c/edit#gid=506160378"
GS_Admin_ID="1w-9j0kvvvbDCaAeUS0cJXLr-W8CpbwCZtv4QuZWjG3c"
GS_Admin_PAGE="ALL_Auto_List"


#0.印出設備有的所有連線資訊+排序
def getUsing_TYPE_IP_MAC():
    global update_list , check_time, check_hostname, check_IP, check_MAC #取得硬體[time, hostname, adapter, IP,MAC]
    from datetime import datetime
    import pandas as pd
    import uuid, psutil, socket
    print("-----------------------------------------------------------------------------------------------------------------------------")
    network_Info=[]         #存取本機電腦的網路連線清單
    time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')          #時間 
    hostname = socket.gethostname()                              #主機名稱  
    
    IP = socket.gethostbyname(socket.gethostname())
    check_IP=IP#待檢查IP資訊

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
            check_time = update_list[0]
            check_hostname = update_list[1]
            check_IP = update_list[2]
            check_MAC = update_list[3]
            break

#開啟瀏覽器
def Open_URL(openURL):
    import webbrowser
    webbrowser.open(openURL)

#取得寫入的起始頁面--->此頁面若無IP以插入的方式新增資料    
def GS_Page_InsertRow():
    global insert_sheet_API
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    import pandas as pd 

    scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(GS_KEY, scope) #權限金鑰
    insert_sheet_API = gspread.authorize(creds) .open_by_key(GS_ID).worksheet(GS_insert_Page)# ID + PAGE

#連線PAGE1-作為硬體資訊的更新頁面--->此頁面已更新固定行的方式變動整行  
def GS_Page_UpdateInfo():
    global update_sheet_API,GS_update_Page
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    import pandas as pd
    GS_update_Page=GS_Page1
    scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(GS_KEY, scope) #權限金鑰
    update_sheet_API = gspread.authorize(creds) .open_by_key(GS_ID).worksheet(GS_update_Page)# ID + PAGE

#取得行數清單頁
def Page_GS_GetROW_result():
    global read_sheet_API,GS_insert_Page
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    import pandas as pd 
    GS_insert_Page=GS_Page2
    scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(GS_KEY, scope) #權限金鑰
    read_sheet_API= gspread.authorize(creds) .open_by_key(GS_ID).worksheet(GS_insert_Page)# ID + PAGE
    
    result = read_sheet_API.get_all_values()
    return result

#獲取頁面資訊與資料行數
def check_valuse():
    global row_end,df
    import pandas as pd
    df = pd.DataFrame(Page_GS_GetROW_result())  #取得回傳值read_test()-->內文陣列values[DEVICE_IP][USED_CELL]
    row_end = df.shape[0]           #總行數統計迴圈使用......印出目前行數
    print("目前清單建立比數:",row_end-2)#-2為去頭尾

#更新IP清單資訊
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

#檢查是否需要插入新的一筆資料
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
    
#1-1取得Admin管理頁面--->回傳資訊
def Page_GS_Admin_result():
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    import pandas as pd
    scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(GS_KEY, scope) #權限金鑰
    read_sheet_API= gspread.authorize(creds).open_by_key(GS_Admin_ID).worksheet(GS_Admin_PAGE)
    result = read_sheet_API.get_all_values()
    return result
#1-2取得管理頁面對應編號行的相關內容
def getGS_Admin_Info():
    global code_Name, GS_URL, GS_ID, GS_Page1, GS_Page2, GS_NAS_Path
    import pandas as pd
    import sys
    print("\n\n【將取得資訊宣告為全域變數】")
    print("-----------------------------------------------------------------------------------------------------------------------------")
    df = pd.DataFrame(Page_GS_Admin_result())  #取得回傳值read_test()-->內文陣列values[DEVICE_IP][USED_CELL]
    
    #取得頁面行數
    row_end = df.shape[0]           
    print("目前清單建立比數:",row_end)#-2為去頭尾
    
    #取得程式編號使用的 [url,id,page1,page2] 
    for x in range(0, int(row_end)):
        
        if df.loc[x][0]=="TRUE" and df.loc[x][1]==code_Number:
            print("檢查程式功能啟動:",df.loc[x][0])
            code_Name = df[2][x]
            GS_URL = df[3][x]
            GS_ID = df[4][x]
            GS_Page1 = df[5][x]
            GS_Page2 = df[6][x]
            GS_NAS_Path = df[7][x]
            print ("程式編號:" , code_Number,"程式命名:" , code_Name)
            print ("URL:" , GS_URL,"\nID:" , GS_ID)
            print ("PAGE1:" , GS_Page1,"\nPAGE2:" , GS_Page2)
            print ("NAS_URL:" ,GS_NAS_Path )
            break
        elif df.loc[x][0]=="FALSE" and df.loc[x][1]==code_Number:
            code_Name = df[2][x]
            print("code_Name:",code_Number,"code_Name:",code_Name,"is not used !")
            sys.exit(0)

#1-3指派文檔NAS路徑--->取用於getGS_Admin_Info()
def assign_NAS_Path():
    global NAS_Path_add_txt
    
    #轉換為windows可用路徑
    Nas_Path = GS_NAS_Path
    #print("替換成Windows 可視別的路徑:",Nas_Path)
    
    #判斷是否存在資料夾
    if os.access(Nas_Path, os.F_OK):
        print("路徑存在")
    else:
        #新增資料夾
        print("路徑不存在新增資料夾")
        os.mkdir(Nas_Path) #取至Admin管理頁面
    #調整路徑存放.txt--->提供給內往IP更新存放相關資訊
    NAS_Path_add_txt = Nas_Path + "\\" + check_IP + ".txt"
    print("Windows 可運行路徑提供:",NAS_Path_add_txt)

# 1-4 將資訊指向NAS寫入文字檔
def NAS_add_txt():
    print("------------------------------連線失敗嘗試->存在NAS路徑存儲.txt------------------------------")
    
    path = NAS_Path_add_txt #全域變數 指派文檔NAS路徑--->取用於getGS_Admin_Info()
    print(path)
    f = open(path, 'w')
    update_list
    #f.write("[")
    f.write(check_time)
    f.write(",")
    f.write(check_hostname)
    f.write(",")
    f.write(check_IP)
    f.write(",")
    f.write(check_MAC)
    #f.write("]")
    f.close()
    print("完成NAS寫入:",path)
    sys.exit(0)

#確認網路是否可以連接Google
def link_GS_or_not():
    import urllib3
    
    http = urllib3.PoolManager(timeout=3.0)
    r = http.request('GET', 'google.com', preload_content=False)
    code = r.status
    r.release_conn()
    if code == 200:
        #print("測試用NAS寫入文字檔",NAS_add_txt())# 測試用---->NAS寫入文字檔
        return True
    else:
        NAS_add_txt()# NAS寫入文字檔

#印出執行程式
print("\n--->執行程式編號:",code_Number)

#取得本機資訊與IP連線
getUsing_TYPE_IP_MAC()
print("\n--->取得硬體連線相關資訊:",update_list,check_IP)

#Admin管理頁程式開始運行
getGS_Admin_Info()#1-1~2取得管理頁面對應編號行的相關內容
assign_NAS_Path()#1-3指派文檔NAS路徑--->取用於getGS_Admin_Info()

#確認網路是否可以連接Google
print("------------------------------檢查網路連線------------------------------")
print(link_GS_or_not())

#此處Admin管理頁面結束
print("------------------------------以上為管理頁面的資訊取得------------------------------")


print("\n\n---------------------------主功能運行-------------------------------------------")
#取得行數清單頁資料
Page_GS_GetROW_result()
#連線PAGE1-作為硬體資訊的更新頁面--->此頁面已更新固定行的方式變動整行
GS_Page_UpdateInfo()
#連線PAGE2-取得寫入的起始頁面--->此頁面若無IP以插入的方式新增資料
GS_Page_InsertRow()#插入-IP & CELL 對照表使用

#更新IP清單資訊
update_IP_systemInfo_Row()
#檢查是否需要插入新的比數或插入新的
check_Insert_or_not()
#再做一次更新IP清單資訊
update_IP_systemInfo_Row()
#結束系統
sys.exit(0)


'''  
def getCSV():
    print("\n\n【抓取 data.csv 關於管理頁面的相關資訊】")
    print("-----------------------------------------------------------------------------------------------------------------------------")
    global GS_KEY,GS_URL,GS_ID,GSheet_listPage_name
    import csv
    import sys

    #nasPath_for_dataCSV = ["\\\\10.231.199.10\\Temp\\.AutoKey-ness\\data.csv"]
    #GS_KEY = ["\\\\10.231.199.10\\Temp\\.AutoKey-ness\\creds.json"]
    nasPath_for_dataCSV = ["\\\\10.231.199.10\\Department\\InformationTechnology\\03.專案組ProjectTeam\\Ness\\1.AD佈署程式\\PG-2022-01_Python_AutoUpdate_PC_Info_by_cell\\Read_Me\\main_v5\\Key-and-data\\data.csv"]
    GS_KEY = ["\\\\10.231.199.10\\Department\\InformationTechnology\\03.專案組ProjectTeam\\Ness\\1.AD佈署程式\\PG-2022-01_Python_AutoUpdate_PC_Info_by_cell\\Read_Me\\main_v5\\Key-and-data\\creds.json"]
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
    print(GS_KEY)
'''  