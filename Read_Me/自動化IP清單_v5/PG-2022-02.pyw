'''
程式開發者:NessHuang
開發日期:2022-10-05
版本:v1
使用平台:window10 / Google Sheet
開發語言:python3

打包:pyinstaller -F txt合併上傳_v1.pyw

PG-2022-02
副程式功能:
-從GS_Admin管理頁面上獲取此程式的相關資訊
-獲取NAS路徑文檔合併轉CSV
-判斷文檔是否執行上傳
 └---NAS路徑無檔案---->中斷程序
 └---NAS路徑有檔案---->接續下方程序
-GS頁面比對IP/MAC 取得指定行數
└-->若有此設備-->更新於指定行數
└-->若無此設備-->插入新的行數-->回前一動



'''

import sys

#第一步指派管理頁面相關資訊
code_Number= "PG-2022-02"#程式編號用於比對取得行數內容
GS_KEY = "\\\\10.231.199.10\\Department\\Form\\HO-ITD\\creds.json"#設定授權金鑰的存放於NAS路徑

#修改以下作為管理頁面的資訊
GS_Admin_URL="https://docs.google.com/spreadsheets/d/1w-9j0kvvvbDCaAeUS0cJXLr-W8CpbwCZtv4QuZWjG3c/edit#gid=506160378"
GS_Admin_ID="1w-9j0kvvvbDCaAeUS0cJXLr-W8CpbwCZtv4QuZWjG3c"
GS_Admin_PAGE="ALL_Auto_List"

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

#2-1取得行數清單頁
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

#2-2獲取頁面資訊與資料行數
def check_valuse():
    global row_end,df
    import pandas as pd
    df = pd.DataFrame(Page_GS_GetROW_result())  #取得回傳值read_test()-->內文陣列values[DEVICE_IP][USED_CELL]
    row_end = df.shape[0]           #總行數統計迴圈使用......印出目前行數
    print("目前清單建立比數:",row_end-2)#-2為去頭尾

#2-3更新IP清單資訊
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
    




'''
到指定路徑old_Path下
將資料夾內文字檔合併
轉存到指定路徑
'''
#獲取目前時間
def get_Today():
    import time
    global today
    
    localtime = time.localtime()
    today = time.strftime("%Y-%m-%d", localtime)  
    return today

#檢查路徑內是否有檔案
def check_Path_hadFile():
    import sys
    import os
    import os.path
    old_Path=GS_NAS_Path
    
    if os.access(old_Path, os.F_OK):
        print("路徑存在")
    else:
        #結束系統
        sys.exit(0)  
    
#舊路徑上的所有TXT合併-->日期.txt
def merge_oldPath_txt():
    import os
    import os.path
    global txt_new_Path,new_Path #提供當天整合後的txt的路徑
    
    old_Path=GS_NAS_Path #取GS的NAS存放路徑
    new_Path=GS_NAS_Path+"\\merge" #TXT合併後存放路徑
    
    #判斷是否存在資料夾
    if os.access(new_Path, os.F_OK):
        print("路徑存在")
    else:
        #新增資料夾
        print("路徑不存在新增資料夾")
        os.mkdir(new_Path) #取至Admin管理頁面
       
    #檢查是否有GS_NAS_Path的路徑
    print("印出合併前存放 .txt的 file路徑-->" + old_Path)
    print("印出合併後存放 .txt的 file路徑-->" + new_Path)
    
    #重組目標路徑根據日期命名
    txt_new_Path = new_Path+'\\'+today+'.txt'
    print("合併後的檔名-->"+txt_new_Path)

    # 獲取路徑內文件列表+印出
    filelist = os.listdir(old_Path)
    print("------------查看old_Path內的檔案清單------------")
    print(filelist)

    # 合并文件，存在 mergeData.TXT 文件中
    print("------------合併TXT內容-->另存於新路徑------------")
    with open(txt_new_Path, 'w', encoding='utf-8') as f:
        # 构建所有文件路路徑
        for filename in filelist:
            if filename=='merge' :
                print("若資料夾內有merge資料夾不建構文件路徑")
                break
            filepath = old_Path + '\\' + filename
            # 按行寫入新的TXT文檔內
            for line in open(filepath):
                f.writelines(line)
            f.write('\n')
    txt="完成合併-->"+old_Path,"已儲存於-->"+txt_new_Path
    return txt

#完成TXT轉換CSV
def mergeTXT_to_CSV():
    import numpy as np
    import pandas as pd
    global csv_new_Path
    
    #引用TXT的存放路徑
    txt = np.genfromtxt(txt_new_Path,dtype='str')
    print(txt)
    
    txtDF = pd.DataFrame(txt)
    #print(txtDF)
    #調整CSV的儲存路徑
    csv_new_Path = new_Path +'\\'+today+'.csv'
    txtDF.to_csv(csv_new_Path,index=False) 
    txt="完成TXT轉換CSV存放於-->" + csv_new_Path
    return txt
#取得CSV內容
def getCSV():
    print("\n\n【抓取 data.csv 關於管理頁面的相關資訊】")
    print("-----------------------------------------------------------------------------------------------------------------------------")
    global csvList,csvRow
    import csv
    import sys

    nasPath_for_dataCSV = csv_new_Path
    csvRow = sum(1 for line in open(nasPath_for_dataCSV))#取得CSV總行數
    with open(nasPath_for_dataCSV, newline='',encoding='utf-8') as csvfile:     
        rows = csv.reader(csvfile,delimiter=",")# 讀取 CSV 檔案內容，分隔符號是 Tab "\t"
        csvList = []# 設定一個空陣列      
        for row in rows:# # 以迴圈輸出每一列資料加到 csvlist 陣列裡
            csvList.append(row) 
    #嘗試讀取內容
    '''
    CSV_TIME=csvList[2][0]
    CSV_AD=csvList[2][1]
    CSV_IP=csvList[2][2]
    CSV_MAC=csvList[2][3]
    print("CSV_TIME",CSV_TIME)
    print("CSV_AD",CSV_AD)
    print("CSV_IP",CSV_IP)
    print("CSV_MAC",CSV_MAC)
    return(csvList)
    '''
#獲取CSV的IP/MAC與GS進行比對
def for_csvList_get_IP_MAC():
    import sys
    global check_IP,check_MAC,update_list
    
    #特殊作法,若只有一筆資料CSV會呈現5行
    #判斷若為5行則終止程序
    print(csvRow)
    if int(csvRow)==5:
        sys.exit(0)
        
    for x in range(1, int(csvRow)):
        
        check_time = csvList[x][0]
        check_hostname = csvList[x][1]
        check_IP=csvList[x][2]
        check_MAC=csvList[x][3]
        update_list = [check_time, check_hostname, check_IP,check_MAC]
        print(csvRow)
        print("印出IP/MAC"+check_IP,check_MAC)
        
        #到GS上做IP+MAC比對
        update_IP_systemInfo_Row()
        #檢查是否需要插入新的比數或插入新的
        check_Insert_or_not()
        #再做一次更新IP清單資訊
        update_IP_systemInfo_Row()
        
        
        
#印出執行程式
print("\n--->執行程式編號:",code_Number)
get_Today()

#Admin管理頁程式開始運行
getGS_Admin_Info()#1-1~2取得管理頁面對應編號行的相關內容

print("------------------------------以上為管理頁面的資訊取得------------------------------")
#將管理頁使用的NAS路徑內的TXT文檔-->TXT合併-->轉檔CSV
check_Path_hadFile()#檢查路徑內是否有檔案
print(merge_oldPath_txt())
print(mergeTXT_to_CSV())

#比對IP+MAC 上傳到指定行數 ->若為新設備自動插入一個固定行數
#1.先取得CSV檔的所有資料
getCSV()
print(csvList)
        

#獲取頁面結果
Page_GS_GetROW_result()        
#連線PAGE1-作為硬體資訊的更新頁面--->此頁面已更新固定行的方式變動整行
#連線PAGE2-取得寫入的起始頁面--->此頁面若無IP以插入的方式新增資料        
GS_Page_UpdateInfo()
GS_Page_InsertRow()
print("------------------------------以上為資料處理取得資訊------------------------------")
print("\n\n---------------------------主功能運行-------------------------------------------")
#開始比對CSV的IP與MAC
for_csvList_get_IP_MAC()
#結束系統
sys.exit(0)
