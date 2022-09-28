'''
pyinstaller -F main_v2.pyw

由CSV文件檔取得
#主要是讓所有的自動化程式到管理頁面進行 "控制" 與 "管控"
'''
print("-------------------------------抓取 data.csv 關於管理頁面的相關資訊-------------------------------")
import csv
with open('data.csv', newline='',encoding='utf-8') as csvfile:    
    rows = csv.reader(csvfile,delimiter=",")# 讀取 CSV 檔案內容，分隔符號是 Tab "\t"
    csvList = []# 設定一個空陣列      
    for row in rows:# # 以迴圈輸出每一列資料加到 csvlist 陣列裡
        csvList.append(row)

#程式編號取得於CSV
print("本程式抓取CSV內的檔名:",csvList[1][0])
program_number= csvList[1][0]

#自動化程式的管理列表------->提供程式連入取得程式應用的相關資訊
print("程式管理頁面URL:",csvList[1][1])
print("程式管理頁面ID:",csvList[1][2])
print("程式管理頁面PAGE:",csvList[1][3])
URL_Info = csvList[1][1]
ID_Info = csvList[1][2]
GSheet_listPage_name = csvList[1][3]

print("-------------------------------保留主程式初始變數-------------------------------")

#保留給主程式使用變數
codeName = ""
Googlesheet_URL = ""
Googlesheet_ID = ""
Page_name1 = ""
Page_name2 = ""


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
    creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope) #權限金鑰
    read_sheet_API= gspread.authorize(creds) .open_by_key(ID_Info).worksheet(GSheet_listPage_name)
    
    result = read_sheet_API.get_all_values()
    return result
'''
抓取主程式使用的必要資訊
啟動功能 / 程式編號 / 程式檔名 / GooglesheetURL / GooglesheetID / PageName1 / PageName2
'''
def getGSheet_():
    global row_end, df, codeName, Googlesheet_URL, Googlesheet_ID, Page_name1, Page_name2
    import sys
    import pandas as pd
    df = pd.DataFrame(all_Auto_GSeet_List())  #取得回傳值read_test()-->內文陣列values[DEVICE_IP][USED_CELL]
    
    #取得頁面行數
    row_end = df.shape[0]           
    print("目前清單建立比數:",row_end)#-2為去頭尾
    
    #取得程式編號使用的 [url,id,page1,page2] 
    for x in range(0, int(row_end)):
        
        if df.loc[x][0]=="TRUE" and df.loc[x][1]==program_number:
            print("檢查程式功能啟動:",df.loc[x][0])
            codeName = df[2][x]
            Googlesheet_URL = df[3][x]
            Googlesheet_ID = df[4][x]
            Page_name1 = df[5][x]
            Page_name2 = df[6][x]
            
            print ("程式編號:" , program_number)
            print ("程式命名:" , codeName)
            
            print ("獲取的URL:" , Googlesheet_URL)
            print ("獲取的ID:" , Googlesheet_ID)
            print ("獲取的PAGE1:" , Page_name1)
            print ("獲取的PAGE2:" , Page_name2)
            
            break
        elif df.loc[x][0]=="FALSE" and df.loc[x][1]==program_number:
            codeName = df[2][x]
            print("codeName:",program_number,"codeName:",codeName,"is not used !")
            sys.exit(0)
#取得所需主程式資訊
getGSheet_()






'''
程式開發者:NessHuang
開發日期:2022-09-23
版本:v1

使用平台:window10 / Google Sheet
開發語言:python

主程式功能:
1取得本機資訊{IP/AD/MAC}
2透過IP先在GOOGLE SHEET上比對
 └---檢查IP&MAC無差異---->在分頁【Get_ROW_LIST】 取得更新的CELL--->3
 └---檢查IP&MAC有差異---->在分頁【Get_ROW_LIST】 寫入IP/MAC/CELL--->回上一動
3.在分頁【IP_MAC_INFO】更新myInfo[get_TIME(),get_hostname(),get_IP(),getMAC()]於指定CELL
'''

#定義使用變數
URL_Info = Googlesheet_URL
ID_Info = Googlesheet_ID

#Google分頁名------->IP_MAC_INFO->以取得的CELL為初始行數寫入TIME_AD_IP_MAC
#Google分頁名------->IP_MAC_CELL比對表->作為取得CELL為初始行數
#設定 IP & MAC 比對表的範圍
showInfo_by_googlePage = Page_name1
checkList_by_googlePage = Page_name2 


df=[]                           #暫存取得的PANDAS陣列----->Google分頁Get_ROW_LIST
row_end=""                      #暫存取得----->Google分頁Get_ROW_LIST總行數
do_insert_or_not=0              #作為判斷值是否為新的IP與MAC 0為初值 1為需要插入IP/MAC/CELL ----->Google分頁Get_ROW_LIST
#next_cell="A10"                #測試給予CELL附值
#check_MAC="c8:d9:d2:03:27:48"  #測試給予MAC附值


'''
功能一取得本機資訊
將各個自訂函式整入myInfo[]
get_TIME(),get_hostname(),get_IP(),getMAC()
'''
def getMAC():
    import uuid             #取得本機MAC位址#https://www.796t.com/content/1548664229.htm
    mac=uuid.UUID(int = uuid.getnode()).hex[-12:]
    return ":".join([mac[e:e+2] for e in range(0,11,2)])

def getSystem_Info():    
    from datetime import datetime                                #datetime函式https://www.delftstack.com/zh-tw/howto/python/how-to-get-the-current-time-in-python/
    import socket                                                #socket函式引用取用IP https://shengyu7697.github.io/python-get-ip/
    import uuid
    import requests         #透過第三方"取得對外IP"#https://shengyu7697.github.io/python-get-ip/
    
    time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')          #時間 
    hostname = socket.gethostname()                              #主機名稱

    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)         
    s.connect(("8.8.8.8", 80))
    ip = s.getsockname()[0]                                      #內部IP
    s.close()
    #ip_external = requests.get('https://api.ipify.org').text     #外部IP----->會跳動抓也無意義
    
    myInfo=[time,hostname,ip,getMAC()]
    return myInfo

'''
Open_URL(URL_Info)
#自訂開啟瀏覽器

insertData_googleSeet_API_Key(ID_Info,checkList_by_googlePage)
#自訂GOOGLE SHEET的頁面連線---->插入IP_CEL的比對資訊為目的

updateData_googleSeet_API_Key(ID_Info,showInfo_by_googlePage)
#自訂GOOGLE SHEET的頁面連線---->更新TIME/AD/IP/MAC為目的
'''
def Open_URL():
    import webbrowser
    webbrowser.open(URL_Info)

    
def insertData_googleSeet_API_Key():
    global insert_sheet_API
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    import pandas as pd 
    #獲取授權與連結#
    scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope) #權限金鑰
    insert_sheet_API = gspread.authorize(creds) .open_by_key(ID_Info).worksheet(checkList_by_googlePage)# ID + PAGE

def updateData_googleSeet_API_Key():
    global update_sheet_API
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    import pandas as pd
    #獲取授權與連結#
    scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("./creds.json", scope) #權限金鑰
    update_sheet_API = gspread.authorize(creds) .open_by_key(ID_Info).worksheet(showInfo_by_googlePage)# ID + PAGE

'''
read_test()
讀取google sheet內文+總行數-->回傳values,rowCount

check_valuse()
透過pandas取的所需參數row_end,df
df->取得回傳值read_test()-->內文陣列values[DEVICE_IP][USED_CELL]
row_end->最後行數
'''
def readData_googleSeet_API_Key():
    global read_sheet_API
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    import pandas as pd 
    #獲取授權與連結#
    scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope) #權限金鑰
    read_sheet_API= gspread.authorize(creds) .open_by_key(ID_Info).worksheet(checkList_by_googlePage)# ID + PAGE
    
    result = read_sheet_API.get_all_values()
    return result

def check_valuse():
    global row_end,df
    import pandas as pd
    df = pd.DataFrame(readData_googleSeet_API_Key())  #取得回傳值read_test()-->內文陣列values[DEVICE_IP][USED_CELL]
    row_end = df.shape[0]           #總行數統計迴圈使用......印出目前行數
    print("目前清單建立比數:",row_end-2)#-2為去頭尾

'''
update_IP_systemInfo_Row()
check_valuse()引用read_test()內文做[行][欄]
if-else判斷式___本機IP == df.loc[0][0]-->回傳df.loc[0][+1]
若IP不等於df.loc[+1][0]直到到迴圈結束
-->IP插入Google sheet的比對取值表 
'''
def update_IP_systemInfo_Row():
    global do_insert_or_not
    check_valuse()
    
    #print("IP:",df.loc[0][0])
    #print("MAC:",df.loc[0][1])
    #print("CELL:",df.loc[0][2])
    
    #for迴圈 逐行檢查IP確認是否相符
    for x in range(1, int(row_end)): 
        if df.loc[x][0]==check_IP and df.loc[x][1]==check_MAC:
            print("找到",check_IP,"對應欄位=",df.loc[x][2])
            select_Cell=df.loc[x][2]
            print("本機資訊:",myInfo)
            update_Row = [ myInfo ]
            print("更新資訊於IP_sysInfo:",myInfo)
            update_sheet_API.update(select_Cell, update_Row) #更新Array
            do_insert_or_not=0
            break
        else :
            print(check_IP,df.loc[x][0],"無IP對應欄位")
            do_insert_or_not=1

'''
insert_ip_cell()
#判斷IP是否有符合項給值,
---有IP---->在分頁【Get_ROW_LIST】 取得CELL---->在分頁【IP_MAC_INFO】更新{IP/AD/MAC}於指定CELL
---無IP---->在分頁【Get_ROW_LIST】 插入IP與CELL---->回到上一動
PS.注意CELL為更新的起始位置使用
'''
def insert_ip_cell():
    global do_insert_or_not
    #迴圈結束無建立IP與CELL對照表---->插入IP與新的CELL
    check_valuse()
    next_cell="A"+str(row_end)

    if do_insert_or_not==1:
        insert_Row=[check_IP,check_MAC,next_cell]
        select_Row=2
        insert_sheet_API.insert_row(insert_Row, select_Row)#插入List
        print("寫入資訊內容:",insert_Row,"插入行數:",select_Row)
        print("完成插入新IP與CELL",)
    do_insert_or_not=0
    print (do_insert_or_not)

'''
下方程式開始
'''
#Open_URL()

readData_googleSeet_API_Key()
updateData_googleSeet_API_Key()#更新-主要查詢IP設備資訊使用
insertData_googleSeet_API_Key()#插入-IP & CELL 對照表使用

print("-------------------------------取得本機資訊-------------------------------")
myInfo=getSystem_Info()               #取得本機資訊

#變數假定IP AND MAC
#check_IP="10.231.220.149"
check_IP=myInfo[2]#待檢查IP資訊
print("本機IP:",myInfo[2])
#check_MAC="c8:d9:d2:03:27:48"
check_MAC=myInfo[3]#待檢查MAC資訊
print("本機MAC:",myInfo[3])

update_IP_systemInfo_Row()
insert_ip_cell()

update_IP_systemInfo_Row()
