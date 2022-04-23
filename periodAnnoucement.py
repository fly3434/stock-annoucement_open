####################################################################
# Warning!
# The password has been eliminated for security reason.
# You can add the password on line:222, if you have the password.
####################################################################










import requests
import pandas as pd
from io import StringIO
import datetime
import sys
import email.message
import smtplib

####################################################################
#109/10/22, 輸入before_day為前n天的日期
def ROC_today_date(before_day):  
    today = datetime.date.today()
    yesterday = (today - datetime.timedelta(days = before_day)).strftime('%Y/%m/%d')
    Y, m, d = yesterday.split('/')
    return str(int(Y)-1911) + '/' + m + '/' + d  #ex: 2020/10/22-> 109/10/22

#抓取證交所資料
def get_stock_price(stock_number): #上市股票
    def crawler(before_day):
        yesterday = (datetime.date.today() - datetime.timedelta(days = before_day)).strftime('%Y%m%d')  #抓取n天前日期
        url = "https://www.twse.com.tw/exchangeReport/MI_INDEX?response=csv&date=" + yesterday + "&type=ALLBUT0999&_=1603028506519"
        proxies = {
            'https':'https://96.113.165.182:3128'
        }
        headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36"     
        }
        res = requests.get(url,headers=headers)

        #分析有用資料
        lines = res.text.split('\n')   #使用"\n"來分開資料
        #lines
        newlines = []
        for line in lines:                  #分開後每行資料用長度來確認是否為我們要的資料
            if len(line.split('",')) > 10:
                newlines.append(line)
        return newlines
    
    before_day = 0                           #從今天的日期當初始值
    while before_day < 20:                   #從今日開始確認是否有資料,沒有則檢查前一天,以此類推
        if len(crawler(before_day)) == 0:
            before_day+=1
        else:                                          #有資料則繼續往下分析資料
            s = '\n'.join(crawler(before_day))         #再將資料用\n接起來
            newcontent = s.replace('=','')             #因原資料有"=",為不需要的字元,取代他
            df = pd.read_csv(StringIO(newcontent))     #讀取該檔案
            break

    #將數字中","去掉
    df = df.astype(str)   
    df = df.apply(lambda s:s.str.replace(',','')) 

    #設定"證券代號"為index
    df = df.set_index('證券代號')                       

    #將dataframe從str轉成float,errors='coerce'為若轉換失敗則回傳NaN
    df = df.apply(lambda s:pd.to_numeric(s, errors='coerce')) 

    #刪除NaN的欄(columns)
    #若isnull(缺值)的數量等於df的總長度,則可證明該Series全為NaN
    #df.isnull().sum() == len(df)
    #反之若isnull(缺值)的數量不等於df的總長度,則為愈留存的columns
    #df.columns[df.isnull().sum() != len(df)]
    #將為true的column留下,並print出整個dataframe
    df = df[df.columns[df.isnull().sum() != len(df)]]

    price = []
    #將輸入的stock_num全部轉為string的list
    if isinstance(stock_number, list):
        for i in stock_number:
            price.append(df.loc[i, '收盤價']) #按照位置取出收盤價格
    elif isinstance(stock_number, str):
        for i in [stock_number]:
            price.append(df.loc[i, '收盤價']) #按照位置取出收盤價格
    elif isinstance(stock_number, int):
        for i in [str(stock_number)]:
            price.append(df.loc[i, '收盤價']) #按照位置取出收盤價格

    return price

def OTC_stock_price(stock_number): #上櫃股票(over-the-counter (OTC) stock)
    def crawler(before_day):
        url = 'https://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_result.php?l=zh-tw&o=csv&d=' + ROC_today_date(before_day) + '&se=EW&_=1603376956472' #猜出來的url
        #原 https://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_result.php?l=zh-tw&o=csv&d=109/10/22&se=EW&_=1603376956472
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.75 Safari/537.36'
        }
        response = requests.get(url, headers= headers)

        lines = response.text.split('\n')    #將資料用\n分拆成好幾行
        newlines = []
        for line in lines:
            if len(line.split(',')) > 10:    #確認每行數量是否大於一定長度(用CSV檔確認),若是則貼於newlines
                newlines.append(line)
        return newlines

    before_day = 0                           #從今天的日期當初始值
    while before_day < 20:                   #從今日開始確認是否有資料,沒有則檢查前一天,以此類推
        if len(crawler(before_day)) == 0:
            before_day+=1
        else:                                #有資料則繼續往下分析資料
            s = '\n'.join(crawler(before_day))             #再將資料用\n接起來
            df = pd.read_csv(StringIO(s))                  #pandas讀CSV
            break

    #將數字中","去掉
    df = df.astype(str)
    df = df.apply(lambda s: s.str.replace(',', ''))

    #將'代號'設為index
    df = df.set_index('代號')

    price = []

    #將輸入的stock_num全部轉為string的list
    if isinstance(stock_number, list):
        for i in stock_number:
            price.append(df.loc[i, '收盤 ']) #按照位置取出收盤價格
    elif isinstance(stock_number, str):
        for i in [stock_number]:
            price.append(df.loc[i, '收盤 ']) #按照位置取出收盤價格
    elif isinstance(stock_number, int):
        for i in [str(stock_number)]:
            price.append(df.loc[i, '收盤 ']) #按照位置取出收盤價格

    return price

def emerging_stock(stock_number): #興櫃股票
    url = 'https://www.tpex.org.tw/storage/emgstk/ch/new.csv'
    #web= https://www.tpex.org.tw/web/emergingstock/lateststats/new.htm?l=zh-tw
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.75 Safari/537.36'
    }
    response = requests.get(url, headers= headers)
    response.encoding = 'big5'

    lines = response.text.split('\n')    #將資料用\n分拆成好幾行
    newlines = []
    for line in lines:
        if len(line.split(',')) > 10:    #確認每行數量是否大於一定長度(用CSV檔確認),若是則貼於newlines
            newlines.append(line)
    s = '\n'.join(newlines)             #再將資料用\n接起來
    df = pd.read_csv(StringIO(s))       #pandas讀CSV

    #將數字中","去掉
    df = df.astype(str)
    df = df.apply(lambda s: s.str.replace(',', ''))

    #將'代號'設為index
    df = df.set_index('代號')

    price = []
    #用股價輸出相對應位置之收盤價
    #將輸入的stock_num全部轉為string的list
    if isinstance(stock_number, list):
        for i in stock_number:
            price.append(df.loc[i, '成交']) #按照位置取出收盤價格
    elif isinstance(stock_number, str):
        for i in [stock_number]:
            price.append(df.loc[i, '成交']) #按照位置取出收盤價格
    elif isinstance(stock_number, int):
        for i in [str(stock_number)]:
            price.append(df.loc[i, '成交']) #按照位置取出收盤價格

    return price

def send_mail(df, receiver, ratio):  
    subject = '今日可申購股票: '  
    subject_bool = True
    for i in df.iterrows():
        #print(i)
        if subject_bool == True:
            subject = subject + i[0]
            subject_bool = False
        else:
            subject = subject + ", " + i[0]
    line1 = '<p>'+ subject + '</p>'
    line2 = '<p>' + "(將於明日 " + df.iloc[0]['申購結束日'] + " 14:00 截止)" + '</p>'
    #print(subject)
    #print(line1) 

    line3 = '<ol>'
    for i in df.iterrows():
        #print(i[0])
        line3 = line3 + '<li>' + i[0] + ' ' + i[1]['證券名稱'] + ',市場別: ' + i[1]['發行市場']   #i[0]為股票代號(index的值),i[1]為該股票後面的值
        line3 = line3 + ',申購價: ' + str(i[1]['實際承銷價(元)']) + ',現價: ' + str(i[1]['收盤價'])
        line3 = line3 + ',抽籤日期: ' + i[1]['抽籤日期'] + ',匯入集保日: ' + i[1]['撥券日期(上市、上櫃日期)'] + '</li>'
    line3 = line3 + '</ol>'
    #print(line3)

    line4 = '<p>'+ "※本程式抓取截止日前一天、且溢價差大於" + str(ratio) + "%之申購股票" + '</p>'
    
    line5 = '<a href="https://www.twse.com.tw/zh/page/announcement/publicForm.html"> Please visit this link for more infomation.</a>'
    content = line1 + line2 + line3 + line4 + line5
    
    ############### mail function ##################
    msg=email.message.EmailMessage()
    msg["From"]="vlegosov@gmail.com"
    #msg["To"]=
    msg['Bcc']=receiver
    msg["Subject"]=subject
    
    
    msg.set_content(content, subtype='html')
    
    server=smtplib.SMTP_SSL("smtp.gmail.com", 465)    #連接伺服器 (可搜尋gmail smtp server)
    server.login("vlegosov@gmail.com", "ENTER THE PASSWORD")  #此密碼為google自動產生之應用程式專用密碼
    server.send_message(msg)
    server.close()

###########################################################################

ratio = 10
this_year = datetime.date.today().strftime("%Y")
url = 'https://www.twse.com.tw/announcement/publicForm?response=csv&yy=' + this_year
#檢視網址: https://www.twse.com.tw/zh/page/announcement/publicForm.html
headers = {
    'User-Agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.75 Safari/537.36'
}
res = requests.get(url, headers=headers)

lines = res.text.split('\n')
newlines = []
for line in lines:
    if len(line.split('",')) > 10:
        newlines.append(line)
s = '\n'.join(newlines)
df = pd.read_csv(StringIO(s))

df = df[df.發行市場 != '中央登錄公債'] #將df排除'中央登錄公債'

#test_day = '109/10/23'
#將申購結束日為今天的股票raw存到today_df
#today_ann = df['申購結束日'] == test_day
today_ann = df['申購結束日'] == ROC_today_date(-1)  #申購結束日等於今日日期(國歷),則回傳true
today_df = df[today_ann]

#若今日申購結束之股票(false長度等於全部長度)，則回傳no data並直接結束execution
# if today_ann.value_counts().loc[False] == len(df):
#     print("No correspond stock found today")
#     sys.exit()

#將股票代號儲存在list: stock_num
stock_num = today_df['證券代號']

#將'證券代號'設為index
today_df = today_df.set_index('證券代號')

stock_price = []
for num in stock_num:
    if today_df.loc[num, '發行市場'] == '上市增資':
        stock_price.append(get_stock_price(num)[0])  #回傳為list,轉成str(藉由貼上回傳list的第一個值)
    elif today_df.loc[num, '發行市場'] == '第一上市公司現金增資':
        stock_price.append(get_stock_price(num)[0])
    elif today_df.loc[num, '發行市場'] == '初上市':                 #待確認是否有其他可能
        stock_price.append(emerging_stock(num)[0])
    #elif today_df.loc[num, '發行市場'] == '第一上市公司初上市':      #待確認
    #    stock_price.append(get_stock_price(num)[0])
    elif today_df.loc[num, '發行市場'] == '上櫃增資':
        stock_price.append(OTC_stock_price(num)[0])
    elif today_df.loc[num, '發行市場'] == '第一上櫃公司現金增資':
         stock_price.append(OTC_stock_price(num)[0])
    elif today_df.loc[num, '發行市場'] == '初上櫃':
        stock_price.append(emerging_stock(num)[0])
    else:  
        stock_price.append(0)  #抓取不到上述選項則stock_price設為0

today_df.insert(11, '收盤價', stock_price)
#print(today_df['收盤價'])
good_stock = today_df['收盤價'].astype(float) / today_df['實際承銷價(元)'].astype(float) >= (100+ratio) / 100
valid_df = today_df[good_stock]

#開Recipients.txt讀收件者
with open('Recipients.txt', 'r') as file:
    recipients = file.read()

if __name__ == "__main__":
    if valid_df.empty == True:  #判斷是否為空df(即今日是否有資料)
        print("No stock can be annouced today.")
    else: 
        #print(valid_df)
        print("Mail is going to be sent...")
        send_mail(valid_df, recipients, ratio)
        print("Mail has been sent.")



