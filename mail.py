import win32com.client
import openpyxl as pyxl
import datetime
import requests

FILEPATH = "C:/Users/FUN-PC132/Downloads/お問い合わせ集計.xlsx"

WB = pyxl.load_workbook(FILEPATH)
WS = WB["Sheet1"]
end_row = WS.max_row

TODAY  =datetime.date.today()
Y = TODAY.year
M = TODAY.month
D = 21#TODAY.day

SELECT_DATE = "{}-{}-{}".format(Y,str(M).zfill(2),str(D).zfill(2))

TargetAddress = "no-reply@shop-pro.jp"

Shop = "FUNオンライン"

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

accounts = outlook.Folders
inbox = outlook.GetDefaultFolder(6)

mails = inbox.Items
print(mails)
counter = 1
comment_counter = 0
for mail in mails:
    
    
    if str(SELECT_DATE) in  str(mail.receivedtime) :


        if ( TargetAddress == mail.sendername ) & ("お問い合わせ" in mail.subject) & ("Re:" not in mail.subject):
            print(mail.subject)
   
            def Sender():

                target1 = "【  お名前  】"
                target2 = "【   MAIL   】"
                
                idx_start = str(mail.body).find(target1)
                find_sender = str(mail.body)[idx_start + len(target1):]
                idx_end = find_sender.find(target2)
                find_sender2 = find_sender[:idx_end + len(target2)].replace(target2,"")

                return find_sender2
            
            def Title():

                target1 = "【 タイトル 】"
                target2 = "【"
                
                idx_start = str(mail.body).find(target1)
                find_title = str(mail.body)[idx_start + len(target1):]
                idx_end = find_title.find(target2)
                find_title2 = find_title[:idx_end + len(target2)].replace(target2,"")

                return find_title2
            
            
            def Text():

                target1 = "【   内容   】"
                target2 = "=============================================================="
                
                idx_start = str(mail.body).find(target1)
                find_text = str(mail.body)[idx_start + len(target1):]
                idx_end = find_text.find(target2)
                find_text2 = find_text[:idx_end + len(target2)].replace(target2,"").replace("　","").lstrip()

                return find_text2
            
            def DateTime():

                target1 = "送信日時："
                target2 = "送信元IP"
                
                idx_start = str(mail.body).find(target1)
                find_datetime = str(mail.body)[idx_start + len(target1):]
                idx_end = find_datetime.find(target2)
                find_datetime2 = find_datetime[:idx_end + len(target2)].replace(target2,"")[:11]
                str_y = find_datetime2[:4]
                str_m = int(find_datetime2[5:7])
                str_d = int(find_datetime2[8:10])
                
                RecreateDate = "{}/{}/{}".format(str_y,str_m,str_d)
                
                return RecreateDate#find_datetime2
            datetime_str = DateTime()
            sender = Sender()
            title = Title()
            text = Text()
            
            print(sender,title,text,datetime_str)
            comment_counter += 1

            WS["B" + str(end_row + counter)].value = Shop
            WS["C" + str(end_row + counter)].value = datetime_str
            WS["D" + str(end_row + counter)].value = sender
            WS["E" + str(end_row + counter)].value = title
            WS["F" + str(end_row + counter)].value = text
            
            counter += 1


#WB.save(FILEPATH)       

print(comment_counter)    
if comment_counter > 0 :
    TOKEN = 'NxPDQg0tpI4oY6BZHo7vkZ3gxPtTIijpWFyN85xL2q1'#テストの部屋トークン
    #TOKEN = 'xE4tqcs5xBQ1mYRS8WsH7Gf5btAd8ypoEahCsRtt54h'#オンラインの部屋
    #TOKEN = 'TNKXBcpEMmK4JAmRPaOyVABkA5GWIJkIHQOnsyfu4MD'#FUNの部屋トークン
    api_url = 'https://notify-api.line.me/api/notify'
    headers = {'Authorization' : 'Bearer ' + TOKEN}

    message_1 = (
        '\n【お問い合わせ報告】\n\nお客様よりお問い合わせが来ております。\n\n集計日 : {}\n件数 : {}件\n\nご確認お願いします。'.format(datetime_str,comment_counter) 

    )
    payload = {'message': message_1}
    requests.post(api_url, headers=headers, params=payload)   
    print("SUCCESSFULL!!")

