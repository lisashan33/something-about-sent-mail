from dotenv import dotenv_values

import smtplib , ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header

import csv

config = dotenv_values("project_mail.env")  # 密碼存取區

sender = "XXXXXXXX@gmail.com" # 寄件者mail
password = config["password"] # 寄件者密碼
receiver = ["XXXXXXXX@gmail.com","OOOOOOOO@gmail.com"] # 收件者

message = MIMEMultipart()

failed_emails = [] # 存取未成功寄件的mail
for i in receiver:
    
    message = MIMEMultipart()
    message["From"] = sender
    message["To"] = i
    message["Subject"] = Header("輸入信件標題","utf-8").encode()

    body = """
輸入純文字的信件內容
"""

    content = MIMEText(body,"plain")
    message.attach(content)


    con = ssl.create_default_context()
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com",465,context=con) as server:
            server.login(sender,password)
            server.sendmail(sender,i,message.as_string())

        print(f"{i}已收到mail")

    except smtplib.SMTPException as e:
        print(f"收件人 {i} 未收到信件，發送錯誤原因: {e}")
        failed_emails.append(i)

    except Exception as e:
        print(f"寄件時發生其他錯誤：{e}")
        failed_emails.append(i)

if failed_emails:
    with open("./output/failed_emails.csv", mode="w", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(['Email'])

        for email in failed_emails:
            writer.writerow([email])
        print(f"未成功發送的郵件明細詳如附件，檔案名稱：failed_emails.csv")
else:
    print("所有mail皆已成功寄出！")

    
