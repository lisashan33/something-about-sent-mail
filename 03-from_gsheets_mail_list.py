from dotenv import dotenv_values

import smtplib , ssl
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.header import Header

import pygsheets

import csv

googlesheet = pygsheets.authorize(
    service_account_file="data-347003-eb11ef6f054b.json") # 更替.json檔

sht = googlesheet.open_by_url(
    "https://docs.google.com/spreadsheets/d/1JWYslMV8nbI_E5xNOQ1vdPBN0yErEyKBB_D04iB8aIM/edit?usp=sharing") # googlesheet連結

config = dotenv_values("project_mail.env") 
sender = "lisa51035@gmail.com"
password = config["password"]

worksheet = sht[0]

# 取得收件者mail
receiver = []
for i, row in enumerate(worksheet):
    email = row[0]
    receiver.append(email)

# print(receiver[1:])

failed_emails = [] # 存取未成功寄件的mail
for i in receiver[1:]:
    
    message = MIMEMultipart()
    message["From"] = sender
    message["To"] = i
    message["Subject"] = Header("輸入信件標題","utf-8").encode()

    body = """
輸入純文字的信件內容
"""
    file_dir = "./input/"
    file_name = "text.csv"
    file_path = file_dir + file_name
    with open(file_path,"rb") as file:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={file_name}")


    content = MIMEText(body,"plain")
    message.attach(content)
    message.attach(part)

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
