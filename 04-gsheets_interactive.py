from dotenv import dotenv_values

import smtplib , ssl
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.header import Header

import pygsheets

import datetime

googlesheet = pygsheets.authorize(
    service_account_file="data-347003-eb11ef6f054b.json") # 更替.json檔

sht = googlesheet.open_by_url(
    "https://docs.google.com/spreadsheets/d/1JWYslMV8nbI_E5xNOQ1vdPBN0yErEyKBB_D04iB8aIM/edit?usp=sharing") # googlesheet連結

config = dotenv_values("project_mail.env") 
sender = "lisa51035@gmail.com"
password = config["password"]

worksheet = sht[0]
cells = worksheet.get_all_records()
# print(cells)

for i, row in enumerate(cells):
    name = row["姓名"]
    gender = row["性別"]
    mail = row["E-mail"]
    check_bool = row["是否寄出成功"]
    date = row["寄件更新日期"]

    if gender == "女":
        gender_txt = "女士"

    elif gender == "男":
        gender_txt = "先生"

    else:
        gender_txt = "先生/女士"


    if check_bool == "TRUE":
        continue

        
    message = MIMEMultipart()
    message["From"] = sender
    message["To"] = mail
    message["Subject"] = Header("輸入信件標題","utf-8").encode()

    body = f"""
            Dear {name}{gender_txt},

            輸入純文字的信件內容

            以上...
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
            server.sendmail(sender,mail,message.as_string())
        worksheet.update_value((i + 2, 4), "TRUE")
        worksheet.update_value((i + 2, 5), datetime.date.today().isoformat())
        print(f"{mail}已收到mail")

    except smtplib.SMTPException as e:
        worksheet.update_value((i + 2, 4), "FALSE")
        worksheet.update_value((i + 2, 5), datetime.date.today().isoformat())
        print(f"{mail}未收到mail")

    except Exception as e:
        worksheet.update_value((i + 2, 4), "FALSE")
        worksheet.update_value((i + 2, 5), datetime.date.today().isoformat())
        print(f"{mail}未收到mail")


print("mail寄件狀態已更新至gsheet！")
