# 自動化mail寄件

本專案是使用 Python 的**自動化郵件寄送工具**，支援透過 SMTP 寄送個人化郵件內容，亦可整合 Google Sheet 做為收件人名單來源，並將寄件結果回填紀錄。

---

## 🚀 環境準備與設定檔案

請先安裝以下 Python 套件：

```bash
pip install python-dotenv pygsheets
```

建立 `.env` 檔案（如：`project_mail.env`）

```
password="你的寄件者密碼"
```

如使用 Google Sheet，需從 GCP 建立服務帳戶，並下載 JSON 金鑰，腳本需指向此檔案路徑：

```python
service_account_file = "./credentials/your-service-account.json"
```
---

## 📁 載入設定與資料

透過 `dotenv_values`套件 從 `.env` 載入 Gmail 密碼。

輸入寄件者 Email

```python
sender_email = "XXXXXXXX@gmail.com"
```

導入收件者Email資訊：

1. **在程式中直接定義列表**

```python
receiver = ["XXXXXXXX@gmail.com","oooooooo@gmail.com"]
```

2. **透過 Google Sheet：**

讀取整張工作表

```python
worksheet.get_all_records()
```

---

## ✉️ 建立郵件

建立 `MIMEMultipart` 郵件容器

```python
message = MIMEMultipart()
```

設定郵件寄件人、收件人、寄件主旨

```python
message["From"] = sender
message["To"] = receiver
message["Subject"] = Header("輸入信件標題","utf-8").encode()
```

撰寫郵件內容

```python
body = """
輸入純文字的信件內容
"""
```

使用 `MIMEText` 封裝內文

```python
content = MIMEText(body,"plain")
message.attach(content)
```

如有附件，需讀取檔案、再使用 `MIMEBase`編碼並附加到郵件中

```python
file_dir = "./input/"
file_name = "text.csv"
file_path = file_dir + file_name
with open(file_path,"rb") as file:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(file.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename={file_name}")
```

---

## 📤 寄送郵件（SMTP）

建立 SSL 連線：

```python
context = ssl.create_default_context()
```

使用 Gmail SMTP：

```python
server = smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context)
```

登入並發送：

```python
server.login(sender_email, password)
server.sendmail(sender_email, recipient_email, message.as_string())
```

---

## 📌 處理結果與紀錄

使用 `try...except` 包裝寄送流程

* 如使用 Google Sheet：
更新該筆資料的「是否寄出成功」欄為 `"TRUE"`
更新「寄件更新日期」欄為當天日期

* 如使用地端未寄件完成：
會印出錯誤訊息並將加入 `failed_emails` 清單
將為寄件完成的筆數寫入 `./output/failed_emails.csv`

---

## 📎 備註

* 若使用 Gmail 寄信，需啟用「應用程式密碼」授權。
* 若收件人名單來自 Google Sheet，需至 google cloud console 申請服務帳戶，並下載JSON金鑰。
* 如若透過python創建新的 Google Sheet ，需開啟 Google Sheet API。

---
