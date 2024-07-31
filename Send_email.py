import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.mime.image import MIMEImage


# Email account credentials
EMAIL_ADDRESS = 'lssmctpe@gmail.com'
EMAIL_PASSWORD = ''


# Path to the Excel file
excel_path = r"C:\Users\lin\Documents\program\toss\Landseed\0614moti 結果\25人_參與者名單(完成) - 排序.xlsx"

df = pd.read_excel(excel_path)

# Get names and emails
name_email_map = pd.Series(df.iloc[:, 16].values, index=df.iloc[:, 2]).to_dict()

# Directory for PDF files
base_dir = os.path.dirname(excel_path)
pdf_folder = os.path.join(base_dir, "Final_PDF")

# Path to the local image
local_image_path = r"C:\Users\lin\Documents\program\toss\Landseed\LINE group.jpg"  # Update to your actual image path

# Function to format names
def format_name(name):
    return name.replace(" ", "")  # Remove spaces from names

# Function to send email
def send_email(to_address, pdf_path):
    # Create the email
    msg = MIMEMultipart('related')
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = to_address
    msg['BCC'] = 'archilin1@gmail.com'  # Add BCC field
    msg['Subject'] = "聯新企業健檢報告"

    # Create the body of the email
    msg_html = f'''
    <p>您好，</p>
    <p>這是上週您參加聯新運醫AI 3D姿態檢測，<br>
    依據您的檢測結果，為您設計最需處理的3項運動。</p>
    <p>請參考您的運動處方建議，如有任何問題，歡迎您先加入Line官方帳號，並利用line詢問我們!!</p>
    <p><img src="cid:image1" alt="Line Group" style="width:200px;height:auto;"></p>
    <p>敬祝您身體常保健康</p>
    <p>聯新運醫<br>
    企業健康促進小組<br>
    02-27216698</p>
    '''
    msg.attach(MIMEText(msg_html, 'html'))

    # Attach the image
    with open(local_image_path, 'rb') as img_file:
        img = MIMEImage(img_file.read())
        img.add_header('Content-ID', '<image1>')
        msg.attach(img)

    # Attach the PDF
    with open(pdf_path, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(pdf_path))
        msg.attach(part)

    # Send the email
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.sendmail(EMAIL_ADDRESS, [to_address, 'archilin1@gmail.com'], msg.as_string())

# Traverse all PDF files
for filename in os.listdir(pdf_folder):
    if filename.endswith("_moti中文報告.pdf"):
        # Extract and format name
        name = filename[:-len("_moti中文報告.pdf")]
        formatted_name = format_name(name)
        
        # Find corresponding email
        email = name_email_map.get(formatted_name)
        if email:
            full_path = os.path.join(pdf_folder, filename)
            try:
                # Send the email
                send_email(email, full_path)
                print(f"Email successfully sent to {email}")
            except Exception as e:
                print(f"Failed to send email to {email}: {e}")
        else:
            print(f"找不到 {name} 的電子郵件")