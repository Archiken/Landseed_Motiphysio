import pandas as pd
import os
from simplegmail import Gmail
import base64

# Initialize Gmail client
gmail = Gmail()

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

# Function to read and convert image to base64
def image_to_base64(image_path):
    with open(image_path, 'rb') as img_file:
        return base64.b64encode(img_file.read()).decode('utf-8')

# Convert image to base64
image_base64 = image_to_base64(local_image_path)
img_tag = f'<img src="data:image/png;base64,{image_base64}" alt="Line Group" style="width:200px;height:auto;">'

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
                # Build and send the email
                gmail.send_message(
                    sender="archilin1@gmail.com",
                    to=email,
                    bcc=["archilin1@gmail.com", "lssmctpe@gmail.com"],  # BCC to self for a copy
                    subject="聯新企業健檢報告",
                    msg_html=f'''
                    <p>您好，</p>
                    <p>這是上週您參加聯新運醫AI 3D姿態檢測，<br>
                    依據您的檢測結果，為您設計最需處理的3項運動。</p>
                    <p>請參考您的運動處方建議，如有任何問題，歡迎您先加入Line官方帳號，並利用line詢問我們!!</p>
                    {img_tag}
                    <p>敬祝您身體常保健康</p>
                    <p>聯新運醫<br>
                    企業健康促進小組<br>
                    02-27216698</p>
                    ''',
                    attachments=[full_path]  # Attach the PDF
                )
                print(f"Email successfully sent to {email}")
            except Exception as e:
                print(f"Failed to send email to {email}: {e}")
        else:
            print(f"找不到 {name} 的電子郵件")
