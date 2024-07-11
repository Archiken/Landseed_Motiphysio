import pandas as pd
import os
from simplegmail import Gmail

# Initialize Gmail client
gmail = Gmail()

# Path to the Excel file
excel_path = r"C:\Users\lin\Documents\program\toss\Landseed\0614moti 結果\25人_參與者名單(完成) - 排序.xlsx"

df = pd.read_excel(excel_path)

# Get names and emails
name_email_map = pd.Series(df.iloc[:, 8].values, index=df.iloc[:, 2]).to_dict()

# Directory for PDF files
base_dir = os.path.dirname(excel_path)
pdf_folder = os.path.join(base_dir, "Final_PDF")

# Direct URL to the LINE group image
line_group_image_url = "https://lh3.googleusercontent.com/pw/AP1GczMf-eW2Cp7eUp3r9Hn1BuuGpcK3VQDORHetMNvPhlKXNIHihirp6dzfLVoNlvevbyJV--JZUYzx213C4t43CHTG90a46_ZEoBoGB-SEwd6BHoyFo2GMjZfCUgiT63YAOF89-HQyNY7GQ8rxBLoXR8shzyku5mvvufXBt1wH-cpI5tooznri4xHMxwKgLPL6XaLU7bcQGn_IE0XOjEjztKhY608kMWMzFhC9ETh4mQuHi2vnWv1w6EIsODYiMffuwkS-vcgSMqDV1Aqu-lrJdFfggyu726mxVFEd4U2_m2wDNwOnHMKyTGbitD5DtqI7kj-xyKCE7qDphAiOWbDIMFlRjvagy8YeSR1QzzAaWaw1-wVz3h6HFJc0Pk3aG9Zk7qPWc1O-eFuqTUEVK1qouabWdNDbh8yVUJMyS8luJJ2y2tReO_N5UHAmqR_9afJu3Y_0d_epBSARzFqkEmh1T5WQjrxd1sKNugIgfO2GgmlQihW5lScYrAwPazF2-CPAPuUX1J8_MgzzpK_mbMvcw6RApXNGBAqJ0tP0VYGLp62B8UsZNPU2J4smWKHIs0IK2SWwIXRH3NKL8Q6mAUsn3jssOJ0MlfkQaZCC-oh9R5VFFv43BKLaKv1psUMExkXDK-n4gE5XU1SoOOHe7KFmUTohtOy9w961ThWSzjhmDY8w8J-FuW4Fo-5f7ceLdjTFWkk8i_eXJKXppdH0od6hX4aQkxctpK0V5IJySovGWr9aWUOxpyCEK-yZC0a7oLazopg3saDSmglk26MzpBysmdA1iYNK0BUBZIpyDjae6N72a84VM9NeONub-P1mxt33QE8OcHhWYz1e23OQcWnFHFnhvLGIAbOzZ3fAMxH9KeUWHRAPN0xKF9zIHga30OneCvrwrye1Y98tU1UiQnSEkrif-c8Aw_we6Wop5imMgzlCGOXuyAQBHxXCcYovwg=w540-h540-s-no?authuser=0"  # This is your image URL

# Function to format names
def format_name(name):
    return name.replace(" ", "")  # Remove spaces from names

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
                    bcc=["archilin1@gmail.com", "hayatosa@gmail.com"],  # BCC to self for a copy
                    subject="聯新企業健檢報告",
                    msg_html=f'''
                    <p>您好，</p>
                    <p>這是上週您參加聯新運醫AI 3D姿態檢測，<br>
                    依據您的檢測結果，為您設計最需處理的3項運動。</p>
                    <p>請參考您的運動處方建議，如有任何問題，歡迎您先加入Line官方帳號，並利用line詢問我們!!</p>
                    <p><img src="{line_group_image_url}" alt="Line Group" style="width:200px;height:auto;"></p>
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
