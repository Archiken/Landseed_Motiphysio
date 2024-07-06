import os
import base64
import re
import time
from simplegmail import Gmail
from datetime import datetime, timedelta
from googleapiclient.errors import HttpError

# 初始化 Gmail 客户端
gmail = Gmail()

# 功能函数：下载 email 附件
def receive_emails_by_date_and_sender(date="2024/06/18", from_search="service@motiphysio.com"):
    # 此處添加你接收郵件的程式碼
    try:
        date_obj = datetime.strptime(date, "%Y/%m/%d")
        day_before = (date_obj - timedelta(days=1)).strftime("%Y/%m/%d")
        day_after = (date_obj + timedelta(days=1)).strftime("%Y/%m/%d")

        query = f'from:{from_search} after:{day_before} before:{day_after}'
        response = gmail.service.users().messages().list(
            userId='me',
            q=query
        ).execute()

        messages = response.get('messages', [])
        
        email_data = []
        for msg in messages:
            message = gmail.service.users().messages().get(
                userId='me', id=msg['id']
            ).execute()

            email_info = {
                "id": msg['id'],
                "subject": next(header['value'] for header in message['payload']['headers'] if header['name'] == 'Subject'),
                "sender": next(header['value'] for header in message['payload']['headers'] if header['name'] == 'From'),
                "body": get_message_body(message['payload']),
                "attachments": get_attachments(message['payload'], msg['id'])
            }
            email_data.append(email_info)
            
            gmail.service.users().messages().modify(
                userId='me',
                id=msg['id'],
                body={'removeLabelIds': ['UNREAD']}
            ).execute()

        return email_data
    except HttpError as error:
        return []
    except Exception as e:
        return []
    

def get_message_body(payload):
    # 此處添加你獲取郵件正文的程式碼
    body = ''
    if 'parts' in payload:
        for part in payload['parts']:
            if part['mimeType'] == 'text/plain':
                body += base64.urlsafe_b64decode(part['body'].get('data', '')).decode('utf-8')
            elif part['mimeType'] == 'text/html':
                body += base64.urlsafe_b64decode(part['body'].get('data', '')).decode('utf-8')
    else:
        body += base64.urlsafe_b64decode(payload['body'].get('data', '')).decode('utf-8')
    
    return body

def get_attachments(payload, message_id):
    # 此處添加你獲取附件的程式碼
    attachments = []
    if 'parts' in payload:
        for part in payload['parts']:
            if 'filename' in part and part['filename']:
                if 'data' in part['body']:
                    data = part['body']['data']
                else:
                    att_id = part['body']['attachmentId']
                    attachment = gmail.service.users().messages().attachments().get(
                        userId='me', messageId=message_id, id=att_id
                    ).execute()
                    data = attachment['data']
                file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
                attachments.append({
                    "filename": part['filename'],
                    "data": file_data
                })
    return attachments

def extract_name(subject):
    # 此處添加你提取名稱的程式碼
    match = re.search(r"results of ([\w\u4e00-\u9fa5 ]+)'s", subject)
    if match:
        return match.group(1).strip()
    else:
        return "image"

def save_attachments(email_data, date):
    # 此處添加你保存附件的程式碼
    folder_name = f"motiphysio{date.replace('/', '')}"
    if not os.path.isdir(folder_name):
        os.mkdir(folder_name)

    for email in email_data:
        png_downloaded = 0
        subject = email['subject']
        name = extract_name(subject)

        for attachment in email['attachments']:
            if attachment['filename'].lower().endswith('.png') and png_downloaded < 2:
                image_filename = f"{name}_{png_downloaded + 1}.png"
                filepath = os.path.join(folder_name, image_filename)
                with open(filepath, "wb") as f:
                    f.write(attachment['data'])
                print(f"Downloaded: {attachment['filename']} to {filepath}")
                png_downloaded += 1
