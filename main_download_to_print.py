import os
from datetime import datetime
from gmail_utils import receive_emails_by_date_and_sender, save_attachments
from attachment_processing import process_attachments_to_excel
from ppt_pdf_utils import process_excel_to_pdf, print_pdf
import time

start_time = time.time()
# 主函数
def main():
    # date = datetime.now().strftime('%Y/%m/%d')
    date = "2024/06/25"
    from_search = "archi_lin1@yahoo.com.tw"
    
    # 步骤 1：下载邮件并保存附件
    emails = receive_emails_by_date_and_sender(date=date, from_search=from_search)
    if not emails:
        print("未找到指定日期和寄件人的邮件。")
        return
    save_attachments(emails, date)
    
    # 步骤 2：将附件转换为 Excel 数据
    process_attachments_to_excel(date)
    
    # 步骤 3：将 Excel 数据制作成 PDF
    process_excel_to_pdf(date)
    
    # 步骤 4：列印 PDF 文件
    folder_path = f"motiphysio{date.replace('/', '')}"
    pdf_folder_path = os.path.join(folder_path, "Final_PDF")
    for pdf_file in os.listdir(pdf_folder_path):
        if pdf_file.endswith(".pdf"):
            pdf_path = os.path.join(pdf_folder_path, pdf_file)
            print_pdf(pdf_path)

if __name__ == "__main__":
    main()

# 记录结束时间
end_time = time.time()

# 计算总耗时
elapsed_time = end_time - start_time
print(elapsed_time)
