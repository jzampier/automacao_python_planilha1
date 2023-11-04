from icecream import ic
import os
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText
from dotenv import load_dotenv

# path and list of files
path = "bases"
files = os.listdir(path)

# Create consolidated worksheet
consolidated_worksheet = pd.DataFrame()

# Loop through files
for file_name in files:
    full_path = os.path.join(path, file_name)
    try:
        sales_table = pd.read_csv(full_path)
        consolidated_worksheet = pd.concat([consolidated_worksheet, sales_table])
    except Exception as e:
        ic(f"Error reading file {full_path}: {e}")

consolidated_worksheet = consolidated_worksheet.sort_values(by='first_name')
consolidated_worksheet = consolidated_worksheet.reset_index(drop=True)
ic(consolidated_worksheet)

# Save consolidated worksheet in an Excel file
consolidated_worksheet.to_excel('Sales.xlsx', index=False)
excel_file: str = 'Sales.xlsx'
# Email routine
load_dotenv()
# environment variables
SENDER = os.getenv('SENDER')
PASSWORD = os.getenv('PASSWORD')
RECEIVER = os.getenv('RECEIVER')


def send_email(attachment_file: str):
    # Prepare the email (message + attachment)
    msg = MIMEMultipart()
    msg['From'] = SENDER
    msg['To'] = RECEIVER
    msg['Subject'] = 'Sales Report'
    body = """<p><b>Relatório de vendas</b></p>
                    <p>Prezado(a) gerente do setor de vendas.</p>
                    <p></p>
                    <p></p>
                    <p>Segue anexado o relatório diário de vendas com os dados atualizados.</p>
                    <p>Caso tenha algum problema, favor nos avisar.</p>
                    <p></p>
                    <p></p>
                    <p>Cordialmente,</p>
                    <p>Análise de dados</p>"""
    msg.attach(MIMEText(body, 'html'))
    file = attachment_file
    try:
        with open(file, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename={file}")
            msg.attach(part)
    except Exception as e:
        ic(f"Error attaching file {file}: {e}")
        return

    # Send the email
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER, PASSWORD)
        server.sendmail(SENDER, RECEIVER, msg.as_string())
        server.quit()
        ic('Email sent!')
    except Exception as e:
        ic(f"Error sending email: {e}")


send_email(excel_file)