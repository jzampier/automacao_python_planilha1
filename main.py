from icecream import ic
import os
from datetime import datetime
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText
from dotenv import load_dotenv


# Create consolidated worksheet
def create_worksheet(folder: str) -> str:
    # path and list of files
    files = os.listdir(folder)
    # Create an empty dataframe
    consolidated_worksheet = pd.DataFrame()
    # Loop through files adding sales data to the dataframe
    for file_name in files:
        full_path = os.path.join(folder, file_name)
        try:
            sales_table = pd.read_csv(full_path)
            consolidated_worksheet = pd.concat([consolidated_worksheet, sales_table])
        except Exception as e:
            ic(f"Error reading file {full_path}: {e}")
    # Sort the dataframe
    consolidated_worksheet = consolidated_worksheet.sort_values(by='first_name')
    consolidated_worksheet = consolidated_worksheet.reset_index(drop=True)
    ic(consolidated_worksheet)

    # Save consolidated worksheet in an Excel file
    consolidated_worksheet.to_excel('Sales.xlsx', index=False)
    return 'Sales.xlsx'


# Email routine
def send_email(attachment_file: str):
    # environment variables
    load_dotenv()
    SENDER = os.getenv('SENDER')
    PASSWORD = os.getenv('PASSWORD')
    RECEIVER = os.getenv('RECEIVER')
    # Prepare the email (message + attachment)
    msg = MIMEMultipart()
    msg['From'] = SENDER
    msg['To'] = RECEIVER
    today_date: str = datetime.today().strftime('%d/%m/%Y')
    msg['Subject'] = f'Relatório de vendas - {today_date}'
    body = f"""<h3>Relatório de vendas do dia {today_date}</h3>
                    <p>Prezado(a) gerente do setor de vendas.</p>
                    <p></p>
                    <p></p>
                    <p>Segue anexado o relatório diário de vendas com os dados atualizados na 
                    data de hoje ({today_date}).</p>
                    <p>Caso tenha algum problema, favor nos avisar.</p>
                    <p></p>
                    <p></p>
                    <p>Cordialmente,</p>
                    <p>Equipe de análise de dados</p>"""
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


excel_file: str = create_worksheet('bases')
send_email(excel_file)
