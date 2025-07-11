import pandas as pd
from dotenv import load_dotenv
from openpyxl import load_workbook
from datetime import datetime
import smtplib
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.mime.text import MIMEText
import os

# === НАСТРОЙКИ ===
INPUT_FILE = 'test_sales.xlsx'
OUTPUT_FILE = 'filtered_sales.xlsx'

load_dotenv()

EMAIL_FROM = os.getenv('EMAIL_FROM')
EMAIL_TO = os.getenv('EMAIL_TO')
EMAIL_SUBJECT = os.getenv('EMAIL_SUBJECT')
SMTP_SERVER = os.getenv('SMTP_SERVER')
SMTP_PORT = int(os.getenv('SMTP_PORT'))
SMTP_USER = os.getenv('SMTP_USER')
SMTP_PASSWORD = os.getenv('SMTP_PASSWORD')
DATE_COLUMN = os.getenv('DATE_COLUMN')
DATE_FORMAT = os.getenv('DATE_FORMAT')

# === ФУНКЦИЯ: ФИЛЬТРАЦИЯ ===
def filter_data_by_date(input_file, output_file, start_date, end_date):
    df = pd.read_excel(input_file, engine='openpyxl')
    df[DATE_COLUMN] = pd.to_datetime(df[DATE_COLUMN], format=DATE_FORMAT)

    filtered_df = df[(df[DATE_COLUMN] >= start_date) & (df[DATE_COLUMN] <= end_date)]
    filtered_df.to_excel(output_file, index=False)
    print(f"[✓] Сохранено: {output_file}")


# === ФУНКЦИЯ: ОТПРАВКА EMAIL ===
def send_email_with_attachment(from_addr, to_addr, subject, body, attachment_path):
    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['To'] = to_addr
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    with open(attachment_path, 'rb') as file:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
        msg.attach(part)

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.send_message(msg)
        print(f"[✓] Email sent to {to_addr}")


# === ОСНОВНОЙ КОД ===
if __name__ == '__main__':
    # Задаём период фильтрации – текущий месяц
    today = datetime.today()
    start = datetime(today.year, today.month, 1)
    end = datetime(today.year, today.month + 1, 1) if today.month < 12 else datetime(today.year + 1, 1, 1)

    print(f"[→] Filtering by period: {start.date()} — {end.date()}")

    filter_data_by_date(INPUT_FILE, OUTPUT_FILE, start, end)
    send_email_with_attachment(
        EMAIL_FROM,
        EMAIL_TO,
        EMAIL_SUBJECT,
        f"Attached is the report for the period: {start.date()} — {end.date()}",
        OUTPUT_FILE
    )
