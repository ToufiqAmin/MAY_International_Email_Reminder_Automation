import pandas as pd
import datetime
import schedule
import time
import smtplib
import logging
import re
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv

# --- Load credentials securely ---
load_dotenv()
SENDER_EMAIL = os.getenv("EMAIL_ADDRESS")
SENDER_PASSWORD = os.getenv("EMAIL_PASSWORD")

# --- Setup logging ---
logging.basicConfig(
    filename="reminder.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# --- Email format validator ---
def is_valid_email(email):
    return re.match(r"[^@]+@[^@]+\.[^@]+", email)

# --- Email sender ---
def send_email(to_email, subject, body):
    print(f"Loaded sender email: {SENDER_EMAIL}")
    print(f"Loaded reciver email: {to_email}")
    if not is_valid_email(to_email):
        logging.warning(f"Invalid email skipped: {to_email}")
        return

    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.send_message(msg)
        logging.info(f"Email sent to {to_email} | Subject: {subject}")
    except Exception as e:
        logging.error(f"Failed to send email to {to_email}: {e}")

# --- Reminder logic ---
def check_and_send_reminders():
    print("Checking for upcoming events...")
    try:
        df = pd.read_excel("events.xlsx", parse_dates=["event_date"])
    except Exception as e:
        logging.error(f"Failed to read events.xlsx: {e}")
        print(f"Failed to read events.xlsx: {e}")
        return

    print("DataFrame: \n",df)

    today = pd.Timestamp(datetime.date.today())
    print(f"Today's date: {today.date()}")
    for _, row in df.iterrows():
        if pd.isnull(row["event_date"]) or not isinstance(row["event_date"], pd.Timestamp):
            logging.warning(f"Invalid event date skipped: {row}")
            print(f"Invalid event date skipped: {row}")
            print("Skipping invalid event date: {today.date()}")
            continue

        days_until = (row["event_date"] - today).days
        if days_until in [7, 2]:
            subject = f"Reminder: {row['event_name']} in {days_until} days"
            body = f"""Dear user,

This is a reminder that '{row['event_name']}' is scheduled for {row['event_date'].date()}.

Regards,
Your M.I.E.R.A.(MAY INTERNATIONAL EMAIL REMINDER AUTOMATION) Bot"""
            send_email(row["email"], subject, body)

# For, automation purposes, uncomment the following lines to enable scheduling.
#  --- Schedule daily check ---
# schedule.every().day.at("09:00").do(check_and_send_reminders)
# schedule.every().day.at("16:00").do(check_and_send_reminders)

# try:
#     while True:
#         if os.path.exists("stop.flag"):
#             logging.info("Automation stopped by flag.")
#             break
#         schedule.run_pending()
#         time.sleep(60)
# except KeyboardInterrupt:
#     logging.info("Automation interrupted manually via KeyboardInterrupt.")
#   aaaa  print("Automation stopped manually.")

# For manual testing, uncomment the following line to run the check once.
if __name__ == "__main__":
    check_and_send_reminders()