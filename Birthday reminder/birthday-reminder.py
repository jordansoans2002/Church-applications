import os
from dotenv import load_dotenv
import datetime
import csv
import pandas as pd
import ssl, smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

load_dotenv(dotenv_path="/opt/church/.env")


# Step 1: Load birthday data from a CSV file
def load_birthdays(file_path, sheet):
    df = pd.read_excel(
        file_path,
        sheet_name=sheet,
        header=None,
        usecols=[0, 1],
        names=['Name', 'DOB']
    )
    # Drop rows where name or dob is missing
    df = df.dropna(subset=['Name', 'DOB'])
    # Ensure DOB is string, remove ordinal suffixes, then parse
    df['DOB'] = (
        df['DOB']
        .astype(str)
        .str.replace(r'(\d+)(st|nd|rd|th)', r'\1', regex=True)
    )
    df['DOB'] = pd.to_datetime(df['DOB'], dayfirst=True, errors='coerce').dt.date
    # Drop any rows where parsing failed
    df = df.dropna(subset=['DOB'])

    # Return list of dicts: { 'Name': str, 'DOB': date }
    return df.to_dict(orient='records')


# Step 2: Find today's birthdays
def get_todays_birthdays(birthdays):
    today = datetime.date.today()
    return [
        p for p in birthdays
        if p['DOB'].month == today.month and p['DOB'].day == today.day
    ]


def get_weekly_birthdays(birthdays):
    today = datetime.date.today()

    # Only check if today is Sunday
    if today.weekday() != int(os.getenv("SUMMARY_DAY_OF_WEEK", "6")):
        return []

    # Get Monday and Sunday for this week
    monday = today - datetime.timedelta(days=today.weekday())  # Monday
    sunday = monday + datetime.timedelta(days=6)               # Sunday

    week_birthdays = []
    for person in birthdays:
        birthday_this_year = person['DOB'].replace(year=today.year)
        if monday <= birthday_this_year <= sunday:
            week_birthdays.append(person)

    return week_birthdays


# Step 3: Send an email
def send_email(message):
    SENDER_EMAIL = os.getenv("EMAIL_USER")
    receiver_emails = os.getenv("EMAIL_RECIPIENTS").split(",")
    receiver_email = "nathanielsimons176@gmail.com"
    subject = "Daily Birthday Reminder"

    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = ", ".join(receiver_emails)
    msg["Subject"] = subject
    msg.attach(MIMEText(message, "plain"))

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as server:
        server.ehlo()
        server.login(SENDER_EMAIL, os.getenv("EMAIL_PASS"))
        server.sendmail(SENDER_EMAIL, receiver_emails, msg.as_string())

# Run the script
if __name__ == "__main__":
    path = os.path.join(os.getcwd(), 'YOUTH FELLOWSHIP THANE (2019-2023).xlsx')
    birthdays = load_birthdays(path, sheet="Youth Members List")
    today_birthdays = get_todays_birthdays(birthdays)
    if today_birthdays:
        message = "🎈 Today's Birthdays:\n\n" + "\n".join(
            [f"Happy birthday {person['Name']} 🥳🥳🥳" for person in today_birthdays]
        )
        send_email(message)

    weekly_birthdays = get_weekly_birthdays(birthdays)
    if weekly_birthdays:
        message = "🎉 Birthdays This Week:\n\n" + "\n".join(
            [f"{person['Name']} - {person['DOB'].strftime('%d %b')}" for person in weekly_birthdays]
        )
        send_email(message)