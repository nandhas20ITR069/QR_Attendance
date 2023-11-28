from flask import Flask, render_template, request
import cv2
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from os.path import basename
from email import encoders
import qrcode
import pandas as pd
from datetime import datetime

app = Flask(__name__)

sender = 'nandhas.20it@kongu.edu'
passwd = 'jxnxegepslxvkshe'

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        session = request.form['session']
        if 'invite' in request.form:
            invite(session)
        elif 'scan' in request.form:
            accept(session)
    return render_template('index.html')

def send_mail(receiver, subject, body, data, name, email, roll_no, year, section, timestamp):
    msg = MIMEMultipart()
    msg['From'] = sender
    msg['To'] = receiver
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    qrcode.make(data).save('data.png')
    filename = attachment = 'data.png'
    with open(filename, 'rb') as f:
        attachment = MIMEBase(*'image/png'.split('/'), filename=basename(filename))
        attachment.set_payload(f.read())
        attachment['Content-Disposition'] = 'attachment; filename="{}"'.format(basename(filename))
        encoders.encode_base64(attachment)
        msg.attach(attachment)

    # Add additional information to the email body
    additional_info = f"Name: {name}\nEmail: {email}\nRoll No: {roll_no}\nYear: {year}\nSection: {section}\nTimestamp: {timestamp}"
    msg.attach(MIMEText(additional_info, 'plain'))

    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=ssl.create_default_context()) as smtp:
        smtp.login(sender, passwd)
        smtp.sendmail(sender, receiver, msg.as_string())

def invite(session):
    header = None
    df = pd.read_excel('student.xlsx')  # Load Excel file
    header = list(df.columns)
    if session in header:
        print('Session already completed')
        return

    for _, row in df.iterrows():
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # Timestamp for sending QR code
        send_mail(row['email'], f'You are invited to attend {session}', '', row['email'],
                  row['Name'], row['email'], row['Roll No'], row['Year'], row['Section'], timestamp)

    header.append(session)
    df[session] = 'NO'

    df.to_excel('student.xlsx', index=False)  # Save changes back to Excel

def accept(session):
    emails = set()
    cam = cv2.VideoCapture(0)
    detector = cv2.QRCodeDetector()
    while True:
        _, img = cam.read()
        data, bbox, _ = detector.detectAndDecode(img)
        if data:
            print("QR Code detected-->", data)
            emails.add(data)
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # Timestamp for scanning QR code
            mark_attendance(data, session, timestamp)  # Mark attendance based on QR code data and timestamp
        cv2.imshow("img", img)
        if cv2.waitKey(1) == ord("q"):
            break
    cam.release()
    cv2.destroyAllWindows()

def mark_attendance(email, session, timestamp):
    df = pd.read_excel('student.xlsx')  # Load Excel file

    for index, row in df.iterrows():
        if row['email'] == email:
            df.at[index, session] = 'YES'  # Mark attendance for the corresponding email in the session column
            df.at[index, f'{session}_Timestamp'] = timestamp  # Add timestamp in a separate column for the session

    df.to_excel('student.xlsx', index=False)  # Save changes back to Excel



if __name__ == '__main__':
    #session = input('SESSION: ')
    #invite(session)
    #accept(session)
    app.run(debug=True)
