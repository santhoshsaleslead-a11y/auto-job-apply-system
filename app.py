from flask import Flask, render_template, request, redirect, send_file
import pandas as pd
import os
import smtplib
from email.mime.text import MIMEText

app = Flask(__name__)

DATA_FILE = "jobs.xlsx"

# Create file if not exists
if not os.path.exists(DATA_FILE):
    df = pd.DataFrame(columns=["Job Title", "Company", "Location", "Email", "Status"])
    df.to_excel(DATA_FILE, index=False)

@app.route('/')
def home():
    df = pd.read_excel(DATA_FILE)
    return render_template("index.html", jobs=df.to_dict(orient="records"))

@app.route('/add_job', methods=['POST'])
def add_job():
    title = request.form['title']
    company = request.form['company']
    location = request.form['location']
    email = request.form['email']

    df = pd.read_excel(DATA_FILE)
    df.loc[len(df)] = [title, company, location, email, "Pending"]
    df.to_excel(DATA_FILE, index=False)

    return redirect('/')

@app.route('/send_mail/<email>')
def send_mail(email):

    sender_email = os.getenv("EMAIL")
    password = os.getenv("PASSWORD")

    msg = MIMEText("""
Dear Hiring Manager,

I am very interested in the Python Full Stack Developer role.
Please find my resume attached.

Regards,
Candidate
""")

    msg['Subject'] = "Application for Python Developer Role"
    msg['From'] = sender_email
    msg['To'] = email

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, password)
        server.send_message(msg)
        server.quit()
    except:
        pass

    return redirect('/')

@app.route('/download')
def dow

