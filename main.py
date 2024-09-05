import pandas as pd
from smtplib import SMTP
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import json

class EmailSender:
    def __init__(self, sender_email, sender_password):
        self.sender_email = sender_email
        self.sender_password = sender_password
        self.group = {}

    def add_recipient(self, name, email):
        self.group[name] = email

    def load_recipients_from_excel(self, file_path, sheet_name="Sheet1"):
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            for index, row in df.iterrows():
                self.add_recipient(row['Name'], row['Email'])
            print(f"Loaded {len(df)} recipients from {file_path}")
        except Exception as e:
            print(f"Failed to load recipients from Excel. Error: {e}")

    def send_email(self, receiver_email, subject, body_html):
        msg = MIMEMultipart()
        msg['From'] = self.sender_email
        msg['To'] = receiver_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body_html, 'html'))  
        
        try:
            with SMTP('smtp.gmail.com', 587) as server:
                server.starttls()
                server.login(self.sender_email, self.sender_password)
                server.sendmail(self.sender_email, receiver_email, msg.as_string())
            print(f"Mail sent successfully to {receiver_email}!")
        except Exception as e:
            print(f"Failed to send mail to {receiver_email}. Error: {e}")

    def send_survey(self, subject, form_link, custom_message=""):
        for name, email in self.group.items():
            body_html = f"""
            <!DOCTYPE html>
            <html>
            <head>
              <meta http-equiv="CONTENT-TYPE" content="text/html; charset=UTF-8">
              <title>Survey Invitation</title>
              <style>
                body {{
                  font-family: Garamond, serif;
                  background-color: #f4f4f4;
                  margin: 0;
                  padding: 0;
                }}

                .form {{
                  width: 90%;
                  max-width: 600px;
                  background-color: #CEE6F3;
                  margin: 50px auto;
                  padding: 20px;
                  border-radius: 10px;
                  box-shadow: 0px 0px 10px rgba(0, 0, 0, 0.1);
                  text-align: center;
                }}

                .form h1 {{
                  font-size: 24px;
                  color: #333333;
                  margin-bottom: 20px;
                }}

                .form p {{
                  font-size: 16px;
                  color: #666666;
                  margin-bottom: 20px;
                }}

                .form a {{
                  display: inline-block;
                  padding: 10px 20px;
                  font-size: 18px;
                  color: #ffffff;
                  background-color: #5a8d78;
                  border-radius: 5px;
                  text-decoration: none;
                  margin-top: 20px;
                }}

                .form a:hover {{
                  background-color: #49705f;
                }}
              </style>
            </head>
            <body>
              
              <div class="form">
                <h1>Your Insights Matter to Our Research</h1>
                  
                <p>Dear {name},</p>
                <p>{custom_message}</p>
                <p>Please fill out the following survey by clicking the button below:</p>
                  
                <a href="{form_link}">Start Survey</a>
                  
                <p>Thank you for your time.</p>
              </div>
              
            </body>
            </html>
            """
            self.send_email(email, subject, body_html)

if __name__ == "__main__":
    with open('config.json') as config_file:
        config = json.load(config_file)
    	
    sender_email = config['email']
    sender_password = config['password']
    email_sender = EmailSender(sender_email, sender_password)
    	
    email_sender.load_recipients_from_excel("Recipient1.xlsx", "Sheet1")
    subject = "We Value Your Feedback - Please Fill Out Our Survey"
    form_link = "https://forms.gle/6wznQyH3tag2zLRv8"      
    custom_message = "We are conducting a study to gather insights on students' experiences and perspectives. Your participation in this survey will significantly contribute to our research."
    email_sender.send_survey(subject, form_link, custom_message)
