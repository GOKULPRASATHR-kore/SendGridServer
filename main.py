import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import csv
import json
from io import StringIO
from flask import Flask, request, jsonify
import boto3
import os
import json
from botocore.exceptions import NoCredentialsError, PartialCredentialsError
from waitress import serve

with open('config.json', 'r') as config_file:
    config = json.load(config_file)

app = Flask(__name__)

smtp_server = config['SMTP_SERVER']
smtp_port = config['SMTP_PORT']
smtp_username = config['SMTP_USERNAME']
smtp_password = config['SMTP_PASSWORD']
bucket_name = config['BUCKET_NAME']

s3 = boto3.client(
                    's3',
                    aws_access_key_id=config['AWS_ACCESS_KEY_ID'],
                    aws_secret_access_key=config['AWS_SECRET_ACCESS_KEY']
                )


def delete_file_from_s3(bucket_name, file_name):
    try:
        response = s3.delete_object(Bucket=bucket_name, Key=file_name)
        return response
    
    except NoCredentialsError:
        print("Credentials not available.")

    except PartialCredentialsError:
        print("Incomplete credentials provided.")

    except Exception as e:
        print(f"An error occurred: {e}")



def success(from_email, to_email, subject, body, invoice_data_array):
    try:
        csv_data = StringIO()
        csv_writer = csv.DictWriter(csv_data, fieldnames=invoice_data_array[0].keys())
        csv_writer.writeheader()
        csv_writer.writerows(invoice_data_array)
    except:
        print("Error creating CSV file")

    message = MIMEMultipart()
    message['From'] = from_email
    message['To'] = to_email
    message['Subject'] = subject

    message.attach(MIMEText(body, 'html'))

    try:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(csv_data.getvalue().encode())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename='invoice_data.csv')
        message.attach(part)
    except:
        print("Error creating CSV file")

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.send_message(message)
        print(f'Email sent 1 {to_email}')

def failure(from_email, to_email, subject, body, follow_up_message,database_message,filename):

    is_allow = 1

    if to_email=="kore@cassinfo.com":
        message = MIMEMultipart()
        message['From'] = from_email
        message['To'] = to_email
        message['Subject'] = subject

        body = body + " " + "NOTE: " + database_message
        message.attach(MIMEText(body, 'html'))

        if len(filename)!=0:
            directory = "attachments"

            if not os.path.exists(directory):
                os.makedirs(directory)
            try:
                s3.download_file(bucket_name, filename, "attachments/"+filename)

                with open("attachments/"+filename, 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', 'attachment', filename=filename)
                    message.attach(part)
            
            except:
                is_allow = 0
                print(" No files are found")
            
            

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.send_message(message)
            print(f'Email sent 2 {to_email}')

            try:
                if len(filename)!=0:
                    if is_allow == 1:
                        delete_file_from_s3(bucket_name, filename)
                        os.remove("attachments/"+filename)
            except:
                pass
            
    # else:
    #     message = MIMEMultipart()
    #     message['From'] = from_email
    #     message['To'] = to_email
    #     message['Subject'] = subject

    #     body = follow_up_message
    #     message.attach(MIMEText(body, 'html'))
        
    #     with smtplib.SMTP(smtp_server, smtp_port) as server:
    #         server.starttls()
    #         server.login(smtp_username, smtp_password)
    #         server.send_message(message)
    #         print(f'Email sent 3 {to_email}')

@app.route('/send_email', methods=['POST'])
def process_data():
    data = request.json

    _from = data.get('From')
    _to = data.get('To')
    subject = data.get('Subject')
    body = data.get('Body')
    filename = data.get('Filename')
    response = data.get('Response')
    invoice_data_array = data.get('InvoiceDataArray')
    invoice_not_found_array = data.get('InvoiceNotFoundArray')
    database_message = data.get('DatabaseMessage')
    follow_up_message = data.get('FollowUpMessage')
    source_from = config['OUTLOOK_INBOX']

    
    if invoice_data_array:
        #to_email_user = _from
        #success(source_from, to_email_user, subject, response, invoice_data_array)

        to_email_outlook = config['OUTLOOK_INBOX']
        outlook_subject = "Success "+""+ subject
        success(source_from, to_email_outlook, outlook_subject, response, invoice_data_array)

    else:
        #to_email_user = _from
        #failure(source_from, to_email_user, subject, body, follow_up_message,database_message,filename)

        to_email_outlook = config['OUTLOOK_INBOX']
        outlook_subject = "Unsuccess "+""+ subject
        failure(source_from, to_email_outlook, outlook_subject, body, follow_up_message,database_message,filename)

    return jsonify({'message': 'Email sent successfully'})

@app.route('/', methods=['GET'])
def home():
    return jsonify({'message':"Server Running Successfully"})

if __name__ == '__main__':
    #serve(app, host='0.0.0.0', port=5000)
    app.run()