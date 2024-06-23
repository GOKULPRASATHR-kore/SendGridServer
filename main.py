import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import csv
import json
from io import StringIO
from flask import Flask, request, jsonify
import os
import json
from botocore.exceptions import NoCredentialsError, PartialCredentialsError
from waitress import serve
import html


with open('config.json', 'r') as config_file:
    config = json.load(config_file)
app = Flask(__name__)
smtp_server = config['SMTP_SERVER']
smtp_port = config['SMTP_PORT']
smtp_username = config['SMTP_USERNAME']
smtp_password = config['SMTP_PASSWORD']
#bucket_name = config['BUCKET_NAME']


def success(to_email_user,from_email, to_email, subject, body, invoice_data_array,invoice_not_found_array,IsAllow_Mail_to_user):
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
    if(to_email=="kore@cassinfo.com"):
        message['Subject'] = to_email_user
    else:
        message['Subject'] = subject
  
    body = body.replace("CAUTION: EXTERNAL EMAIL. ", "")
    body = html.unescape(body)
    print(body)
    phrase_one = "For more detailed information, review the attachment of this email, if you have any questions remember we are always here to help."
    phrase_two = "Also we can't able to process some of your invoices, please bear with us and a representative from our side will contact you sooner"
    if(to_email=="kore@cassinfo.com"):
        if(len(invoice_not_found_array)==0):
            html_content = f"""
                    <body>
                        <p>{body}</p>
                        <p>{phrase_one}</p>
                        <h4><p>For future inquiries and to get immediate answers, please utilize <a href="https://my.cassport.com" target="_blank">CassPort</a> as your initial source of invoice status and payment information.</p></h4>
                        <b><p>Note: Success</p></b>
                        <table align="center" width="100%" bgcolor="white" class="email-container">
                        <tr>
                            <td>
                                <div style="text-align: center; width: 100%;">
                                    <div style="font-size:1px;line-height:1px;max-height:0px;max-width:0px;opacity:0;overflow:hidden;mso-hide:all;font-family: sans-serif;"></div>
                                    <table align="center" width="100%" color="color" bgcolor="white" class="email-container">
                                        <tr>
                                            <td style="padding: 10px 15px; text-align: left;">
                                                <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;"><strong>Carrier Support</strong></p>
                                                <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;">Cass Information Systems, Inc.</p>
                                                <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;">314-506-5959 08:00AM (CT) - 5:00PM (CT)</p>
                                                <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;"><a href="mailto:CarrierSupport@cassinfo.com" style="text-decoration: none; color: #0000ff; text-decoration: underline;">CarrierSupport@cassinfo.com</a></p>
                                                <p style="margin: 10px 0 0 0;"><img src="https://i.ibb.co/ZNNm8Qb/image007.jpg" alt="Footer Image" style="max-width: 50%; height: 28;"/></p>
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                                </td>
                        </tr>
                    </table>
                    </body>
                    """
        else:
            html_content = f"""
                    <body>
                        <p>{body}</p>
                        <p>{phrase_one}</p>
                        <p>{phrase_two}</p>
                        <h4><p>For future inquiries and to get immediate answers, please utilize <a href="https://my.cassport.com" target="_blank">CassPort</a> as your initial source of invoice status and payment information.</p></h4>
                        <b><p>Note: Success</p></b>
                        <table align="center" width="100%" bgcolor="white" class="email-container">
                        <tr>
                            <td>
                                <div style="text-align: center; width: 100%;">
                                    <div style="font-size:1px;line-height:1px;max-height:0px;max-width:0px;opacity:0;overflow:hidden;mso-hide:all;font-family: sans-serif;"></div>
                                    <table align="center" width="100%" color="color" bgcolor="white" class="email-container">
                                        <tr>
                                            <td style="padding: 10px 15px; text-align: left;">
                                                <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;"><strong>Carrier Support</strong></p>
                                                <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;">Cass Information Systems, Inc.</p>
                                                <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;">314-506-5959 08:00AM (CT) - 5:00PM (CT)</p>
                                                <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;"><a href="mailto:CarrierSupport@cassinfo.com" style="text-decoration: none; color: #0000ff; text-decoration: underline;">CarrierSupport@cassinfo.com</a></p>
                                                <p style="margin: 10px 0 0 0;"><img src="https://i.ibb.co/ZNNm8Qb/image007.jpg" alt="Footer Image" style="max-width: 50%; height: 28;"/></p>
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                                </td>
                        </tr>
                    </table>
                    </body>
                    """
    else:
        footer = """
                    <table align="center" width="100%" bgcolor="white" class="email-container">
                        <tr>
                            <td>
                                <div style="text-align: center; width: 100%;">
                                    <div style="font-size:1px;line-height:1px;max-height:0px;max-width:0px;opacity:0;overflow:hidden;mso-hide:all;font-family: san
s-serif;"></div>
<table align="center" width="100%" color="color" bgcolor="white" class="email-container">
                                        <tr>
                                            <td style="padding: 10px 15px; text-align: left;">
                                                <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;"><strong>Carrier Support</stron
g></p>
                                                <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;">Cass Information Systems, Inc.
</p>
                                                <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;">314-506-5959 08:00AM (CT) - 5:
00PM (CT)</p>
                                                <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;"><a href="mailto:CarrierSupport
@cassinfo.com" style="text-decoration: none; color: #0000ff; text-decoration: underline;">CarrierSupport@cassinfo.com</a></p>
                                                <p style="margin: 10px 0 0 0;"><img src="https://i.ibb.co/ZNNm8Qb/image007.jpg" alt="Footer Image" style="max-widt
h: 50%; height: 28;"/></p>
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                                </td>
                        </tr>
                    </table>
                 """
        if(len(invoice_not_found_array)==0):
            html_content = f"""
                            <body>
                                <p>{body}</p>
                                <p>{phrase_one}</p>
                                <h4><p>For future inquiries and to get immediate answers, please utilize  <a href="https://my.cassport.com" target="_blank">CassPort</a> as your initial source of invoice status and payment information.</p></h4>
                                {footer}
                            </body>
                    """
        else:
            html_content = f"""
                            <body>
                                <p>{body}</p>
                                <p>{phrase_one}</p>
                                <p>{phrase_two}</p>
                                <h4><p>For future inquiries and to get immediate answers, please utilize <a href="https://my.cassport.com" target="_blank">CassPort</a> as your initial source of invoice status and payment information.</p></h4>
                                {footer}
                            </body>
                    """
    message.attach(MIMEText(html_content, 'html'))
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

# #TO send the email to user {sucessfull_Scenario}
    if(IsAllow_Mail_to_user == 1):
        message = MIMEMultipart()
        message['From'] = from_email
        message['To'] = to_email_user
        if(to_email=="kore@cassinfo.com"):
            message['Subject'] = to_email_user
        else:
            message['Subject'] = subject
    
        body = body.replace("CAUTION: EXTERNAL EMAIL. ", "")
        body = html.unescape(body)
        print(body)
        phrase_one = "For more detailed information, review the attachment of this email, if you have any questions remember we are always here to help."
        phrase_two = "Also we can't able to process some of your invoices, please bear with us and a representative from our side will contact you sooner"
        if(to_email=="kore@cassinfo.com"):
            if(len(invoice_not_found_array)==0):
                html_content = f"""
                        <body>
                            <p>{body}</p>
                            <p>{phrase_one}</p>
                            <h4><p>For future inquiries and to get immediate answers, please utilize  <a href="https://my.cassport.com" target="_blank">CassPort</a> as your initial source of invoice status and payment information.</p></h4>
                            <b><p>Note: Success</p></b>
                            <table align="center" width="100%" bgcolor="white" class="email-container">
                            <tr>
                                <td>
                                    <div style="text-align: center; width: 100%;">
                                        <div style="font-size:1px;line-height:1px;max-height:0px;max-width:0px;opacity:0;overflow:hidden;mso-hide:all;font-family: sans-serif;"></div>
                                        <table align="center" width="100%" color="color" bgcolor="white" class="email-container">
                                            <tr>
                                                <td style="padding: 10px 15px; text-align: left;">
                                                    <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;"><strong>Carrier Support</strong></p>
                                                    <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;">Cass Information Systems, Inc.</p>
                                                    <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;">314-506-5959 08:00AM (CT) - 5:00PM (CT)</p>
                                                    <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;"><a href="mailto:CarrierSupport@cassinfo.com" style="text-decoration: none; color: #0000ff; text-decoration: underline;">CarrierSupport@cassinfo.com</a></p>
                                                    <p style="margin: 10px 0 0 0;"><img src="https://i.ibb.co/ZNNm8Qb/image007.jpg" alt="Footer Image" style="max-width: 50%; height: 28;"/></p>
                                                </td>
                                            </tr>
                                        </table>
                                    </div>
                                    </td>
                            </tr>
                        </table>
                        </body>
                        """
            else:
                html_content = f"""
                        <body>
                            <p>{body}</p>
                            <p>{phrase_one}</p>
                            <p>{phrase_two}</p>
                            <h4><p>For future inquiries and to get immediate answers, please utilize  <a href="https://my.cassport.com" target="_blank">CassPort</a> as your initial source of invoice status and payment information.</p></h4>
                            <b><p>Note: Success</p></b>
                            <table align="center" width="100%" bgcolor="white" class="email-container">
                            <tr>
                                <td>
                                    <div style="text-align: center; width: 100%;">
                                        <div style="font-size:1px;line-height:1px;max-height:0px;max-width:0px;opacity:0;overflow:hidden;mso-hide:all;font-family: sans-serif;"></div>
                                        <table align="center" width="100%" color="color" bgcolor="white" class="email-container">
                                            <tr>
                                                <td style="padding: 10px 15px; text-align: left;">
                                                    <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;"><strong>Carrier Support</strong></p>
                                                    <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;">Cass Information Systems, Inc.</p>
                                                    <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;">314-506-5959 08:00AM (CT) - 5:00PM (CT)</p>
                                                    <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;"><a href="mailto:CarrierSupport@cassinfo.com" style="text-decoration: none; color: #0000ff; text-decoration: underline;">CarrierSupport@cassinfo.com</a></p>
                                                    <p style="margin: 10px 0 0 0;"><img src="https://i.ibb.co/ZNNm8Qb/image007.jpg" alt="Footer Image" style="max-width: 50%; height: 28;"/></p>
                                                </td>
                                            </tr>
                                        </table>
                                    </div>
                                    </td>
                            </tr>
                        </table>
                        </body>
                        """
        else:
            footer = """
                        <table align="center" width="100%" bgcolor="white" class="email-container">
                            <tr>
                                <td>
                                    <div style="text-align: center; width: 100%;">
                                        <div style="font-size:1px;line-height:1px;max-height:0px;max-width:0px;opacity:0;overflow:hidden;mso-hide:all;font-family: san
    s-serif;"></div>
    <table align="center" width="100%" color="color" bgcolor="white" class="email-container">
                                            <tr>
                                                <td style="padding: 10px 15px; text-align: left;">
                                                    <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;"><strong>Carrier Support</stron
    g></p>
                                                    <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;">Cass Information Systems, Inc.
    </p>
                                                    <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;">314-506-5959 08:00AM (CT) - 5:
    00PM (CT)</p>
                                                    <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;"><a href="mailto:CarrierSupport
    @cassinfo.com" style="text-decoration: none; color: #0000ff; text-decoration: underline;">CarrierSupport@cassinfo.com</a></p>
                                                    <p style="margin: 10px 0 0 0;"><img src="https://i.ibb.co/ZNNm8Qb/image007.jpg" alt="Footer Image" style="max-widt
    h: 50%; height: 28;"/></p>
                                                </td>
                                            </tr>
                                        </table>
                                    </div>
                                    </td>
                            </tr>
                        </table>
                    """
            if(len(invoice_not_found_array)==0):
                html_content = f"""
                                <body>
                                    <p>{body}</p>
                                    <p>{phrase_one}</p>
                                    <h4><p>For future inquiries and to get immediate answers, please utilize  <a href="https://my.cassport.com" target="_blank">CassPort</a> as your initial source of invoice status and payment information.</p></h4>
                                    {footer}
                                </body>
                        """
            else:
                html_content = f"""
                                <body>
                                    <p>{body}</p>
                                    <p>{phrase_one}</p>
                                    <p>{phrase_two}</p>
                                    <h4><p>For future inquiries and to get immediate answers, please utilize  <a href="https://my.cassport.com" target="_blank">CassPort</a> as your initial source of invoice status and payment information.</p></h4>
                                    {footer}
                                </body>
                        """
        message.attach(MIMEText(html_content, 'html'))
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
            print(f'Email sent 2 {to_email_user}')
    
    else:
        print("NO Mail send to user")






def failure(to_email_user,from_email, to_email, subject, body, follow_up_message,database_message,IsAllow_Mail_to_user):
    is_allow = 1
    body = body.replace("CAUTION: EXTERNAL EMAIL. ", "")
    database_message = html.unescape(database_message)
    
    if to_email=="kore@cassinfo.com":
        message = MIMEMultipart()
        message['From'] = from_email
        message['To'] = to_email
        message['Subject'] = to_email_user
        html_content = f"""
                            <body>
                                <h3>Email Subject: {subject}</h3>
                                <p>{body}</p>
                                <h4><p>For future inquiries and to get immediate answers, please utilize  <a href="https://my.cassport.com" target="_blank">CassPort</a> as your initial source of invoice status and payment information.</p></h4>
                                <b><p>Note: Unsuccess- {database_message}</p></b>
                                <table align="center" width="100%" bgcolor="white" class="email-container">
                        <tr>
                            <td>
                                <div style="text-align: center; width: 100%;">
                                    <div style="font-size:1px;line-height:1px;max-height:0px;max-width:0px;opacity:0;overflow:hidden;mso-hide:all;font-family: sans-serif;"></div>
                                    <table align="center" width="100%" color="color" bgcolor="white" class="email-container">
                                        <tr>
                                            <td style="padding: 10px 15px; text-align: left;">
                                                <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;"><strong>Carrier Support</strong></p>
                                                <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;">Cass Information Systems, Inc.</p>
                                                <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;">314-506-5959 08:00AM (CT) - 5:00PM (CT)</p>
                                                <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;"><a href="mailto:CarrierSupport@cassinfo.com" style="text-decoration: none; color: #0000ff; text-decoration: underline;">CarrierSupport@cassinfo.com</a></p>
                                                <p style="margin: 10px 0 0 0;"><img src="https://i.ibb.co/ZNNm8Qb/image007.jpg" alt="Footer Image" style="max-width: 50%; height: 28;"/></p>
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                                </td>
                        </tr>
                    </table>
                            </body>
                        """
        message.attach(MIMEText(html_content, 'html'))
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.send_message(message)
            print("Email to outlook")
            print(f'Unsuccess: Email sent 2 {to_email}')
        

    if(IsAllow_Mail_to_user == 1):
        footer = """
                        <table align="center" width="100%" bgcolor="white" class="email-container">
                            <tr>
                                <td>
                                    <div style="text-align: center; width: 100%;">
                                        <div style="font-size:1px;line-height:1px;max-height:0px;max-width:0px;opacity:0;overflow:hidden;mso-hide:all;font-family: sans-serif;"></div>
                                        <table align="center" width="100%" color="color" bgcolor="white" class="email-container">
                                            <tr>
                                                <td style="padding: 10px 15px; text-align: left;">
                                                    <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;"><strong>Carrier Support</strong></p>
                                                    <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;">Cass Information Systems, Inc.</p>
                                                    <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;">314-506-5959 08:00AM (CT) - 5:00PM (CT)</p>
                                                    <p style="margin: 0; font-size: 12px; font-family: Arial, sans-serif; color: #000;"><a href="mailto:CarrierSupport@cassinfo.com" style="text-decoration: none; color: #0000ff; text-decoration: underline;">CarrierSupport@cassinfo.com</a></p>
                                                    <p style="margin: 10px 0 0 0;"><img src="https://i.ibb.co/ZNNm8Qb/image007.jpg" alt="Footer Image" style="max-width: 50%; height: 28;"/></p>
                                             </td>
                                            </tr>
                                        </table>
                                    </div>
                                </td>
                            </tr>
                        </table>
                    """
        message = MIMEMultipart()
        message['From'] = from_email
        message['To'] = to_email_user
        message['Subject'] = subject
        body = follow_up_message
        html_content = f"""
                                <body>
                                    <p>{follow_up_message}</p>        
    <h4><p>For future inquiries and to get immediate answers, please utilize  <a href="https://my.cassport.com" target="_blank">CassPort</a> as your initial source of invoice status and payment information.</p></h4>
    {footer}                           
     </body>
                                """
        message.attach(MIMEText(html_content, 'html'))
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.send_message(message)
            print(f'Email sent 3 {to_email}')
    else:
        print("NO Mail send to user")
    
@app.route('/send_email', methods=['POST'])
def process_data():
    data = request.json
    _from = data.get('From')
    _to = data.get('To')
    subject = data.get('Subject')
    body = data.get('Body')
    response = data.get('Response')
    invoice_data_array = data.get('InvoiceDataArray')
    invoice_not_found_array = data.get('InvoiceNotFoundArray')
    database_message = data.get('DatabaseMessage')
    follow_up_message = data.get('FollowUpMessage')
    IsAllow_Mail_to_user = data.get('IsAllowMailtoUser')
    source_from = config['OUTLOOK_INBOX']
    if invoice_data_array:
        to_email_user = _from
        # success(to_email_user,source_from, to_email_user, subject, response, invoice_data_array,invoice_not_found_array)
        to_email_outlook = config['OUTLOOK_INBOX']
        outlook_subject = subject
        success(to_email_user,source_from, to_email_outlook, outlook_subject, response, invoice_data_array,invoice_not_found_array,IsAllow_Mail_to_user)
    else:
        to_email_user = _from
        # failure(to_email_user,source_from, to_email_user, subject, body, follow_up_message,database_message)
        to_email_outlook = config['OUTLOOK_INBOX']
        outlook_subject = subject
        failure(to_email_user,source_from, to_email_outlook, outlook_subject, body, follow_up_message,database_message,IsAllow_Mail_to_user)
    return jsonify({'message': 'Email sent successfully'})
@app.route('/', methods=['GET'])
def home():
    return jsonify({'message':"Server Running Successfully"})
if __name__ == '__main__':
    serve(app, host='0.0.0.0', port=5000)
    #app.run(host='0.0.0.0', port=5000)
