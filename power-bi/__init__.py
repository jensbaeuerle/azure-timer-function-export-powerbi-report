import datetime
import logging

import azure.functions as func

import requests
import json

import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText




def main(mytimer: func.TimerRequest) -> None:
    utc_timestamp = datetime.datetime.utcnow().replace(
        tzinfo=datetime.timezone.utc).isoformat()

    
    reportID = "<Report-ID>"
    # data for the body
    data = "{format: 'PDF'}"

    # header
    headers = {
    'Authorization': '<Bearer Token>',
    'Content-Type': 'application/json'
    }
    ## Reports - Export To File
    def exportToFile(data, headers, reportID):
        # defining the api-endpoint  
        url = "https://api.powerbi.com/v1.0/myorg/reports/{}/ExportTo".format(reportID)

        # sending post request and saving response as response object 
        response = requests.request("POST", url, headers=headers, data = data)

        # print(response.text.encode('utf8'))
        print(response.text)
        y = json.loads(response.text)['id']
        return y

    # Reports - Get Export To File Status
    def getFileStatus(headers, reportID, exportID):
        # defining the api-endpoint  
        url = "https://api.powerbi.com/v1.0/myorg/reports/{}/exports/{}".format(reportID, exportID)
        # sending post request and saving response as response object 
        response = requests.request("GET", url, headers=headers)

        print(response.text)
        y = json.loads(response.text)['percentComplete']
        print(y)
        return y

    # Reports - Get File Of Export To File
    def getFile(headers_file, reportID, exportID):
        # set up the SMTP server
        s = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)
        s.starttls()
        # s.login('Jens.Baeuerle@outlook.de', input("Type your password and press enter:"))
        s.login('<E-Mail>', '<Password>')
        

        subject = "An email with attachment from Python"
        body = "This is an email with attachment sent from Python"
        sender_email = "<E-Mail>"
        receiver_email = "<E-Mail>"

        # Create a multipart message and set headers
        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = receiver_email
        message["Subject"] = subject
        message["Bcc"] = receiver_email  # Recommended for mass emails

        # defining the api-endpoint  
        url = "https://api.powerbi.com/v1.0/myorg/reports/{}/exports/{}/file".format(reportID, exportID)
        # sending post request and saving response as response object 
        response = requests.request("GET", url, headers=headers)

        # print(response.text.encode('utf8'))
        # with open('C:/Users/jebaeuer/Desktop/customer-repo/powerbi.pdf', 'wb') as f:
        #     f.write(response.content)
        # print("File is saved!")

        # Add body to email
        message.attach(MIMEText(body, "plain"))

        filename = "document.pdf"  # In same directory as script

        # Open PDF file in binary mode

        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(response.content)

        # Encode file in ASCII characters to send by email    
        encoders.encode_base64(part)

        # Add header as key/value pair to attachment part
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename}",
        )

        # Add attachment to message and convert message to string
        message.attach(part)
        text = message.as_string()

        s.sendmail(sender_email, receiver_email, text)
        s.quit()

    # if mytimer.past_due:
    #     logging.info('The timer is past due!')

    exportID = exportToFile(data, headers, reportID)
    # getFileStatus(headers, reportID, exportID)
    # getFile(headers, reportID, 'MS9CbG9iSWRWMi1mNjA3MGY5My02NmZjLTQwZjgtYmQwMS0zZjdhNTEyYmRjMzBPSVZnWkdydE5XM0ROMDBmWHkzVDhkSTJja09USEpxb1RRZXNYUldVVElRPS4=')

    # percentage = 1
    # while percentage < 100:
    #     print(percentage)
    #     percentage = percentage + 1
    # else:
    #     print("Wir sind durch")

    
    percentage = 1
    while percentage <= 99:
        print(percentage)
        percentage = getFileStatus(headers, reportID, exportID)
    else:
        getFile(headers, reportID, exportID)

    logging.info('Python timer trigger function ran at %s', utc_timestamp)
    logging.info("Hallo mein Name ist Jens")
