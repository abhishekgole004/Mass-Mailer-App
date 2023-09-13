import json
import os
import boto3
import openpyxl
from io import BytesIO
ses = boto3.client('ses', region_name='us-east-1')

def lambda_handler(event, context):
    # TODO implement
    
    bucket = event['Records'][0]['s3']['bucket']['name']
    key = event['Records'][0]['s3']['object']['key']
    print("event", event)
    print("key", key)
    # Download the XLSX file from S3
    s3 = boto3.client('s3')
    response = s3.get_object(Bucket=bucket, Key=key)
    xlsx_data = response['Body'].read()

    # Parse the XLSX data
    xlsx_file = BytesIO(xlsx_data)
    workbook = openpyxl.load_workbook(xlsx_file)

    # Assuming the first row contains headers, you can map the columns by their names
    sheet = workbook.active
    headers = [cell.value for cell in sheet[1]]
    print(headers)
    for row in sheet.iter_rows(min_row=2, values_only=True):
        row_data = dict(zip(headers, row))
        to_email = row_data['mails']
        email_subject = row_data['Subject']
        email_body = row_data['body']

        # Send the email
        send_email(to_email, email_subject, email_body)

def send_email(to_email, email_subject, email_body):
    sender_email = 'abhisheklale1812@gmail.com'

    try:
        # Send the email using SES
        response = ses.send_email(
            Source=sender_email,
            Destination={
                'ToAddresses': [to_email]
            },
            Message={
                'Subject': {
                    'Data': email_subject
                },
                'Body': {
                    'Text': {
                        'Data': email_body
                    }
                }
            }
        )
        print(f"Email sent to {to_email} with subject: {email_subject}")
    except Exception as e:
        print(f"Error sending email to {to_email}: {str(e)}")

    
