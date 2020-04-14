# Author: Eduardo Marossi
# Version: 1.0.0
import argparse
import json
import logging
import os
from argparse import ArgumentError
import pickle
from urllib.parse import urlparse
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from mail_send import email_providers, EmailBackend, EmailMessage


def load_mail_credentials(path):
    with open(path) as f:
        data = json.loads(f.read())

    required = ['provider', 'username', 'app_password', 'message']
    for k in required:
        if k not in data:
            raise ArgumentError('Missing required information in mail_credentials json file.')

    return data


def read_sheet(credentials_path, sheet_id, sheet_range):
    # Source: https://developers.google.com/sheets/api/quickstart/python
    creds = None
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_id,
                                range=sheet_range).execute()
    values = result.get('values', [])

    return values

def get_header_lines_number(header_lines):
    if '-' not in header_lines:
        line_start = line_end = int(header_lines) - 1
    else:
        line_start = int(header_lines.split('-')[0]) - 1
        line_end = int(header_lines.split('-')[1]) - 1

    return line_start, line_end


def get_header_columns(data, header_lines):
    header_names = {}
    line_start, line_end = get_header_lines_number(header_lines)

    header_size = 0
    for l in range(line_start, line_end+1):
        if len(data[l]) > header_size:
            header_size = len(data[l])

    print(line_start)
    print(line_end)
    for l in range(line_start, line_end+1):
        print(data[l])
        if len(data[l]) != header_size:
            continue

        for c in range(0, len(data[l])):
            if c not in header_names:
                header_names[c] = data[l][c]
            else:
                header_names[c] = header_names[c] + ' ' + data[l][c]
            header_names[c] = header_names[c].strip()
    return header_names

def prepare_message(row_data, mail_credentials):
    pass



if __name__ == '__main__':
    parser = argparse.ArgumentParser('Sends email for every row in a Google Sheets')
    parser.add_argument('sheet_url', type=str)
    parser.add_argument('sheet_name', type=str)
    parser.add_argument('sheet_range', type=str)
    parser.add_argument('--header_lines', type=str, default=None)
    parser.add_argument('--rows_start', type=int, default=None)
    parser.add_argument('--mail_column', type=str, default=None)
    parser.add_argument('--dry-run', default=False, action='store_true')
    parser.add_argument('--mail_credentials_path', type=str, default='mail_credentials.json')
    parser.add_argument('--google_credentials_path', type=str, default='credentials.json')
    parser.add_argument('-d', '--debug', default=False, action='store_true')
    args = parser.parse_args()

    if args.debug:
        logging.basicConfig(level=logging.DEBUG)

    if not os.path.exists(args.mail_credentials_path):
        raise FileNotFoundError('Missing mail credentials file')

    if not os.path.exists(args.google_credentials_path):
        raise FileNotFoundError('Missing google credentials file')

    mail_credentials = load_mail_credentials(args.mail_credentials_path)
    sheet_id = urlparse(args.sheet_url).path.split('/')[3] # /spreadsheets/d/SHEET_ID/edit

    data = read_sheet(args.google_credentials_path, sheet_id, '{}!{}'.format(args.sheet_name, args.sheet_range))

    if args.header_lines is None:
        args.header_lines = input('Specify lines where the header is (1-N, example: 1-5): ')

    if args.rows_start is None:
        try:
            args.rows_start = int(input('Sheet data starts in line (default: line after header): '))
        except ValueError:
            _, args.rows_start = get_header_lines_number(args.header_lines)

    headers = get_header_columns(data, args.header_lines)
    print('Header columns found: ', end='')
    print(', '.join(list(headers.values())))

    mail_column = args.mail_column
    while mail_column not in list(headers.values()):
        mail_column = input('Mail column not found, please specify: ')

    mail_index = 0
    for k, v in headers.items():
        if v == mail_column:
            mail_index = k

    mails = []
    print(args.rows_start)
    for l in range(args.rows_start-1, len(data)):
        temp = ''
        for k, v in headers.items():
            temp += '{}: {}<br/>'.format(v, data[l][k])
        message = mail_credentials['message'].replace('{data}', temp)
        mail = EmailMessage(mail_credentials['subject'], message, mail_credentials['username'], ['eduardom44@gmail.com'])
        mail.content_subtype = "html"  # Main content is now text/html
        mails.append(mail)

    sender = EmailBackend(username=mail_credentials['username'], password=mail_credentials['app_password'], **email_providers[mail_credentials['provider']])
    if args.dry_run:
        for mail in mails:
            print(mail)
    else:
        print(sender.send_messages(mails))






