import os
import pickle
import re
import sys
import urllib
from urllib.parse import urlparse
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import logging


def sheet_id_from_url(url):
    return urlparse(url).path.split('/')[3]  # /spreadsheets/d/SHEET_ID/edit


def read_sheet(credentials_path, sheet_url, sheet_range, download_sheet=None):
    """
    Fetch Google Sheet values using specified range
    :param credentials_path: json file containing google credentials
    :param sheet_id: google sheet id
    :param sheet_range: range in Sheet format (example A1:E3)
    :return: array / sheet values
    """
    sheet_id = sheet_id_from_url(sheet_url)
    # Source: https://developers.google.com/sheets/api/quickstart/python
    creds = None
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        logging.info('Existing Google session token found')
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            logging.info('Session expired. Restarting session on Google')
            creds.refresh(Request())
        else:
            logging.info('Getting authorization from Google')
            flow = InstalledAppFlow.from_client_secrets_file(
                credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            logging.info('Saving Google session token.')
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds, cache_discovery=False)
    # Call the Sheets API
    sheet = service.spreadsheets()
    logging.info('Fetching spreadsheet data')
    result = sheet.values().get(spreadsheetId=sheet_id,
                                range=sheet_range).execute()
    values = result.get('values', [])
    logging.debug(values)

    if download_sheet is not None:
        from google.auth.transport.requests import AuthorizedSession

        authed_session = AuthorizedSession(creds)
        exportUrl = sheet_url.replace('edit', 'export')

        url = exportUrl

        response = authed_session.request(
            'GET', url)

        with open(download_sheet, 'wb') as csvFile:
            csvFile.write(response.content)

    return values


def get_header_lines_number(header_lines):
    """
    Return interval (line number) where header starts and ends.
    :param header_lines: String with interval defined (example 1-3)
    :return: line_start and line_end (int)
    """
    if '-' not in header_lines:
        line_start = line_end = int(header_lines) - 1
    else:
        line_start = int(header_lines.split('-')[0]) - 1
        line_end = int(header_lines.split('-')[1]) - 1

    return line_start, line_end


def get_header_columns(data, header_lines):
    """
    Retrive non-empty header columns with column index
    :param data: Sheet data
    :param header_lines: interval (line number) where is header contained.
    :return: dict containing column index (key) with header name (value)
    """
    header_names = {}
    line_start, line_end = get_header_lines_number(header_lines)

    header_size = 0
    for l in range(line_start, line_end+1):
        if len(data[l]) > header_size:
            header_size = len(data[l])

    for l in range(line_start, line_end+1):
        if len(data[l]) != header_size:
            continue

        for c in range(0, len(data[l])):
            if data[l][c].strip() == '':
                continue

            if c not in header_names:
                header_names[c] = data[l][c]
            else:
                header_names[c] = header_names[c] + ' ' + data[l][c]
            header_names[c] = header_names[c].strip()
    logging.debug(header_names)
    return header_names
