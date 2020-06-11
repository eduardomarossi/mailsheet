# Author: Eduardo Marossi
# Version: 1.1.0
import argparse
import logging
import os
from tempfile import mkstemp

import excel
from gsheet import read_sheet, get_header_lines_number, get_header_columns, sheet_id_from_url
from mail_util import load_mail_credentials, find_mail_column_index, prepare_mails
from mail_send import email_providers, EmailBackend


APP_VERSION = '1.1.0'

if __name__ == '__main__':
    argparse.ArgumentParser()
    parser = argparse.ArgumentParser(prog='mailsheet {}'.format(APP_VERSION),
                                     description='Sends email for every row in a Excel or Google Sheet')
    parser.add_argument('sheet_url', type=str, help='Google Sheets URL / Excel file')
    parser.add_argument('sheet_name', type=str, help='Sheet name')
    parser.add_argument('sheet_range', type=str, help='Sheet range. Example A1:B2')
    parser.add_argument('--header-lines', type=str, default=None, help='Interval (lines) where header is located. Examples: 1 or 1-3.')
    parser.add_argument('--rows-start', type=int, default=None, help='Line number where table data starts. Default: Header line + 1')
    parser.add_argument('--mail-column', type=str, default=None, help='Specifies mail column name (header)')
    parser.add_argument('--dry-run', default=False, action='store_true', help='Do not send mail. Show results')
    parser.add_argument('--mail-credentials-path', type=str, default='mail_credentials.json', help='Custom path for mail credentials json file. Default: mail_credentials.json')
    parser.add_argument('--google-credentials-path', type=str, default='credentials.json', help='Custom path for google credentials json file. Default: credentials.json')
    parser.add_argument('-d', '--debug', default=False, action='store_true', help='Enable debug. Default: off')
    parser.add_argument('--debug-force-to', default=None, type=str, help='Forces all mail to field to specified value.')
    parser.add_argument('-c', '--add-cc', default=[], action='append', help='Adds mail to cc field.')
    parser.add_argument('--debug-send-interval-start', default=None, type=int, help='Start sending mail after start interval')
    parser.add_argument('--debug-send-interval-end', default=None, type=int, help='End sending mail after end interval.')
    parser.add_argument('-v', '--verbose', default=False, action='store_true', help='Verbose output. Default: off')
    parser.add_argument('--sends-as-file', default=True, action='store_true', help='Sends resulting sheet with header and row data. Recommended if you want tot preserve formattting.')
    args = parser.parse_args()

    if args.debug:
        logging.basicConfig(level=logging.DEBUG)

    if not os.path.exists(args.mail_credentials_path):
        raise FileNotFoundError('Missing mail credentials file')

    if not os.path.exists(args.google_credentials_path):
        raise FileNotFoundError('Missing google credentials file')

    if args.header_lines is None:
        args.header_lines = input('\nSpecify lines numbers where the header is (1-N, example: 1-5): ')

    if args.rows_start is None:
        try:
            args.rows_start = int(input('Sheet data starts in line number (default: line after header): ')) - 1
        except ValueError:
            _, args.rows_start = get_header_lines_number(args.header_lines)
            args.rows_start += 1

    mail_credentials = load_mail_credentials(args.mail_credentials_path)

    file_path = None
    if 'google.com' in args.sheet_url:
        if args.sends_as_file:
            handle, file_path = mkstemp(suffix='.xlsx')
            os.close(handle)
        data = read_sheet(args.google_credentials_path, args.sheet_url, '{}!{}'.format(args.sheet_name, args.sheet_range), file_path)
        print('Google Docs temp file: {}'.format(file_path))
    else:
        file_path = args.sheet_url
        data = excel.read_sheet(args.sheet_url, args.sheet_name, args.sheet_range)

    headers = get_header_columns(data, args.header_lines)

    print('\nHeader columns found: ', end='')
    print(', '.join(list(headers.values())))

    mail_column = args.mail_column
    while mail_column not in list(headers.values()):
        mail_column = input('Mail column not found, please specify: ')

    mail_index = find_mail_column_index(headers, mail_column)


    if args.sends_as_file:
        mails = prepare_mails(headers, data[args.rows_start:], mail_index, mail_credentials['subject'],
                              mail_credentials['message'], mail_credentials['username'], file_path, args.rows_start+1, args.sheet_name)
    else:
        mails = prepare_mails(headers, data[args.rows_start:], mail_index, mail_credentials['subject'], mail_credentials['message'], mail_credentials['username'])

    sender = EmailBackend(username=mail_credentials['username'], password=mail_credentials['app_password'], **email_providers[mail_credentials['provider']])

    if args.debug_force_to is not None:
        for m in mails:
            m.to = [args.debug_force_to]

    for m in mails:
        m.cc.extend(args.add_cc)

    if args.debug_send_interval_start is not None and args.debug_send_interval_end is not None:
        mails = mails[args.debug_send_interval_start:args.debug_send_interval_end]
    elif args.debug_send_interval_start is not None:
        mails = mails[args.debug_send_interval_start:]
    elif args.debug_send_interval_end is not None:
        mails = mails[:args.debug_send_interval_end]

    if args.dry_run:
        print('Results in {} mails:'.format(len(mails)))
        for mail in mails:
            print(mail)
            print(' ')
    else:
        print('Sending mails...')
        print('Sent {} mails'.format(sender.send_messages(mails)))

    if 'google.com' in args.sheet_url and args.sends_as_file:
        os.unlink(file_path)




