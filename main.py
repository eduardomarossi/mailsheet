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
import yaml

APP_VERSION = '1.1.0'

if __name__ == '__main__':
    argparse.ArgumentParser()
    parser = argparse.ArgumentParser(prog='mailsheet {}'.format(APP_VERSION),
                                     description='Sends email for every row in a Excel or Google Sheet')
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

    with open('config.yml', 'r') as file:
        config = yaml.load(file, Loader=yaml.FullLoader)

    mail_credentials = load_mail_credentials(args.mail_credentials_path)

    file_path = None
    if 'google.com' in config["sheet"]["url"]:
        if args.sends_as_file:
            handle, file_path = mkstemp(suffix='.xlsx')
            os.close(handle)
        data = read_sheet(args.google_credentials_path, config["sheet"]["url"], '{}!{}'.format(config["sheet"]["name"], config["sheet"]["range"]), file_path)
        print('Google Docs temp file: {}'.format(file_path))
    else:
        file_path = config["sheet"]["url"]
        data = excel.read_sheet(config["sheet"]["url"], config["sheet"]["name"], config["sheet"]["range"])

    if args.sends_as_file:
        mails = prepare_mails(data[config["sheet"]["start-row"]:], config, mail_credentials, file_path)
    else:
        mails = prepare_mails(data[config["sheet"]["start-row"]:], mail_index, config["email"]['subject'], mail_credentials['message'], mail_credentials['username'])

    sender = EmailBackend(username=mail_credentials['username'], password=mail_credentials['app_password'], **email_providers[mail_credentials['provider']])

    if args.debug_force_to is not None:
        for m in mails:
            m.to = [args.debug_force_to]

    for m in mails:
        m.cc.extend([x.strip() for x in config["email"]["cc"].split(';')])

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

    if 'google.com' in config["sheet"]["url"] and args.sends_as_file:
        os.unlink(file_path)




