# Author: Eduardo Marossi
# Version: 1.1.0
import argparse
import logging
import os
from tempfile import mkstemp

import excel
from gsheet import read_sheet, get_header_lines_number, get_header_columns, sheet_id_from_url
from mail_util import load_mail_credentials, find_mail_column_index, prepare_mails, format_google_url
from mail_send import email_providers, EmailBackend
import yaml

APP_VERSION = '1.1.0'

if __name__ == '__main__':
    argparse.ArgumentParser()
    parser = argparse.ArgumentParser(prog='mailsheet {}'.format(APP_VERSION),
                                     description='Sends email for every row in a Excel or Google Sheet')
    parser.add_argument('--config', default=None,  type=str, help='Config file')
    parser.add_argument('--dry-run', default=False, action='store_true', help='Do not send mail. Show results')
    parser.add_argument('-d', '--debug', default=False, action='store_true', help='Enable debug. Default: off')
    parser.add_argument('--debug-force-to', default=None, type=str, help='Forces all mail to field to specified value.')
    parser.add_argument('--debug-send-interval-start', default=None, type=int, help='Start sending mail after start interval')
    parser.add_argument('--debug-send-interval-end', default=None, type=int, help='End sending mail after end interval.')
    parser.add_argument('-v', '--verbose', default=False, action='store_true', help='Verbose output. Default: off')
    parser.add_argument('--sends-as-file', default=True, action='store_true', help='Sends resulting sheet with header and row data. Recommended if you want tot preserve formattting.')
    args = parser.parse_args()

    if args.debug:
        logging.basicConfig(level=logging.DEBUG)

    mail_credentials_path = 'mail_credentials.json'
    google_credentials_path = 'google_credentials.json'

    if args.config:
        config_file = args.config
    else:
        config_file = 'config.yml'

    with open(config_file, 'r') as file:
        config = yaml.load(file, Loader=yaml.FullLoader)
    config['sheet']['url'] = format_google_url(config['sheet']['url'])

    mail_credentials = load_mail_credentials(mail_credentials_path)

    file_path = None
    if 'google.com' in config["sheet"]["url"]:
        if args.sends_as_file:
            handle, file_path = mkstemp(suffix='.xlsx')
            os.close(handle)
        data = read_sheet(google_credentials_path, config["sheet"]["url"], '{}!{}'.format(config["sheet"]["name"], config["sheet"]["range"]), file_path)
        print('Google Docs temp file: {}'.format(file_path))
    else:
        file_path = config["sheet"]["url"]
        data = excel.read_sheet(config["sheet"]["url"], config["sheet"]["name"], config["sheet"]["range"])

    if args.sends_as_file:
        mails = prepare_mails(data[config["sheet"]["header-rows"]:], config, mail_credentials, file_path)
    else:
        mails = prepare_mails(data[config["sheet"]["start-row"]:], mail_index, config["email"]['subject'], mail_credentials['message'], mail_credentials['username'])

    sender = EmailBackend(username=mail_credentials['username'], password=mail_credentials['app_password'], **email_providers[mail_credentials['provider']])

    if args.debug_force_to is not None:
        for m in mails:
            m.to = [args.debug_force_to]

    for m in mails:
        if config['email']['cc'] is not None:
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




