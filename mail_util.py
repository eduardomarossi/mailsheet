import json
import os
from argparse import ArgumentError
from datetime import datetime
from shutil import copyfile
from tempfile import mkstemp

from excel import open_sheet_keep_row
from mail_send import EmailMessage


def symbols_replace(template, symbols):
    o = template
    for k, v in symbols.items():
        o = o.replace(k, v)
    return o


def prepare_mails(header, data, mail_column, mail_subject, mail_template, mail_username, symbols, file_path=None, data_start=None, sheet_name=None):
    mails = []

    dir_path = ''
    if file_path is not None:
        now = datetime.now()
        dir_path = 'temp{:04d}-{:02d}-{:02d}'.format(now.year, now.month, now.day)
        if not os.path.exists(dir_path):
            os.mkdir(dir_path)

    for l in range(0, len(data)):
        mail_to = data[l][mail_column]
        mail_attach = None

        if file_path is not None:
            now = datetime.now()
            starting_name = mail_to[:mail_to.find('@')] + '-{}-{}'.format(now.second, now.microsecond)
            mail_attach = os.path.join(dir_path, starting_name + '.xlsx')
            open_sheet_keep_row(file_path, mail_attach, sheet_name, data_start, l)

        mail = prepare_mail(header, data[l], mail_subject, mail_template, mail_username, mail_to, symbols, mail_attach)
        mails.append(mail)
    return mails


def prepare_mail(header, row_data, mail_subject, mail_template, mail_username, mail_to, symbols, mail_attach=None):
    data = ''
    for k, v in header.items():
        data += '{}: {}<br/>'.format(v, row_data[k])
    symbols['{data}'] = data

    subject = symbols_replace(mail_subject, symbols)
    username = symbols_replace(mail_username, symbols)
    message = symbols_replace(mail_template, symbols)
    to = symbols_replace(mail_to, symbols)

    mail = EmailMessage(subject, message, username, [to])
    mail.content_subtype = "html"
    if mail_attach is not None:
        mail.attach_file(mail_attach, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    return mail


def load_mail_credentials(path):
    """
    Loads mail credentials from json file. Check if required fields are defined.
    :param path: location of json file
    :return: dict with provider, username, app_password and message
    """
    with open(path) as f:
        data = json.loads(f.read())

    required = ['provider', 'username', 'app_password', 'message']
    for k in required:
        if k not in data:
            raise ArgumentError('Missing required information in mail_credentials json file.')

    return data


def find_mail_column_index(headers, mail_column):
    for k, v in headers.items():
        if v == mail_column:
            return k
    return None