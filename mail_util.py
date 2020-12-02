import json
import os
from argparse import ArgumentError
from datetime import datetime
from shutil import copyfile
from tempfile import mkstemp

from excel import open_sheet_keep_row
from mail_send import EmailMessage
import markdown2

def symbols_replace(template, symbols):
    o = template
    for k, v in symbols.items():
        o = o.replace(k, v)
    return o

# https://stackoverflow.com/questions/4528982/convert-alphabet-letters-to-number-in-python
def letter_num(char):
    return(char)
    return ([ord(char) - 96 for char in char.lower()] - 1)

def format_google_url(url):
    return(url.split('edit')[0]+'edit')

def prepare_mails(data, config, credentials, file_path=None):
    mails = []

    dir_path = ''
    if file_path is not None:
        now = datetime.now()
        dir_path = 'temp{:04d}-{:02d}-{:02d}'.format(now.year, now.month, now.day)
        if not os.path.exists(dir_path):
            os.mkdir(dir_path)

    for l in range(0, len(data)):
        mail_to = data[l][config["sheet"]["email-col"] - 1]
        mail_attach = None

        if file_path is not None:
            now = datetime.now()
            starting_name = mail_to[:mail_to.find('@')] + '-{}-{}'.format(now.second, now.microsecond)
            mail_attach = os.path.join(dir_path, starting_name + '.xlsx')
            open_sheet_keep_row(file_path, mail_attach, config["sheet"]["name"], config["sheet"]["header-rows"] + config["sheet"]["start-row"], l)

        mail = prepare_mail(config["email"]["subject"], config["email"]["msg"], credentials['username'], mail_to, mail_attach)
        mails.append(mail)
    return mails


def prepare_mail(mail_subject, mail_msg, mail_username, mail_to, mail_attach=None):
    data = ''
    symbols = {}
    symbols['{data}'] = data

    subject = symbols_replace(mail_subject, symbols)
    username = symbols_replace(mail_username, symbols)
    message = symbols_replace(markdown2.markdown(mail_msg), symbols)
    to = symbols_replace(mail_to, symbols).split(';')
    to = [x.strip() for x in to]
    mail = EmailMessage(subject, message, username, to)
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
