import json
from argparse import ArgumentError
from mail_send import EmailMessage


def symbols_replace(template, symbols):
    o = template
    for k, v in symbols.items():
        o = o.replace(k, v)
    return o


def prepare_mails(header, data, mail_column, mail_subject, mail_template, mail_username, symbols):
    mails = []
    for l in range(0, len(data)):
        mail = prepare_mail(header, data[l], mail_subject, mail_template, mail_username, data[l][mail_column], symbols)
        mails.append(mail)
    return mails


def prepare_mail(header, row_data, mail_subject, mail_template, mail_username, mail_to, symbols):
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