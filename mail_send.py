# Based on: https://github.com/django/django/blob/master/django/core/mail/backends/smtp.py
# and https://docs.djangoproject.com/pt-br/3.0/_modules/django/core/mail/message/
import mimetypes
import smtplib
import ssl
from email import message_from_string, encoders, generator, charset
from email.message import Message
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.utils import formatdate, make_msgid
from io import BytesIO
from pathlib import Path
from email.mime.message import MIMEMessage
from email.mime.text import MIMEText


email_providers = {'outlook': {'host': 'smtp.office365.com', 'port': 587, 'use_tls': True},
                   'gmail': {'host': 'smtp.gmail.com', 'port': 587, 'use_tls': True}
                  }


class EmailBackend:
    def __init__(self, host, port, username, password, use_ssl=False, use_tls=False, fail_silently=False):
        self.host = host
        self.port = port
        self.username = username
        self.password = password
        self.use_ssl = use_ssl
        self.use_tls = use_tls
        self.connection = None
        self.fail_silently = fail_silently

    def open_connection(self):
        self.connection = self.connection_class(self.host, self.port)

        if not self.use_ssl and self.use_tls:
            self.connection.starttls()

        if self.username and self.password:
            self.connection.login(user=self.username, password=self.password)

        return self.connection

    def close_connection(self):
        if self.connection is None:
            return
        try:
            try:
                self.connection.quit()
            except (ssl.SSLError, smtplib.SMTPServerDisconnected):
                # This happens when calling quit() on a TLS connection
                # sometimes, or when the connection was already disconnected
                # by the server.
                self.connection.close()
            except smtplib.SMTPException:
                raise
        finally:
            self.connection = None

    def send_messages(self, email_messages):
        if not email_messages:
            return 0

        new_conn_created = self.open_connection()
        if not self.connection or new_conn_created is None:
            # We failed silently on open().
            # Trying to send would be pointless.
            return 0

        num_sent = 0
        for message in email_messages:
            print('Sending mail to {}...'.format(message.to))
            sent = self._send(message)
            if sent:
                  num_sent += 1
        if new_conn_created:
            self.close_connection()

        return num_sent

    @property
    def connection_class(self):
        return smtplib.SMTP_SSL if self.use_ssl else smtplib.SMTP

    def _send(self, email_message):
        """A helper method that does the actual sending."""
        if not email_message.recipients():
            return False
        encoding = email_message.encoding or charset.Charset('utf-8')
        from_email = email_message.from_email
        recipients = email_message.recipients()
        message = email_message.message()
        try:
            self.connection.sendmail(from_email, recipients, as_bytes(message, linesep='\r\n'))
        except smtplib.SMTPException:
            if not self.fail_silently:
                raise
            return False
        return True


def as_bytes(message, unixfrom=False, linesep='\n'):
    fp = BytesIO()
    g = generator.BytesGenerator(fp, mangle_from_=False)
    g.flatten(message, unixfrom=unixfrom, linesep=linesep)
    return fp.getvalue()


class EmailMessage:
    """A container for email information."""
    content_subtype = 'plain'
    mixed_subtype = 'mixed'
    encoding = None     # None => use settings default

    def __init__(self, subject='', body='', from_email=None, to=None, bcc=None,
                 attachments=None, headers=None, cc=None,
                 reply_to=None):
        """
        Initialize a single email message (which can be sent to multiple
        recipients).
        """
        if to:
            if isinstance(to, str):
                raise TypeError('"to" argument must be a list or tuple')
            self.to = list(to)
        else:
            self.to = []
        if cc:
            if isinstance(cc, str):
                raise TypeError('"cc" argument must be a list or tuple')
            self.cc = list(cc)
        else:
            self.cc = []
        if bcc:
            if isinstance(bcc, str):
                raise TypeError('"bcc" argument must be a list or tuple')
            self.bcc = list(bcc)
        else:
            self.bcc = []
        if reply_to:
            if isinstance(reply_to, str):
                raise TypeError('"reply_to" argument must be a list or tuple')
            self.reply_to = list(reply_to)
        else:
            self.reply_to = []
        self.from_email = from_email
        self.subject = subject
        self.body = body or ''
        self.attachments = []
        if attachments:
            for attachment in attachments:
                if isinstance(attachment, MIMEBase):
                    self.attach(attachment)
                else:
                    self.attach(*attachment)
        self.extra_headers = headers or {}

    def message(self):
        encoding = self.encoding
        msg = MIMEText(self.body, self.content_subtype, encoding)
        msg = self._create_message(msg)
        msg['Subject'] = self.subject
        msg['From'] = self.extra_headers.get('From', self.from_email)
        self._set_list_header_if_not_empty(msg, 'To', self.to)
        self._set_list_header_if_not_empty(msg, 'Cc', self.cc)
        self._set_list_header_if_not_empty(msg, 'Reply-To', self.reply_to)

        # Email header names are case-insensitive (RFC 2045), so we have to
        # accommodate that when doing comparisons.
        header_names = [key.lower() for key in self.extra_headers]
        if 'date' not in header_names:
            # formatdate() uses stdlib methods to format the date, which use
            # the stdlib/OS concept of a timezone, however, Django sets the
            # TZ environment variable based on the TIME_ZONE setting which
            # will get picked up by formatdate().
            msg['Date'] = formatdate()
        if 'message-id' not in header_names:
            msg['Message-ID'] = make_msgid()
        for name, value in self.extra_headers.items():
            if name.lower() != 'from':  # From is already handled
                msg[name] = value
        return msg

    def recipients(self):
        """
        Return a list of all recipients of the email (includes direct
        addressees as well as Cc and Bcc entries).
        """
        return [email for email in (self.to + self.cc + self.bcc) if email]

    def attach(self, filename=None, content=None, mimetype=None):
        """
        Attach a file with the given filename and content. The filename can
        be omitted and the mimetype is guessed, if not provided.
        If the first parameter is a MIMEBase subclass, insert it directly
        into the resulting message attachments.
        For a text/* mimetype (guessed or specified), when a bytes object is
        specified as content, decode it as UTF-8. If that fails, set the
        mimetype to DEFAULT_ATTACHMENT_MIME_TYPE and don't decode the content.
        """
        if isinstance(filename, MIMEBase):
            assert content is None
            assert mimetype is None
            self.attachments.append(filename)
        else:
            assert content is not None
            mimetype = mimetype or mimetypes.guess_type(filename)[0]
            basetype, subtype = mimetype.split('/', 1)

            if basetype == 'text':
                if isinstance(content, bytes):
                    try:
                        content = content.decode()
                    except UnicodeDecodeError:
                        # If mimetype suggests the file is text but it's
                        # actually binary, read() raises a UnicodeDecodeError.
                        mimetype = 'application/octet-stream'

            self.attachments.append((filename, content, mimetype))

    def attach_file(self, path, mimetype=None):
        """
        Attach a file from the filesystem.
        Set the mimetype to DEFAULT_ATTACHMENT_MIME_TYPE if it isn't specified
        and cannot be guessed.
        For a text/* mimetype (guessed or specified), decode the file's content
        as UTF-8. If that fails, set the mimetype to
        DEFAULT_ATTACHMENT_MIME_TYPE and don't decode the content.
        """
        path = Path(path)
        with path.open('rb') as file:
            content = file.read()
            self.attach(path.name, content, mimetype)

    def _create_message(self, msg):
        return self._create_attachments(msg)

    def _create_attachments(self, msg):
        if self.attachments:
            encoding = self.encoding
            body_msg = msg
            msg = MIMEMultipart(_subtype=self.mixed_subtype, encoding=encoding)
            if self.body or body_msg.is_multipart():
                msg.attach(body_msg)
            for attachment in self.attachments:
                if isinstance(attachment, MIMEBase):
                    msg.attach(attachment)
                else:
                    msg.attach(self._create_attachment(*attachment))
        return msg

    def _create_mime_attachment(self, content, mimetype):
        """
        Convert the content, mimetype pair into a MIME attachment object.
        If the mimetype is message/rfc822, content may be an
        email.Message or EmailMessage object, as well as a str.
        """
        basetype, subtype = mimetype.split('/', 1)
        if basetype == 'text':
            encoding = self.encoding
            attachment = MIMEText(content, subtype, encoding)
        elif basetype == 'message' and subtype == 'rfc822':
            # Bug #18967: per RFC2046 s5.2.1, message/rfc822 attachments
            # must not be base64 encoded.
            if isinstance(content, EmailMessage):
                # convert content into an email.Message first
                content = content.message()
            elif not isinstance(content, Message):
                # For compatibility with existing code, parse the message
                # into an email.Message object if it is not one already.
                content = message_from_string(str(content))

            attachment = MIMEMessage(content, subtype)
        else:
            # Encode non-text attachments with base64.
            attachment = MIMEBase(basetype, subtype)
            attachment.set_payload(content)
            encoders.encode_base64(attachment)
        return attachment

    def _create_attachment(self, filename, content, mimetype=None):
        """
        Convert the filename, content, mimetype triple into a MIME attachment
        object.
        """
        attachment = self._create_mime_attachment(content, mimetype)
        if filename:
            try:
                filename.encode('ascii')
            except UnicodeEncodeError:
                filename = ('utf-8', '', filename)
            attachment.add_header('Content-Disposition', 'attachment', filename=filename)
        return attachment

    def _set_list_header_if_not_empty(self, msg, header, values):
        """
        Set msg's header, either from self.extra_headers, if present, or from
        the values argument.
        """
        if values:
            try:
                value = self.extra_headers[header]
            except KeyError:
                value = ', '.join(str(v) for v in values)
            msg[header] = value

    def __str__(self):
        return '[subject={}, from={}, to={}, body={}]'.format(self.subject, self.from_email, self.to, self.body)


class EmailMultiAlternatives(EmailMessage):
    """
    A version of EmailMessage that makes it easy to send multipart/alternative
    messages. For example, including text and HTML versions of the text is
    made easier.
    """
    alternative_subtype = 'alternative'

    def __init__(self, subject='', body='', from_email=None, to=None, bcc=None,
                 connection=None, attachments=None, headers=None, alternatives=None,
                 cc=None, reply_to=None):
        """
        Initialize a single email message (which can be sent to multiple
        recipients).
        """
        super().__init__(
            subject, body, from_email, to, bcc, connection, attachments,
            headers, cc, reply_to,
        )
        self.alternatives = alternatives or []

    def attach_alternative(self, content, mimetype):
        """Attach an alternative content representation."""
        assert content is not None
        assert mimetype is not None
        self.alternatives.append((content, mimetype))

    def _create_message(self, msg):
        return self._create_attachments(self._create_alternatives(msg))

    def _create_alternatives(self, msg):
        encoding = self.encoding
        if self.alternatives:
            body_msg = msg
            msg = MIMEMultipart(_subtype=self.alternative_subtype, encoding=encoding)
            if self.body:
                msg.attach(body_msg)
            for alternative in self.alternatives:
                msg.attach(self._create_mime_attachment(*alternative))
        return msg