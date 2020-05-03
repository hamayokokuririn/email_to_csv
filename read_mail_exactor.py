#!/usr/bin/env python

import sys
import email
from email.header import decode_header

class MailParser(object):

    def __init__(self, mail_file_path):
        self.mail_file_path = mail_file_path
        with open(mail_file_path, 'rb') as email_file:
            self.email_message = email.message_from_bytes(email_file.read())
        self.subject = None
        self.to_address = None
        self.from_address = None
        self.body = ""
        self.textPlain = ""
        self._parse()


    def get_attr_data(self):
        result = """\
            From: {}
            To: {}
            -------
            Body:
            {}
            -------
            """.format(self.from_address,
            self.to_address,
            self.body)
    
        return result

        
    def _parse(self):
        self.subject = self._get_decoded_header("Subject")
        self.to_address = self._get_decoded_header("To")
        self.from_address = self._get_decoded_header("From")

        for part in self.email_message.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            
            attach_fname = part.get_filename()

            if not attach_fname:
                charset = str(part.get_content_charset())
                if charset:
                    self.body += part.get_payload(decode=True).decode(charset, errors='replace')
                else:
                    self.body += part.get_payload(decode=True)
            

    def _get_decoded_header(self, key_name):
        ret = ""
        raw_obj = self.email_message.get(key_name)
        if raw_obj is None:
            return ""
        for fragment, encoding in decode_header(raw_obj):
            if not hasattr(fragment, "decode"):
                ret += fragment
                continue
            if encoding:
                ret += fragment.decode(encoding)
            else:
                ret += fragment.decode("UTF-8")
        return ret

def get_text_plain(msg):
    body = ''
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        attach_fname = part.get_filename()
        if not attach_fname:
            charset = str(part.get_content_charset())
            if charset:
                if part.get_content_type() == 'text/plain':
                    body += part.get_payload(decode=True).decode(charset, errors='replace')
            else:
                body += part.get_payload(decode=True)

    return body

if __name__ == "__main__":
    path = 'c:/Users/xxxxxxx'
    mail = MailParser(path)
    #print(mail.get_attr_data)

    print(get_text_plain(mail.email_message))
