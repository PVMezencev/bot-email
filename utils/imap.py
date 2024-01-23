import email
import imaplib
import re
import time
from datetime import datetime

from email.header import decode_header

from utils.helpers import EMAIL_VALID_PATTERN, body_decode, sanitize_html

EMAIL_DATE_PATTERN = "%Y-%m-%d %H:%M:%S"


# Чтение входящих писем.
def read(conf: dict) -> list:
    imap_login = conf.get('login')
    imap_password = conf.get('password')
    imap_host = conf.get('host')
    imap_port = conf.get('port')
    imap_inbox = conf.get('inbox')
    imap_archive = conf.get('archive', '')
    imap_read_only = conf.get('read_only', True)
    filter_from = conf.get('filter_from', '')
    filter_from_domain = conf.get('filter_from_domain', '')

    search_filter = 'ALL'
    if imap_read_only:
        search_filter = '(UNSEEN)'
    if filter_from != '':
        search_filter += f' FROM {filter_from}'
    if filter_from_domain != '':
        search_filter += f' HEADER FROM {filter_from_domain}'
    store_filter = '(\Deleted)'
    if imap_read_only:
        store_filter = '\Seen'

    try:
        imap = imaplib.IMAP4_SSL(imap_host, imap_port)
    except Exception as e:
        raise Exception(f'{datetime.utcnow().isoformat(sep="T")}: imaplib.IMAP4_SSL() {e} {type(e)}')

    try:
        imap.login(imap_login, imap_password)
    except Exception as e:
        raise Exception(f'{datetime.utcnow().isoformat(sep="T")}: imap.login(imap_login, imap_password) {e} {type(e)}')

    # Список папок.
    # print(imap.list())
    try:
        imap.select(f'"{imap_inbox}"')
    except Exception as e:
        imap.close()  # Закроем сессию imap.
        imap.logout()  # Отключимся от почтового сервера.
        raise Exception(f'{datetime.utcnow().isoformat(sep="T")}: imap.select() {e} {type(e)}')

    try:
        result, mails_to = imap.uid('search', search_filter)
    except Exception as e:
        imap.close()  # Закроем сессию imap.
        imap.logout()  # Отключимся от почтового сервера.
        raise Exception(f'{datetime.utcnow().isoformat(sep="T")}: imap.uid() {e} {type(e)}')

    email_uuids = []
    if mails_to[0] is not None:
        email_uuids += mails_to[0].split()

    if len(email_uuids) == 0:
        return []

    for email_uid in email_uuids:
        text = ''  # Хранилище текста сообщения.
        try:
            result, mail_data = imap.uid('fetch', email_uid, '(RFC822)')
        except Exception as e:
            _err = f'{datetime.utcnow().isoformat(sep="T")}: {e}'
            print(_err)
            continue

        if mail_data[0] is None:
            continue
        raw_email = mail_data[0][1]
        try:
            raw_email_string = raw_email.decode('utf-8', errors="ignore")
        except Exception as e:
            _err = f'{datetime.utcnow().isoformat(sep="T")}: Ошибка при раскодировке письма {e}'
            print(_err)
            continue

        email_message = email.message_from_string(raw_email_string)

        email_re = re.compile(EMAIL_VALID_PATTERN)

        header_from = ''
        email_header_from = str(email_message['From'])
        if email_re.findall(email_header_from):
            header_from = str(email_re.findall(email_header_from)[0]).lower()

        header_to = []
        email_header_to = str(email_message['To'])
        if email_re.findall(email_header_to):
            for t in email_re.findall(email_header_to):
                header_to.append(str(t).lower())

        header_cc = []
        email_header_cc = str(email_message['Cc'])
        if email_re.findall(email_header_cc):
            for t in email_re.findall(email_header_cc):
                header_cc.append(str(t).lower())

        header_bc = []
        email_header_bc = str(email_message['Bc'])
        if email_re.findall(email_header_bc):
            for t in email_re.findall(email_header_bc):
                header_bc.append(str(t).lower())

        header_rt = ''
        email_header_rt = str(email_message['Reply-To'])
        if email_re.findall(email_header_rt):
            header_rt = str(email_re.findall(email_header_rt)[0]).lower()

        date = email_message['Date']

        # =?utf-8?B?0JLRiyDRg9GB0L/QtdGI0L3QviDRgdC80LXQvdC40LvQuCDQv9Cw0YDQvg==?=
        #  =?utf-8?B?0LvRjCDRg9GH0ZHRgtC90L7QuSDQt9Cw0L/QuNGB0Lgg0L3QsCBGaXJzdFY=?=
        #  =?utf-8?B?RFM=?=

        # '=?utf-8?B?RndkOiDQn9GA0L7QstC10YDQutCw?='
        subject = 'Без темы'
        _raw_subject = email_message['Subject']
        if _raw_subject is not None:
            try:
                _bin_subject, _enc_subject = decode_header(_raw_subject)[0]
                if _enc_subject is not None:
                    subject = _bin_subject.decode(_enc_subject)
                else:
                    subject = str(_bin_subject)
            except Exception as e:
                _err = f'{datetime.utcnow().isoformat(sep="T")}: Ошибка при раскодировке темы {e}'
                print(_err)

        # Попытка получить дату письма.
        try:
            mail_date = datetime.strptime(date, '%a, %d %b %Y %H:%M:%S %z')
        except:
            mail_date = datetime.utcnow()

        attach = list()

        body_text = ''
        body_html = ''
        if email_message.is_multipart():
            for payload in email_message.walk():
                content_type = payload.get_content_type()
                if content_type.startswith('application') \
                        or content_type.startswith('image'):
                    # Получим имя файла, как оно называется в почте.
                    fn = payload.get_filename()
                    if fn is None:
                        continue
                    d_fn = email.header.decode_header(fn)
                    fn = str(email.header.make_header(d_fn))
                    if not (fn):
                        # Ну, куда без имени файла, пропускаем.
                        continue
                    _content = payload.get_payload(decode=True)
                    if _content is not None and len(_content) != 0:
                        attach.append({
                            'name': fn,
                            'content': _content,
                            'content_type': content_type,
                        })
                else:
                    try:
                        chs = payload.get_content_charset()
                        p = payload.get_payload(decode=True).strip()
                    except AttributeError:
                        continue
                    body = body_decode(chs, p)
                    fn = payload.get_filename()
                    if fn is not None:
                        d_fn = email.header.decode_header(fn)
                        fn = str(email.header.make_header(d_fn))
                    if fn is not None:
                        _content = payload.get_payload(decode=True)
                        if _content is not None and len(_content) != 0:
                            attach.append({
                                'name': fn,
                                'content': _content,
                                'content_type': content_type,
                            })
                    else:
                        if content_type == 'text/plain':
                            body_text += body
                        else:
                            body = sanitize_html(body)
                            body_html += body
        else:
            content_type = email_message.get_content_type()
            try:
                chs = email_message.get_content_charset()
                p = email_message.get_payload(decode=True).strip()
            except AttributeError:
                continue
            if p is None:
                print(f'{datetime.utcnow().isoformat(sep="T")}: не удалось получить содержимое письма {subject}')
                continue
            if content_type == 'text/plain':
                body = body_decode(chs, p)
                body_text += body
            elif content_type == 'application/octet-stream':
                fn = email_message.get_filename()
                if fn is not None:
                    d_fn = email.header.decode_header(fn)
                    fn = str(email.header.make_header(d_fn))
                attach.append({
                    'name': fn,
                    'content': p,
                    'content_type': 'application/octet-stream',
                })
            else:
                body = body_decode(chs, p)
                body = sanitize_html(body)
                body_html += body
        if body_html != '':
            text += body_html
            text += '\n'
            _content = body_html.encode()
            if _content is not None and len(_content) != 0:
                attach.append({
                    'name': 'content_body.html',
                    'content': _content,
                    'content_type': 'text/html',
                })
        else:
            text += body_text
            _content = body_text.encode()
            if _content is not None and len(_content) != 0:
                attach.append({
                    'name': 'content_body.txt',
                    'content': _content,
                    'content_type': 'text/plain',
                })
            text += '\n'

        if imap_archive != '':
            imap.uid('COPY', email_uid, imap_archive)

        imap.uid('STORE', email_uid, '+FLAGS', store_filter)

        imap.expunge()

        # Пауза.
        time.sleep(3)

        eml = {
            'subject': subject,
            'date': mail_date,
            'from': header_from,
            'to': header_to,
            'cc': header_cc,
            'bc': header_bc,
            'reply_to': header_rt,
            'body': text,
            'attachments': attach,
            'raw': raw_email_string,
        }

        yield eml

    imap.close()  # Закроем сессию imap.
    imap.logout()  # Отключимся от почтового сервера.
