import email
import os
import re
from datetime import datetime
from email.header import decode_header

# from utils.helpers import EMAIL_VALID_PATTERN, body_decode, sanitize_html
EMAIL_VALID_PATTERN = r'([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)'

def body_decode(chs, payload) -> str:
    if chs is not None:
        try:
            body = payload.decode(encoding=chs)
        except:
            body = payload.decode(encoding=chs, errors='ignore')
    else:
        try:
            body = payload.decode()
        except:
            body = payload.decode(errors='ignore')
    if '\\u0' in body:
        body = body.encode(encoding=chs).decode('unicode-escape')
    return body

def parse(eml):
    text = ''  # Хранилище текста сообщения.
    try:
        raw_email_string = eml.decode('utf-8', errors="ignore")
    except Exception as e:
        _err = f'{datetime.utcnow().isoformat(sep="T")}: Ошибка при раскодировке письма {e}'
        print(_err)
        return

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
        # Date: Sat, 9 Sep 2023 01:15:44 +0000
        mail_date = datetime.strptime(date, '%a, %d %b %Y %H:%M:%S %z')
    except:
        try:
            # Date: Wed, 09 Feb 2022 05:27:06 GMT
            mail_date = datetime.strptime(date, '%a, %d %b %Y %H:%M:%S %Z')
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
                        # body = sanitize_html(body)
                        body_html += body
    else:
        content_type = email_message.get_content_type()
        try:
            chs = email_message.get_content_charset()
            p = email_message.get_payload(decode=True).strip()
        except AttributeError:
            return
        if p is None:
            print(f'{datetime.utcnow().isoformat(sep="T")}: не удалось получить содержимое письма {subject}')
            return
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
            # body = sanitize_html(body)
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

    return eml


def list_files(root, ext=None):
    for file in os.listdir(root):
        path = os.path.join(root, file)
        if not os.path.isdir(path):
            if ext is None:
                yield path
            elif str(path).endswith(ext):
                yield path
            else:
                continue
        else:
            for p in list_files(path, ext):
                yield p


if __name__ == '__main__':
    backups_dir = 'backups'
    new_dir = 'attachments'

    for d in list_files('backups'):
        dir_name = os.path.dirname(d)
        file_name = os.path.basename(d)

        try:
            os.makedirs(new_dir)
        except:
            pass

        with open(d, 'rb') as eml:
            payload = eml.read()
            data = parse(payload)

            for att in data.get('attachments'):
                new_attach_name = os.path.join(new_dir, file_name + att.get('name'))
                with open(new_attach_name, 'wb') as a:
                    a.write(att.get('content'))
