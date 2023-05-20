from utils.htmltruncate import *
from utils.helpers import *

import email
import imaplib
import io
from datetime import datetime

import yaml
import asyncio
import sys

from aiogram import Bot, Dispatcher, types
from aiogram.types import MediaGroup, InputFile, ParseMode, BotCommand
from aiogram.utils.exceptions import RetryAfter, CantParseEntities, BadRequest
from email.header import decode_header


# Специальная обёртка для исключения, чтоб завершить все асинхронные задачи.
class ErrorThatShouldCancelOtherTasks(Exception):
    pass


# Отправка сообщения ботом (вспомогательная функция).
async def __send_message(bot: Bot, chat_id: int, text: str, mk=None) -> int:
    result = 0
    try:
        try:
            text = truncate(text, 1024, ellipsis=' ...')
        except UnbalancedError:
            text = text[:1024] + ' ...'
        except IndexError:
            text = text[:1024] + ' ...'
        text = tokenizer_html5lib(text)
        message_resp = await bot.send_message(
            chat_id=chat_id,
            text=text,
            parse_mode=types.ParseMode.HTML,
            reply_markup=mk,
            disable_web_page_preview=True,
        )
        result = message_resp.message_id
    except RetryAfter as e:
        print(f'{datetime.utcnow().isoformat(sep="T")}: __send_message(): RetryAfter {e}')
        to_sleeps = re.findall(r'(\d+)', f'{e}')
        if len(to_sleeps) != 0:
            to_sleep = to_sleeps[0]
            await asyncio.sleep(int(to_sleep))
    except CantParseEntities as e:
        message_resp = await bot.send_message(
            chat_id=chat_id,
            text='не удалось разобрать текст...',
            parse_mode=types.ParseMode.HTML,
            reply_markup=mk,
            disable_web_page_preview=True,
        )
        result = message_resp.message_id
    except BadRequest as e:
        if f'{e}' in ['Entities_too_long', 'Message is too long']:
            message_resp = await bot.send_message(
                chat_id=chat_id,
                text='слишком длинный для Телеграм текст...',
                parse_mode=types.ParseMode.HTML,
                reply_markup=mk,
                disable_web_page_preview=True,
            )
            result = message_resp.message_id
        else:
            print(f'{datetime.utcnow().isoformat(sep="T")}: __send_message(): {e}')
    except Exception as e:
        print(f'{datetime.utcnow().isoformat(sep="T")}: __send_message(): {e}')
    finally:
        return result


# Отправка файлов (вспомогательная функция).
async def __send_files(bot: Bot, chat_id: int, reply_message_id: int, files: list) -> bool:
    result = False
    split_files = split_list_by(files, 10)
    cnt_split_files = 0
    len_split_files = len(split_files)
    try:
        while len_split_files > 0:
            media = MediaGroup()
            for f in split_files[cnt_split_files]:
                _content = f.get('content')
                _bytes_file = io.BytesIO(_content)
                _document = InputFile(path_or_bytesio=_bytes_file, filename=f.get('name'))
                media.attach_document(document=_document)
            try:
                await bot.send_media_group(chat_id=chat_id, reply_to_message_id=reply_message_id, media=media)
                len_split_files -= 1
                cnt_split_files += 1
            except RetryAfter as e:
                print(f'{datetime.utcnow().isoformat(sep="T")}: send_files_telegram() RetryAfter {e}')
                to_sleeps = re.findall(r'(\d+)', f'{e}')
                if len(to_sleeps) != 0:
                    to_sleep = to_sleeps[0]
                    await asyncio.sleep(int(to_sleep))
                else:
                    await asyncio.sleep(10)
        result = True
    except Exception as e:
        print(f'{datetime.utcnow().isoformat(sep="T")}: send_files_telegram() {e}')
    finally:
        # s = await bot.get_session()
        # await s.close()
        return result


# Отправка текстового сообщения.
async def send_message(bot: Bot, chat_id: int, text: str, mk=None):
    send_counter = 3
    message_id = await __send_message(bot=bot, chat_id=chat_id, text=text, mk=mk)
    while message_id == 0:
        await asyncio.sleep(2)
        if send_counter <= 0:
            print(f'{datetime.utcnow().isoformat(sep="T")}: не удачная попытка отправки')
            _headers = text.split('\n')
            if len(_headers) >= 4:
                _subtext = '\n'.join(_headers[:4]) + '\n' + 'не удалось разобрать структуру письма...'
                message_id = await __send_message(bot=bot, chat_id=chat_id, text=_subtext, mk=mk)
            break
        message_id = await __send_message(bot=bot, chat_id=chat_id, text=text, mk=mk)
        send_counter -= 1
    return message_id


# Отправка одного фото.
async def send_photo(bot: Bot, chat_id: int, photo: dict, caption: str, mk=None) -> bool:
    result = False
    _content = photo.get('content')
    _bytes_file = io.BytesIO(_content)
    _document = InputFile(path_or_bytesio=_bytes_file, filename=photo.get('name'))
    try:
        await bot.send_photo(chat_id=chat_id, photo=_document, caption=caption, parse_mode=ParseMode.MARKDOWN,
                             reply_markup=mk)
        result = True
    except RetryAfter as e:
        print(f'{datetime.utcnow().isoformat(sep="T")}: send_photo_telegram() RetryAfter {e}')
        to_sleeps = re.findall(r'(\d+)', f'{e}')
        if len(to_sleeps) != 0:
            to_sleep = to_sleeps[0]
            await asyncio.sleep(int(to_sleep))
        else:
            await asyncio.sleep(10)
    except Exception as e:
        print(f'{datetime.utcnow().isoformat(sep="T")}: send_photo_telegram() {e}')
    finally:
        return result


# Отправка вложений.
async def send_attach(bot: Bot, chat_id: int, text: str, files: list):
    send_counter = 3
    message_id = await send_message(bot=bot, chat_id=chat_id, text=text)
    while message_id == 0:
        await asyncio.sleep(2)
        if send_counter <= 0:
            print(f'{datetime.utcnow().isoformat(sep="T")}: не удачная попытка отправки')
            break
        message_id = await send_message(bot=bot, chat_id=chat_id, text=text)
        send_counter -= 1

    if len(files) == 0:
        print(f'{datetime.utcnow().isoformat(sep="T")}: нет файлов')
        return
    if message_id == 0 or message_id is None:
        print(f'{datetime.utcnow().isoformat(sep="T")}: не удалось отправить описание')
        return
    sen_counter = 3
    while not await __send_files(bot=bot, chat_id=chat_id, reply_message_id=int(message_id), files=files):
        if sen_counter <= 0:
            print(f'{datetime.utcnow().isoformat(sep="T")}: не удачная попытка отправки')
            break
        sen_counter -= 1
        await asyncio.sleep(2)


# Чтение входящих писем.
async def parse_inbox(user: dict, user_request=False):
    b = user.get('bot')
    tid = int(user.get('telegram_id'))
    imap_login = user.get('login')
    imap_password = user.get('password')
    imap_host = user.get('host')
    imap_port = user.get('port')
    imap_inbox = user.get('inbox')
    imap_archive = user.get('archive', '')
    imap_read_only = user.get('read_only', True)

    search_filter = 'ALL'
    if imap_read_only:
        search_filter = '(UNSEEN)'
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
        if user_request:
            await send_message(bot=b, chat_id=tid, text='ящик пустой')
        return

    for email_uid in email_uuids:
        text = ''  # Хранилище текста сообщения.
        try:
            result, mail_data = imap.uid('fetch', email_uid, '(RFC822)')
        except Exception as e:
            _err = f'{datetime.utcnow().isoformat(sep="T")}: {e}'
            print(_err)
            await send_message(bot=b, chat_id=tid, text=_err)
            continue
        if mail_data[0] is None:
            continue
        raw_email = mail_data[0][1]
        try:
            raw_email_string = raw_email.decode('utf-8', errors="ignore")
        except Exception as e:
            _err = f'{datetime.utcnow().isoformat(sep="T")}: Ошибка при раскодировке письма {e}'
            print(_err)
            await send_message(bot=b, chat_id=tid, text=_err)
            continue

        email_message = email.message_from_string(raw_email_string)

        email_re = re.compile(EMAIL_VALID_PATTERN)
        header_from = ''
        email_header_from = str(email_message['From'])
        if email_re.findall(email_header_from):
            header_from = str(email_re.findall(email_header_from)[0]).lower()

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

        # Форматируем в MD: дата - моноширинный, тема - полужирный.
        text += '<b>' + subject + '</b>' + '\n' + '<code>' + date + '</code>' + '\n' + 'От <b>' + header_from + '</b>' + '\n' + '_' + '\n'
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
            body = body_decode(chs, p)
            if content_type == 'text/plain':
                body_text += body
            else:
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

        # Флаг, указывающий, что во вложении только картинки, чтоб собрать медиагруппу и отправить в телеграмм
        # не как файлы, а как изображения.
        attach_is_image_only = True
        for a in attach:
            if not a.get('content_type').startswith('image'):
                attach_is_image_only = False
                break

        if len(attach) == 0:
            if text != '':
                # Если удалось собрать текст для сообщения, то разошлём его по получателям.
                await send_message(bot=b, chat_id=tid, text=text)
        elif len(attach) == 1 and attach_is_image_only:
            await send_photo(bot=b, chat_id=tid, photo=attach[0], caption=text)
        else:
            await send_attach(bot=b, chat_id=tid, text=text, files=attach)

        if imap_archive != '':
            imap.uid('COPY', email_uid, imap_archive)

        imap.uid('STORE', email_uid, '+FLAGS', store_filter)
        await asyncio.sleep(3)

        imap.expunge()
    imap.close()  # Закроем сессию imap.
    imap.logout()  # Отключимся от почтового сервера.


async def bot_command_handler(d: Dispatcher, user: dict):
    register_handlers(d)
    await set_commands(user)
    while True:
        try:
            await d.start_polling()
        except (KeyboardInterrupt, SystemExit, RuntimeError) as e:
            break
        except RetryAfter as e:
            print(f'{datetime.utcnow().isoformat(sep="T")}: disp.start_polling() RetryAfter {e}')
            to_sleeps = re.findall(r'(\d+)', f'{e}')
            if len(to_sleeps) != 0:
                to_sleep = to_sleeps[0]
                await asyncio.sleep(int(to_sleep))
        except Exception as e:
            print(f'{datetime.utcnow().isoformat(sep="T")}: disp.start_polling(): {e}')
            await asyncio.sleep(2)
        raise ErrorThatShouldCancelOtherTasks


# Запуск чтения почты один раз (если запускать скрипт через крон).
async def pareser_start_once(user: dict):
    try:
        await parse_inbox(user)
    except Exception as e:
        print(f'{datetime.utcnow().isoformat(sep="T")}: parse_inbox(): {e}')


# Запуск чтения почты раз в минуту в бесконечном цикле.
async def pareser_start_cycle(user: dict):
    while True:
        try:
            await parse_inbox(user)
            await asyncio.sleep(60)
        except Exception as e:
            print(f'{datetime.utcnow().isoformat(sep="T")}: parse_inbox(): {e}')
            await asyncio.sleep(60)


async def cmd_read_handler(message: types.Message):
    try:
        await parse_inbox(user, user_request=True)
    except Exception as e:
        await message.answer(f'Ошибка: {e}')


async def cmd_me_handler(message: types.Message):
    await message.answer(f'Мой TelegramID: `{message.from_user.id}`\nTelegramID чата: `{message.chat.id}`', parse_mode=types.ParseMode.MARKDOWN_V2)


# Регистрация команд, отображаемых в интерфейсе Telegram
async def set_commands(user: dict):
    b = user.get('bot')
    commands = [
        BotCommand(command="/read", description="Запросить почту из ящика"),
        BotCommand(command="/me", description="Показать информацию обо мне"),
    ]
    await b.set_my_commands(commands)


def register_handlers(d: Dispatcher):
    d.register_message_handler(cmd_read_handler, commands=['read'])
    d.register_message_handler(cmd_me_handler, commands=['me'])


# Главная функция.
async def main(user_data: dict, d: Dispatcher = None, cycle=False):
    tasks = []
    if d is not None:
        tasks.append(asyncio.create_task(bot_command_handler(d, user_data)))

    if cycle:
        tasks.append(asyncio.create_task(pareser_start_cycle(user_data)))
    else:
        tasks.append(asyncio.create_task(pareser_start_once(user_data)))

    try:
        await asyncio.gather(*tasks)
    except ErrorThatShouldCancelOtherTasks:
        for t in tasks:
            t.cancel()
    finally:
        if d is not None:
            s = await d.bot.get_session()
            await s.close()


# Константы.
CFG_PATH = 'config-bot-email.yml'
EMAIL_VALID_PATTERN = r'([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)'

# Начало выполнения.
if __name__ == '__main__':
    try:
        with open(CFG_PATH, 'r') as yml_file:
            CFG = yaml.load(yml_file, Loader=yaml.FullLoader)
    except FileNotFoundError:
        sys.exit(f'Укажите файл конфигурации {CFG_PATH}')

    bot = Bot(token=CFG.get('bot'))
    disp = None

    cfg_imap = CFG.get('imap')
    if cfg_imap is None:
        sys.exit(f'Укажите настройки для imap {CFG_PATH}')

    if CFG.get('start_bot', False):
        disp = Dispatcher(bot)

    user = {
        'bot': bot,
        'login': cfg_imap.get('login'),
        'password': cfg_imap.get('password'),
        'inbox': cfg_imap.get('inbox'),
        'archive': cfg_imap.get('archive'),
        'read_only': cfg_imap.get('read_only', True),
        'host': cfg_imap.get('host'),
        'port': cfg_imap.get('port'),
        'telegram_id': CFG.get('my_telegram_id'),
    }

    try:
        asyncio.run(main(user, disp, CFG.get('is_cycle', False)))
    except KeyboardInterrupt:
        sys.exit(0)
