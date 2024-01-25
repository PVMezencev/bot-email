import argparse
import os

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

from utils.imap import read

MAX_MESSAGE = 4096


# Специальная обёртка для исключения, чтоб завершить все асинхронные задачи.
class ErrorThatShouldCancelOtherTasks(Exception):
    pass


# Отправка сообщения ботом (вспомогательная функция).
async def __send_message(bot: Bot, chat_id: int, text: str, mk=None) -> int:
    if bot is None:
        return 0
    result = 0
    try:
        try:
            text = truncate(text, MAX_MESSAGE, ellipsis=' ...')
        except UnbalancedError:
            text = text[:MAX_MESSAGE] + ' ...'
        except IndexError:
            text = text[:MAX_MESSAGE] + ' ...'
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
    if bot is None:
        return True
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
    if bot is None:
        return
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
    if bot is None:
        return True
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
    if bot is None:
        return
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

    filter_ext = user.get('filter_ext', '')

    counter = 0

    try:
        for eml in read(user):
            counter += 1
            subject = eml.get('subject', '')
            mail_date = eml.get('date', '')
            header_from = eml.get('header_from', '')
            attach = eml.get('attachments', [])

            # Форматируем в MD: дата - моноширинный, тема - полужирный. 'Sat, 15 Jul 2023 04:55:05 +0400'
            text = '<b>' + subject + '</b>' + '\n' + '<code>' + mail_date.strftime(
                '%a, %d %b %Y %H:%M:%S %z') + '</code>' + '\n' + 'От <b>' + header_from + '</b>' + '\n' + '_' + '\n'
            text += eml.get('body', '')
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

            save_to = user.get('attaches_save_to')
            if save_to is None or save_to == '':
                for a in attach:
                    aname = a.get('name')
                    if 'content_body' in aname:
                        continue
                    if filter_ext != '':
                        if not aname.endswith(filter_ext):
                            continue
            else:
                save_to_full = os.path.join(save_to,
                                            mail_date.strftime('%Y'),
                                            mail_date.strftime('%m'),
                                            mail_date.strftime('%d'),
                                            )
                if not os.path.exists(save_to_full):
                    try:
                        os.makedirs(save_to_full, 0o755)
                    except Exception as e:
                        raise e
                for a in attach:
                    aname = a.get('name')
                    acontent = a.get('content')
                    if 'content_body' in aname:
                        continue
                    if filter_ext != '':
                        if not aname.endswith(filter_ext):
                            continue
                    fp = os.path.join(save_to_full, aname)
                    with open(fp, 'wb') as aw:
                        aw.write(acontent)

            backup_to = user.get('backup_save_to', '')
            if backup_to != '':
                save_to_full = os.path.join(backup_to,
                                            mail_date.strftime('%Y'),
                                            mail_date.strftime('%m'),
                                            mail_date.strftime('%d'),
                                            )
                if not os.path.exists(save_to_full):
                    try:
                        os.makedirs(save_to_full, 0o755)
                    except Exception as e:
                        raise e

                subj_prepare = eml.get("subject")[:10].replace(" ", "_").replace(":", "_").replace("/", "_").replace("\\", "_")
                emlname = f'{counter}_{eml.get("from")}_{subj_prepare}.eml'
                fp = os.path.join(save_to_full, emlname)
                with open(fp, 'w') as aw:
                    aw.write(eml.get("raw"))

    except Exception as e:
        raise Exception(e)

    if counter == 0 and user_request:
        await send_message(bot=b, chat_id=tid, text='ящик пустой')


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
    await message.answer(f'Мой TelegramID: `{message.from_user.id}`\nTelegramID чата: `{message.chat.id}`',
                         parse_mode=types.ParseMode.MARKDOWN_V2)


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


# Начало выполнения.
if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Обработка почты.', argument_default='')
    parser.add_argument('--config', help='путь к файлу конфигурации', type=str, default='config-bot-email.yml')
    params = parser.parse_args()

    CFG_PATH = params.config

    try:
        with open(CFG_PATH, 'r') as yml_file:
            CFG = yaml.load(yml_file, Loader=yaml.FullLoader)
    except FileNotFoundError:
        sys.exit(f'Укажите файл конфигурации {CFG_PATH}')

    bot = None
    disp = None
    if CFG.get('bot', '') != '':
        bot = Bot(token=CFG.get('bot'))

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
        'archive': cfg_imap.get('archive', ''),
        'read_only': cfg_imap.get('read_only', True),
        'host': cfg_imap.get('host'),
        'port': cfg_imap.get('port'),
        'filter_from': cfg_imap.get('filter_from'),
        'filter_from_domain': cfg_imap.get('filter_from_domain'),
        'telegram_id': CFG.get('my_telegram_id'),
        'filter_ext': CFG.get('filter_ext'),
        'attaches_save_to': CFG.get('attaches_save_to'),
        'backup_save_to': CFG.get('backup_save_to'),
    }

    try:
        asyncio.run(main(user, disp, CFG.get('is_cycle', False)))
    except KeyboardInterrupt:
        sys.exit(0)
