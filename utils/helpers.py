import re

from bs4 import BeautifulSoup
from html5lib import treebuilders, treewalkers, serializer
import html5lib

from markdownify import markdownify as md


def split_list_by(big_list: list, limit: int) -> list:
    """
    Разделить список на мелкие части длиной limit.
    :param big_list: исходный список для разделения.
    :param limit: по сколько элементов максимум в новых списках.
    :return: список раздроблённых списков.
    """
    # Получим длину исходного списка.
    big_len = len(big_list)
    if big_len <= limit:
        # Если длина исходного списка не превышает заданного размера, вернём её в качестве первого элемента списка.
        return [big_list]
    # Заготовка для ответа.
    super_list = list()
    # Счётчик итераций.
    cntr = 0
    while big_len > limit:
        # Пока полученная длина исходного списка больше лимита...
        # ...добавим срез исходного списка в список ответа
        # в первой итерации будет big_list[0:limit],
        # во второй и последующих со сдвигом на cntr раз.
        super_list.append(big_list[limit * cntr:limit * cntr + limit])
        # Уменьшаем значение длины исходного списка (не саму длину списка, а его значение - копию).
        big_len -= limit
        # Инкрементим счётчик.
        cntr += 1
    # После того, как значение длины исходного списка сравнялось с лимитом или стало меньше - остался не добавленный
    # отрезок от исходного списка, заберём его в ответ.
    super_list.append(big_list[limit * cntr:])
    # Вуа-ля!
    return super_list


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


def __sanitize_html(tag, valid_tags=None):
    for t in tag.children:
        if t.name is None:
            continue
        t = __sanitize_html(t, valid_tags)
    if tag.name not in valid_tags:
        if str(tag.text).strip() == '':
            tag.hidden = True
        else:
            tag.name = 'pre'
    elif str(tag.text).strip() == '':
        tag.hidden = True
    else:
        if tag.name.lower() == 'a':
            _href = tag.attrs.get('href', '')
            if _href != '' and not _href.startswith('@') and not _href.startswith('#'):
                tag.attrs = {
                    'href': _href,
                }
            else:
                tag.hidden = True
    return tag


def sanitize_html(value: str):
    soup = BeautifulSoup(value, 'lxml')

    root_html = soup.find('html')
    if root_html is None:
        root = soup
    else:
        root = root_html.find('body')
    if root is None:
        root = soup

    result_text = ''
    for tag in root.find_all():
        if tag.name.lower() == 'a':
            _href = tag.attrs.get('href', '')
            if _href != '' and not _href.startswith('@') and not _href.startswith('#'):
                tag.attrs = {
                    'href': _href,
                }
            result_text += tag.renderContents().decode()
        elif str(tag.text).strip() != '':
            result_text += f"<pre>{str(tag.text).strip()}</pre>"

    result_text = clean_html(result_text)
    result_text = clean_newline(result_text)
    return result_text


def clean_html(text: str) -> str:
    text = re.sub("(<div.*?div>)", "", text, flags=re.DOTALL)
    text = re.sub("(<table.*?table>)", "", text, flags=re.DOTALL)
    text = text.replace('<br>', '\n')
    text = text.replace('<div>', '\n')
    text = text.replace('</div>', '')
    text = re.sub("(<\?php.*?\?>)", "", text, flags=re.DOTALL)
    text = re.sub("(<!--.*?-->)", "", text, flags=re.DOTALL)
    text = re.sub("(<!DOCTYPE.*?>)", "", text, flags=re.DOTALL)
    text = re.sub("(<!doctype.*?>)", "", text, flags=re.DOTALL)
    text = re.sub("(<head.*?head>)", "", text, flags=re.DOTALL)
    text = re.sub("(<HEAD.*?HEAD>)", "", text, flags=re.DOTALL)
    text = re.sub("(<style.*?style>)", "", text, flags=re.DOTALL)
    text = re.sub("(<STYLE.*?STYLE>)", "", text, flags=re.DOTALL)
    text = re.sub("(<script.*?script>)", "", text, flags=re.DOTALL)
    text = re.sub("(<SCRIPT.*?SCRIPT>)", "", text, flags=re.DOTALL)
    text = re.sub("(<body.*?body>)", "", text, flags=re.DOTALL)
    return text.strip()


def clean_newline(text: str) -> str:
    if len(text) == 0:
        return ''
    text = text.replace('\r\n', '\n')
    text = text.replace(' ', ' ')
    text = text.replace('&nbsp;', ' ')
    text = text.replace('&zwnj;', '')
    text = text.replace('‌', '')
    text = text.replace(' ', '')
    text = text.replace('|', '')
    text = text.replace('---', '')
    pre_symb = str(text[0])
    _str = ''
    for s in text:
        if s in [' ', '\n', '\r', '_'] and pre_symb == s:
            pre_symb = s
            continue
        pre_symb = s
        _str += f'{s}'
    clean_str = ''
    for l in _str.split('\n'):
        if l.strip() == '':
            continue
        clean_str += l + '\n'

    return clean_str.strip()


def tokenizer_html5lib(string):
    p = html5lib.HTMLParser(tree=treebuilders.getTreeBuilder("dom"))
    dom_tree = p.parseFragment(string)
    walker = treewalkers.getTreeWalker("dom")
    stream = walker(dom_tree)

    s = serializer.HTMLSerializer(omit_optional_tags=False)
    return ''.join(s.serialize(stream))


def html2md(html_content, valid_tags=None) -> str:
    soup = BeautifulSoup(html_content, 'lxml')

    root = soup.find('body')
    if root is None:
        root = soup

    for tag in root.find_all():
        if tag.name not in valid_tags or str(tag.text).strip() == '':
            tag.hidden = True
        else:
            if tag.name.lower() == 'a':
                _href = tag.attrs.get('href', '')
                if _href != '' and not _href.startswith('@') and not _href.startswith('#'):
                    tag.attrs = {
                        'href': _href,
                    }
                else:
                    tag.hidden = True
    return clean_newline(md(root.renderContents().decode(), strip=['img', 'table']))