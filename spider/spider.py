# encoding=utf-8
import os.path
import random
import re
import time

import docx
import pandas as pd
import requests
from bs4 import BeautifulSoup
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from pandas import DataFrame

ROOT_PATH = 'http://sousuo.gov.cn/s.htm?t=zhengce&q=%E4%B9%A1%E6%9D%91&timetype=timezd&mintime=2018-02-01&maxtime=2021-11-21&sort=&sortType=1&searchfield=&pcodeJiguan=&childtype=&subchildtype=&tsbq=&pubtimeyear=&puborg=&pcodeYear=&pcodeNum=&filetype=&p=0&n=5&inpro=&sug_t=zhengce'

QUERY_PATH_BASE = 'http://sousuo.gov.cn/data?t=zhengce&q=%E4%B9%A1%E6%9D%91&timetype=timezd&mintime=2018-02-01&maxtime=2021-11-21&sort=&sortType=1&searchfield=&pcodeJiguan=&childtype=&subchildtype=&tsbq=&pubtimeyear=&puborg=&pcodeYear=&pcodeNum=&filetype=&p={}&n=5&inpro='
MAX_PAGE = 500

RESULT_DOCX_PATH = 'result/'

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                  'AppleWebKit/537.36 (KHTML, like Gecko) '
                  'Chrome/67.0.3396.79 Safari/537.36'
}


def sleep_random(const):
    randNum = round(random.random(), 3)
    interval = const + randNum
    print('sleep for ' + str(interval) + ' second...')
    time.sleep(interval)


def get_url_list_from_html(root_url):
    res = requests.get(root_url, headers)

    soup = BeautifulSoup(res.text, "lxml")
    links = soup.find_all(href=re.compile('content'))

    titles = []
    urls = []
    dates = []

    for link in links:
        if not link.get('onclick'):
            continue
        detail = get_title_and_date(link.text)
        titles.append(detail[0])
        dates.append(detail[1])
        url = link.get('href')
        urls.append(url)

    data = {'date': dates, 'title': titles, 'url': urls}
    return DataFrame(data)


def get_title_and_date(content):
    pattern = re.compile('(\d+)\-(\d+)\-(\d+)')
    matcher = re.search(pattern, content)
    date = matcher.group()
    title = content.replace(date, '').strip()
    return [title, date]


def get_url_list_from_query(query_url):
    titles = []
    urls = []
    dates = []
    categories = []
    for i in range(MAX_PAGE):
        sleep_random(1)
        query = query_url.format(i)
        res = requests.get(query, headers)
        json = res.json()
        if is_page_empty(json):
            break

        details = get_url_lists_from_json(json)

        titles.extend(details['title'])
        urls.extend(details['url'])
        dates.extend(details['date'])
        categories.extend(details['category'])

        print('get details on page[{}], current num is [{}]'.format(str(i), str(len(dates))))

    data = {'title': titles, 'url': urls, 'date': dates, 'category': categories}
    return DataFrame(data)


def is_page_empty(json):
    contents = json['searchVO']['catMap']
    keys = contents.keys()
    remain = 0
    for key in keys:
        remain += contents[key]['currentNum']

    return remain == 0


def get_url_lists_from_json(json):
    titles = []
    urls = []
    dates = []
    categories = []

    contents = json['searchVO']['catMap']
    keys = contents.keys()
    for key in keys:
        cat_contents = contents[key]['listVO']
        for item in cat_contents:
            detail = get_single_from_json(item, key)
            titles.append(detail['title'])
            urls.append(detail['url'])
            dates.append(detail['date'])
            categories.append(detail['category'])

    return {'title': titles, 'url': urls, 'date': dates, 'category': categories}


def get_single_from_json(s_json, category):
    pattern = re.compile('\</?em\>')
    title_raw = s_json['title']
    title = re.sub(pattern, '', title_raw)

    return {'title': title, 'url': s_json['url'], 'date': s_json['pubtimeStr'], 'category': category}


def get_articles(frame, result_dir):
    for row in frame.itertuples():
        sleep_random(1)
        get_article_from_single_page(getattr(row, 'date') + '_' + getattr(row, 'title'), getattr(row, 'url'),
                                     result_dir)


def get_article_from_single_page(title, url, result_dir):
    print('start to convert title[{}]...'.format(title))
    res = requests.get(url, headers)
    res.encoding = res.apparent_encoding

    soup = BeautifulSoup(res.text, "lxml")
    if soup.select('.pages_content'):
        text = soup.select('.pages_content')[0].text
    elif soup.select('.b12c'):
        text = soup.select('.b12c')[0].text
    else:
        print('article\'s content is unrecognize at[{}]'.format(title))
        text = soup.text
    text_to_docx(title, text, result_dir)
    print('end to convert title[{}]...'.format(title))


def text_to_docx(name, text, result_dir):
    pattern = re.compile('[^\u4e00-\u9fa5\w]')
    filename = re.sub(pattern, '', name)
    if not os.path.exists(result_dir):
        os.makedirs(result_dir)
    result_file_name = result_dir + filename + '.docx'
    if os.path.exists(result_file_name):
        return

    file = docx.Document()
    file.styles['Normal'].font.name = u'宋体'
    file.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
    file.styles['Normal'].font.size = Pt(10.5)
    file.styles['Normal'].font.color.rgb = RGBColor(0, 0, 0)

    file.add_paragraph(text)

    file.save(result_file_name)


if __name__ == '__main__':
    # frame = get_url_list_from_html(ROOT_PATH)
    # frame.to_csv('test.csv',encoding='utf_8_sig', index=False)

    # frame1 = get_url_list_from_query(QUERY_PATH_BASE)
    # frame1.to_csv('test2.csv', encoding='utf_8_sig', index=False)

    frame2 = pd.read_csv('test2.csv')
    get_articles(frame2, RESULT_DOCX_PATH)
