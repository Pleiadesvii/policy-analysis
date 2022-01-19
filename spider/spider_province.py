# encoding=utf-8
import os.path
import random
import re
import time
from urllib import parse

import docx
import pandas as pd
import requests
from bs4 import BeautifulSoup
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from pandas import DataFrame

ROOT_PATH = 'http://search.zj.gov.cn/jsearchfront/search.do?websiteid=330000000000000&searchid=&pg=&p={}&tpl=1569&cateid=372&fbjg=&word=%E4%B9%A1%E6%9D%91&temporaryQ=&synonyms=&checkError=1&isContains=1&q=%E4%B9%A1%E6%9D%91&jgq=&eq=&begin=20180101&end=20211130&timetype=5&_cus_pq_ja_type=&pos=title&sortType=1'

QUERY_PATH_BASE = 'http://search.zj.gov.cn/jsearchfront/interfaces/cateSearch.do'

MAX_PAGE = 112

RESULT_DOCX_PATH = 'result/'

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                  'AppleWebKit/537.36 (KHTML, like Gecko) '
                  'Chrome/67.0.3396.79 Safari/537.36',
    'Cookie': 'user_sid=119f0a1ce655491fa157a32ba539c08a; JSESSIONID=73F9F3D8E0FAF35155431C00F19B5CB2; SERVERID=daf947b71579cb2e324dcfdb35f0f984|1640786272|1640785824'
}


def sleep_random(const):
    randNum = round(random.random(), 3)
    interval = const + randNum
    print('sleep for ' + str(interval) + ' second...')
    time.sleep(interval)


def get_url_list_from_html(root_url):
    titles = []
    urls = []
    dates = []

    for i in range(MAX_PAGE - 1):
        sleep_random(1)
        html = root_url.format(i + 1)
        res = requests.get(html, headers)

        soup = BeautifulSoup(res.text, "lxml")
        links = get_links(soup)

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


def get_links(soup):
    # find all links by regex filter to all herf
    # links = soup.find_all(href=re.compile('content'))

    # find all links by div
    divs = soup.find_all(name='div', attrs={
        "class": "comprehensiveItem"
    })
    # TODO: div to links
    return divs


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
        res = doRequest(query, headers, i)
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


def doRequest(url, headers, index):
    # get
    # return requests.get(url, headers)
    # post
    headersReal = {
        'Host': 'search.zj.gov.cn',
        'Proxy-Connection': 'keep-alive',
        'Origin': 'http://search.zj.gov.cn',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Referer': 'http://search.zj.gov.cn/jsearchfront/search.do?websiteid=330000000000000&searchid=&pg=&p=1&tpl=1569&cateid=372&fbjg=&word=%E4%B9%A1%E6%9D%91&temporaryQ=&synonyms=&checkError=1&isContains=1&q=%E4%B9%A1%E6%9D%91&jgq=&eq=&begin=20180101&end=20211130&timetype=5&_cus_pq_ja_type=&pos=title&sortType=1',
        'X-Requested-With': 'XMLHttpRequest',
    }

    body = {
        'websiteid': 330000000000000,
        'pg': 10,
        'p': index + 1,
        'tpl': 1569,
        'cateid': 372,
        'word': '乡村',
        'checkError': 1,
        'isContains': 1,
        'q': '乡村',
        'begin': 20180101,
        'end': 20211130,
        'timetype': 5,
        'pos': 'title',
        'sortType': 1
    }
    data = parse.urlencode(body)
    headersReal.update(headers)
    return requests.request("POST", url, headers=headersReal, data=data)


def is_page_empty(json):
    # origin
    # contents = json['searchVO']['catMap']
    # keys = contents.keys()
    # remain = 0
    # for key in keys:
    #     remain += contents[key]['currentNum']
    #
    # return remain == 0
    return False


def get_url_lists_from_json(json):
    # get from origin json
    # return get_url_lists_from_origin_json(json)

    # get from html json
    titles = []
    urls = []
    dates = []
    categories = []
    contents = json['result']
    for item in contents:
        soup = BeautifulSoup(item, "lxml")
        data = soup.find(name='div', attrs={
            'class': 'titleWrapper'
        })
        catagory = data.find(name='input', attrs={
            'class': 'tagclass'
        })
        categories.append(catagory.get('value'))
        href = data.find(name='a')
        title = href.text.replace(' ', '').replace('\r', '').replace('\n', '')
        titles.append(title)
        m = re.match(r'.*url=([^&]*)&.*', href.get('href'))
        urls.append(parse.unquote(m.group(1)))
        sourceTime = soup.find(name='div', attrs={
            'class': 'sourceTime'
        })
        date = sourceTime.find(name='span', text=re.compile('时间'))
        m = re.match(r'.*:(.*)', date.text)
        dates.append(m.group(1))

    return {'title': titles, 'url': urls, 'date': dates, 'category': categories}


def get_url_lists_from_origin_json(json):
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
        sleep_random(0)
        get_article_from_single_page(getattr(row, 'date') + '_' + getattr(row, 'title'), getattr(row, 'url'),
                                     result_dir)


def get_article_from_single_page(title, url, result_dir):
    print('start to convert title[{}]...'.format(title))
    try:
        res = requests.get(url, headers)
    except:
        print('convert failed, continue')
        return

    res.encoding = res.apparent_encoding

    soup = BeautifulSoup(res.text, "lxml")
    if soup.select('.pages_content'):
        text = soup.select('.pages_content')[0].text
    elif soup.select('.b12c'):
        text = soup.select('.b12c')[0].text
    elif soup.select('.zoom'):
        text = soup.select('.zoom')[0].text
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
    # frame.to_csv('test_html_province_1.csv', encoding='utf_8_sig', index=False)

    # frame1 = get_url_list_from_query(QUERY_PATH_BASE)
    # frame1.to_csv('test_province_1.csv', encoding='utf_8_sig', index=False)

    frame2 = pd.read_csv('test_province_1.csv')
    get_articles(frame2, RESULT_DOCX_PATH)
