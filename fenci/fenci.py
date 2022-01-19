# encoding=utf-8
import os
import time

import docx
import jieba
import jieba.analyse
from matplotlib import pyplot as plt
from wordcloud import WordCloud

FILE_DIR = [
    # 'D:/Data/code/policy-analysis/input-files/zywj/',
    # 'D:/Data/code/policy-analysis/input-files/qtwj/',
    # 'D:/Data/code/policy-analysis/input-files/cebx/',
    # 'D:/Data/code/policy-analysis/input-files/xiaolunwen/'
    # 'D:/Data/code/policy-analysis/input-files/guowuyuan/'
    'D:/Data/code/policy-analysis/input-files/zhejiangsheng/'
]
FILE_SUFFIX = '.docx'

STOP_WORDS_FLAG = True
STOP_WORDS_DIR = 'stopwords/'
STOP_WORDS_NAME = 'merge_stopwords.txt'
STOP_WORDS_OUT_NAME = 'all_out_stopwords.txt'
STOP_WORDS_MINE_NAME = 'mine_stopwords.txt'

CUSTOM_DICT_FLAG = False
DICT_DIR = 'dict/'
DICT_NAME = 'custom.txt'

RESULT_DIR = 'result/'
FENCI_NAME = 'fenci-result'
STAT_NAME = 'stat-result'
DATA_RESULT_SUFFIX = '.csv'
CLOUD_NAME = 'fenci-cloud'
PIC_RESULT_SUFFIX = '.jpg'

FONT_NAME = 'fonts/ms-yahei.ttf'

ANALYSE_TOPK = 30
ANALYSE_POS = ('n')

STAT_TOPK = 100

# jieba result
result = {}


def read_docx_file(file_path):
    doc = docx.Document(file_path)
    word_str = ''
    for para in doc.paragraphs:
        word_str += para.text
    return word_str


def merge_stopwords(file_dir, file_merge):
    if os.path.exists(file_merge):
        os.remove(file_merge)
    with open(file_merge, 'w', encoding='utf-8') as f:
        files = os.listdir(file_dir)
        for file in files:
            if os.path.isdir(file):
                continue
            if os.path.splitext(file)[1] != '.txt':
                continue
            with open(file_dir + file, 'r', encoding='utf-8') as f_sub:
                f.writelines(f_sub.readlines())


def file_words_analyse(str):
    # TF-IDF method
    # stop words
    if STOP_WORDS_FLAG:
        jieba.analyse.set_stop_words(STOP_WORDS_DIR + STOP_WORDS_NAME)
    # custom idf dict
    if CUSTOM_DICT_FLAG:
        jieba.analyse.set_idf_path(DICT_DIR + DICT_NAME)
    tags = jieba.analyse.extract_tags(str,
                                      topK=ANALYSE_TOPK,
                                      withWeight=True,
                                      allowPOS=ANALYSE_POS)
    # TestRank method
    # tags = jieba.analyse.textrank(str, topK=ANALYSE_TOPK, withWeight=True)
    # print(tags)
    return tags


def analyse_files_batch(dir_list, suffix):
    for dir in dir_list:
        analyse_files(dir, suffix)


def analyse_files(dir, suffix):
    files = os.listdir(dir)
    for file in files:
        if os.path.isdir(file):
            continue
        if file.find('~$') != -1:
            continue
        if os.path.splitext(file)[1] != suffix:
            continue
        print('------start analyse file:' + file)
        tags = file_words_analyse(read_docx_file(dir + file))
        result[file] = tags


def generate_and_check_result_name(result_type, suffix):
    # generate result name
    result_file_name = RESULT_DIR + result_type + start_time_str
    if STOP_WORDS_FLAG:
        result_file_name += '_stop'
    if CUSTOM_DICT_FLAG:
        result_file_name += '_customD'
    result_file_name += suffix
    # check exist and remove
    if os.path.exists(result_file_name):
        os.remove(result_file_name)
    return result_file_name


def convert_to_freq_dict():
    words_freq = {}
    for key in result.keys():
        for item in result[key]:
            if item[0] in words_freq:
                words_freq[item[0]] += item[1]
            else:
                words_freq[item[0]] = item[1]
    return words_freq


def generate_fenci_csv():
    result_file_name = generate_and_check_result_name(FENCI_NAME, DATA_RESULT_SUFFIX)
    with open(result_file_name, 'w', encoding='utf-8-sig') as result_file:
        for key in result.keys():
            str_list_word = []
            str_list_rate = []
            str_list_word.append(key)
            str_list_rate.append(key)
            for item in result[key]:
                str_list_word.append(item[0])
                str_list_rate.append(str(item[1]))
            result_file.writelines(",".join(str_list_word))
            result_file.writelines('\n')
            result_file.writelines(",".join(str_list_rate))
            result_file.writelines('\n')


def generate_stat_csv():
    result_file_name = generate_and_check_result_name(STAT_NAME, DATA_RESULT_SUFFIX)
    words_freq = convert_to_freq_dict()
    with open(result_file_name, 'w', encoding='utf-8-sig') as result_file:
        sorted_words_stat = sorted(words_freq.items(), key=lambda x: x[1], reverse=True)
        sorted_words_stat_topk = sorted_words_stat[:STAT_TOPK]
        word_list = [x for x, _ in sorted_words_stat_topk]
        stat_list = [str(y) for _, y in sorted_words_stat_topk]
        result_file.writelines(",".join(word_list))
        result_file.writelines('\n')
        result_file.writelines(",".join(stat_list))
        result_file.writelines('\n')


def generate_cloud():
    words_freq = convert_to_freq_dict()
    pic_result_name = generate_and_check_result_name(CLOUD_NAME, PIC_RESULT_SUFFIX)
    cloud = WordCloud(
        font_path=FONT_NAME,
        background_color='white',
        max_words=150,
        width=1000,
        height=1000
    ).generate_from_frequencies(words_freq)
    cloud.to_file(pic_result_name)
    plt.imshow(cloud)
    plt.axis("off")
    plt.show()


def generate_results():
    generate_fenci_csv()
    generate_stat_csv()
    generate_cloud()


if __name__ == '__main__':
    start_time_str = time.strftime("_%m%d_%H%M%S", time.localtime())
    merge_stopwords(STOP_WORDS_DIR, STOP_WORDS_DIR + STOP_WORDS_NAME)
    analyse_files_batch(FILE_DIR, FILE_SUFFIX)
    generate_results()
