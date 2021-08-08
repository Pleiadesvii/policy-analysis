# encoding=utf-8
import os
import jieba
import jieba.analyse
import docx
import time


FILE_DIR = ['D:/Data/code/policy-fenci/zywj/',
            'D:/Data/code/policy-fenci/qtwj/',
            'D:/Data/code/policy-fenci/cebx/']
FILE_SUFFIX = '.docx'

RESULT_DIR = 'D:/Data/code/policy-fenci/'
RESULT_NAME = 'fenci-result'
RESULT_SUFFIX = '.csv'

ANALYSE_TOPK = 10
ANALYSE_POS = ('n')

result = {}


def read_docx_file(file_path):
    doc = docx.Document(file_path)
    word_str = ''
    for para in doc.paragraphs:
        word_str += para.text
    return word_str


def file_words_analyse(str):
    tags = jieba.analyse.extract_tags(
        str, topK=ANALYSE_TOPK, withWeight=True, allowPOS=ANALYSE_POS)
    print(tags)
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


def generate_csv():
    time_str = time.strftime("_%m%d_%H%M%S", time.localtime())
    result_file_name = RESULT_DIR + RESULT_NAME + time_str + RESULT_SUFFIX
    if os.path.exists(result_file_name):
        os.remove(result_file_name)
    with open(result_file_name, 'w', encoding='utf-8-sig') as result_file:
        for key in result.keys():
            str_list = []
            str_list.append(key)
            for item in result[key]:
                str_list.append(item[0])
                str_list.append(str(item[1]))
            result_file.writelines(",".join(str_list))
            result_file.writelines('\n')


if __name__ == '__main__':
    analyse_files_batch(FILE_DIR, FILE_SUFFIX)
    print('------all result')
    print(result)
    generate_csv()
