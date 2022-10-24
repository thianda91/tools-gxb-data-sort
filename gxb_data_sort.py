#!python3

from xdaLibs import iniconfig
from colorama import Fore, Back, init as init_colorama
from msoffcrypto import OfficeFile
import io
import re
from pandas import read_excel, concat, DataFrame
import time
from pathlib import Path
import warnings
import os
# import openpyxl
# import numpy
import traceback
from sys import exit

__title__ = ''.join(chr(x) for x in [24037, 20449, 37096, 32593, 31449, 39640, 21361, 23458, 25143, 25968, 25454, 25972, 29702, 23567, 24037, 20855])
__version__ = 'v0.7'
__author__ = ''.join(chr(x) for x in [20110, 26174, 36798, 46, 38081, 23725])

DEBUG_FILE = 'debug_log.txt'

# concat() DataFrame 后再 set_index() 会打印 FutureWarning
# 此方法屏蔽此警告
warnings.filterwarnings('ignore')

# 1. 读配置，确认台账是否存在，确认要处理的文件是否存在
# 2. 读取台账，加载全部信息
# 3. 遍历文件，处理并补充到台账
start_time = time.time()
init_colorama(autoreset=True)
configfile = 'config.ini'
if Path(configfile).is_file():
    config = iniconfig.IniConfig(configfile)
else:
    print('配置文件“ {} ”不存在，无法运行。'.format(configfile))
    exit()

COL_NAME_TIME = config.get('common', 'col_name_time')
PROVINCE = config.get('common', 'province')
COL_NAMES = config.get('common', 'col_names').split('\t')
keywords = ['index', 'first_time', 'last_time', 'count']
INDEX = config.get('common', keywords[0])
FIRST_TIME = config.get('common', keywords[1])
LAST_TIME = config.get('common', keywords[2])
COUNT = config.get('common', keywords[3])
READ_SHEET_NAME = config.get('common', 'read_sheet_name')
BOOK_SHEET_NAME = config.get('common', 'book_sheet_name')
BACKUP_MODE = config.get('common', 'backup_mode')

def check_config() -> bool:
    '''
    运行前检查配置文件参数是否存在
    '''
    flag = True
    for i, v in enumerate([INDEX, FIRST_TIME, LAST_TIME, COUNT]):
        if not v in COL_NAMES:
            print(Fore.LIGHTRED_EX + '无法正确运行，请检查配置文件 col_names 字段是否包含 {} 字段 （"{}"）'.format(keywords[i], v))
            flag = False
    return flag

def open_excel_with_key(file: str, key: str) -> DataFrame:
    '''
    打开带密码的 excel 文件
    '''
    if key == '0' or key == '':
        # 没密码
        io_file = file
    else:
        # 有密码
        io_file = io.BytesIO()
        if key == '1':
            # 密码需要自动识别
            key = re.search(r'[0-9]+', file).group()[:4]

        with open(file, 'rb') as f:
            excel = OfficeFile(f)
            excel.load_key(key)
            excel.decrypt(io_file)
    i_df = read_excel(io=io_file, sheet_name=READ_SHEET_NAME, dtype={COL_NAME_TIME: str})
    del io_file
    return i_df


def init_table_from_empty(filename: str, col_names: list) -> None:
    '''
    台账初始化
    按列名创建一个空台账
    '''
    df = DataFrame({}, columns=col_names)
    df.to_excel(filename, index=False)


def self_compare(book: DataFrame) -> DataFrame:
    '''
    dataFrame 自比较
    用户号码 去重 与 统计
    '''
    book.sort_values(by=FIRST_TIME, inplace=True)
    # 统计重复次数
    dup_series = book[INDEX].value_counts().rename(COUNT)
    # 去重，保留最近时间
    book_last = book.drop_duplicates(subset=INDEX, keep='last')[[INDEX, FIRST_TIME]]
    book_last.rename(columns={FIRST_TIME: LAST_TIME}, inplace=True)
    # 去重，保留最早时间
    book = book.drop_duplicates(subset=INDEX, keep='first')
    book = book.set_index(INDEX)
    # book[LAST_TIME] = book[FIRST_TIME]
    # 更新最新时间
    book.update(book_last.set_index(INDEX))
    # 更新重复次数
    book.update(dup_series)
    return book


def df_compare(book: DataFrame, df: DataFrame) -> DataFrame:
    '''
    dataFrame 互比较
    在 2 个 dataFrame 用户号码 去重 与 统计
    '''
    # 先自去重，返回结果已设置 index
    df = self_compare(df)
    # 统计是否包含（是否重复）
    df['include'] = df.index.isin(book.index)
    # 包含的更新，次数相加
    df_dup = df[df['include']]
    book[COUNT] = book[COUNT].add(df_dup[COUNT], fill_value=0)
    book.update(df_dup[LAST_TIME])
    # 不包含的新增
    _df = df[~df['include']].drop(labels='include', axis=1)
    book = concat([book, _df])
    return book


def backup_book(book_file):
    '''
    先复制，备份台账文件
    '''
    _name, _suffix = book_file.split('.')
    _time = '{:.0f}'.format(time.time())
    if '1' == BACKUP_MODE:
        _bak_name = '{}_bak-{}.{}'.format(_name, _time, _suffix)
    else:
        _bak_name = '{}_bak.{}'.format( _name, _suffix)
    _script = 'copy {} {} /y'.format(book_file, _bak_name)
    os.system(_script)
    print('>>> 台账已备份为: ', Fore.LIGHTBLUE_EX + _bak_name)


def main():
    print(Fore.LIGHTBLUE_EX + '╭---------------------------------------╮')
    print(Fore.CYAN + '| {} {} |'.format(__title__, __version__))
    print('|          **** 作者： {} **** |'.format(__author__))
    print(Fore.LIGHTBLUE_EX + '└---------------------------------------┘')
    if not check_config():
        exit('退出运行..')
    book_file = config.get('common', 'path')
    is_first_run = False
    if not Path(book_file).is_file():
        # 配置文件不存在
        # print(Fore.RED + '台账文件路径不存在，请修改配置文件 {}，然后重新运行'.format(configfile))
        print(Fore.RED + '台账文件路径不存在，将自动生成：{}'.format(book_file))
        init_table_from_empty(book_file, COL_NAMES)
        is_first_run = True

    book_df = read_excel(io=book_file, sheet_name=BOOK_SHEET_NAME, index_col=INDEX, dtype={FIRST_TIME: str, LAST_TIME: str})
    # book_list = book_df[book_df['省份'] == PROVINCE][INDEX].to_list()
    filename_keyword = config.get('common', 'import_filename')
    filenames = [f for f in Path.cwd().glob('*{}*'.format(filename_keyword)) if Path(f).is_file() if '~$' not in str(f)]
    if filenames == []:
        print(Fore.RED + '未识别到有效的工信部导出文件，请修改配置文件中的 “import_filename” 并重试。')
        input()
        exit()
    df_list = []
    print(Fore.MAGENTA + '正在运行，请等待...')
    # 循环读取，汇总到 list
    for f in filenames:
        key = config.get('common', 'key')
        i_df = open_excel_with_key(str(f), key)
        i_df_filter = i_df[i_df['省份'] == PROVINCE]
        df_list.append(i_df_filter)

        print('')
    # 连接成一个 dataFrame
    df = concat(objs=df_list, ignore_index=True)
    # 修改导入数据的列名，与台账的列名一致
    df.rename(columns={COL_NAME_TIME: FIRST_TIME}, inplace=True)
    if book_df.empty:
        book_df = concat([book_df, df])
        book_df = self_compare(book_df)
        print(Fore.MAGENTA + Back.LIGHTYELLOW_EX + '>>>>> self_compare')
    else:
        df[LAST_TIME] = ''
        df[COUNT] = None
        book_df = df_compare(book_df, df)
        print(Fore.MAGENTA + Back.LIGHTYELLOW_EX + '>>>>> df_compare')

    if not is_first_run:
        backup_book(book_file)

    book_df.to_excel(excel_writer=book_file, sheet_name=BOOK_SHEET_NAME)
    # book_df.to_excel(excel_writer=book_file, sheet_name=BOOK_SHEET_NAME, columns=COL_NAMES)

    end_time = time.time()
    run_time = end_time - start_time
    print(Fore.GREEN + '运行成功结束. 总用时: {:.2f} s'.format(run_time))


if __name__ == '__main__':
    if time.time() > time.mktime(time.strptime('20221130', '%Y%m%d')):
        print(Fore.RED + '**** 本工具为旧版本，已无法使用。')
        input()
        exit()
    try:
        main()
    except Exception as err:
        print(err)
        traceback.print_exc(file=open(DEBUG_FILE, 'w'))
        print(Fore.LIGHTRED_EX + '** 出错啦。请反馈目录中的文件：', DEBUG_FILE)
        input('按回车退出')
    exit()
