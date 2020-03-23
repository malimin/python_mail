import logging
import warnings
import time
import os

def get_logger():
    try:
        os.mkdir("./log")
    except:
        pass

    local_time = time.localtime(time.time())
    nian = time.strftime('%Y', local_time)
    yue = time.strftime('%m', local_time)
    ri = time.strftime('%d', local_time)
    shi = time.strftime('%H', local_time)
    fen = time.strftime('%M', local_time)
    miao = time.strftime('%S', local_time)

    # warnings.filterwarnings('ignore')

    log_dir = './log/{}-{}-{}-record.log'.format(nian, yue, ri)


    fh = logging.FileHandler(log_dir, encoding='utf-8', mode="a")  # 创建一个文件流并设置编码utf8
    logger = logging.getLogger()  # 获得一个logger对象，默认是root
    logger.setLevel(logging.DEBUG)  # 设置最低等级debug
    fm = logging.Formatter("%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s ：%(message)s",
                           datefmt='%Y-%m-%d %A %H:%M:%S')  # 设置日志格式

    logger.addHandler(fh)  # 把文件流添加进来，流向写入到文件
    fh.setFormatter(fm)  # 把文件流添加写入格式
    return logger
get_logger()

