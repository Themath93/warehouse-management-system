from asyncio.log import logger
import logging
from spark_infra.util import cal_std_day


def get_logger(name):
    """
    log를 키록하는 함수
    home 디렉터리에서 오늘날짜.log파일에 발생한 로그들을 입력한다.
    """
    co_logger = logging.getLogger(name)
    handler = logging.FileHandler('./log/'+cal_std_day(0)+'.log')
    co_logger.addHandler(handler)
    return co_logger