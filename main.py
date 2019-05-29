#coding=utf-8
# 网页自动测试脚本
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
import os
import xlrd
import logging


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 日志配置
logger = logging.getLogger(__name__)
format = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
# 文件日志
handler = logging.FileHandler(os.path.join(BASE_DIR, "log"), encoding="utf-8")
handler.setFormatter(format)
handler.setLevel(logging.ERROR)
logger.addHandler(handler)
# 控制台日志
console = logging.StreamHandler()
console.setLevel(logging.INFO)
console.setFormatter(format)
logger.addHandler(console)
logger.setLevel(logging.INFO)


def load_config():
    # 从excel中加载配置
    res = {}
    config = os.path.join(BASE_DIR, "配置.xls")
    excel = xlrd.open_workbook(config)
    sheet = excel.sheet_by_index(0)

    for i in range(1, sheet.nrows):
        url, num = sheet.row_values(i)
        logger.info("浏览%s次%s", int(num), url)
        res[url] = int(num)

    return res


def add_pv():
    stat = {}
    opt = Options()
    opt.add_argument("--incognito")
    opt.add_argument("--headless")
    opt.add_argument('--log-level=3')
    opt.add_argument('--disable-extensions')
    opt.add_argument('test-type')
    targets = load_config()
    for url, num in targets.items():
        logger.info("--------------------------------")
        logger.info("正在浏览%s", url)
        stat.setdefault(url, [0, 0])
        for i in range(num):
            chrome = webdriver.Chrome(executable_path=os.path.join(BASE_DIR, "chromedriver.exe"), options=opt)
            chrome.set_page_load_timeout(9)
            try:
                chrome.get(url)
                stat[url][0] += 1
            except TimeoutException as e:
                logger.warning('请求页面超时，继续下一次请求')
                stat[url][1] += 1
            # 关闭浏览器
            chrome.close()

            # 每10次输出一次
            x = i + 1
            if x % 10 == 0:
                logger.info("已完成%s次浏览", x)
    logger.info("**********************************************")
    logger.info("全部任务已完成")
    for url, (succ, fail) in stat.items():
        logger.info("浏览%s次%s", succ+fail, url)
        logger.info("成功: %s次", succ)
        logger.info("失败: %s次", fail)
        logger.info("----------------------------------------------------")
    input("按回车键退出程序")


if __name__ == "__main__":
    try:
        add_pv()
    except Exception as e:
        logger.exception(e)


