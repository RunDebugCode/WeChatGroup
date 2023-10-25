# -*- coding:UTF-8 -*-
"""
功能模块
"""
import datetime
import os
import sys
import time

import win32api
import win32con

from Config import settings


def sleep(n):
    time.sleep(n)


def now_time():
    now = datetime.datetime.now()
    return now.strftime('%Y/%m/%d %H:%M:%S')


def ok_messagebox(msg):
    win32api.MessageBox(None, str(msg), 'Tips:', win32con.MB_OK)


def warn_messagebox(msg):
    win32api.MessageBox(None, str(msg), 'Error:', win32con.MB_ICONERROR)


def author(ev=None):
    print(ev)
    msg = f'名字   ：  {settings.TITLE2}\n' \
          '作者    ：  SOUL\n' \
          f'Now   ：  {now_time()}\n' \
          '日期    ：  2020年11月22日 15点43分\n' \
          'Power  ：  Python 3.6.4  Pycharm 2021.1.3(Community Edition)'
    ok_messagebox(msg)


def help_message(ev=None):
    print(ev)
    msg = '1、登录微信\n' \
          '2、点击第一个浏览选择获取的名单保存的位置\n' \
          '3、点击获取，然后在规定的秒数内依次打开：\n' \
          '\t（1）打开目标群聊\n' \
          '\t（2）点击右上角聊天信息\n' \
          '\t（3）点击 [查看更多 ⌵]（注：如果没有可忽略）\n' \
          '4、等待获取成功'
    ok_messagebox(msg)


def ResourcePath(relative_path):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def GetFilePath(file):
    return ResourcePath(file)


def main():
    help_message()


if __name__ == '__main__':
    main()
