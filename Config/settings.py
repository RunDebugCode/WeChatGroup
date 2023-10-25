# -*- coding:UTF-8 -*-
import configparser
import winreg

import win32api

from Lib import common

# 分辨率权重计算
BASIC_X = 1920
BASIC_Y = 1080
X = win32api.GetSystemMetrics(0)
Y = win32api.GetSystemMetrics(1)

WEIGHT_X = X / BASIC_X
WEIGHT_Y = Y / BASIC_Y
WIDTH1 = 500 * WEIGHT_X
HEIGHT1 = 200 * WEIGHT_Y
WIDTH2 = 350 * WEIGHT_X
HEIGHT2 = 100 * WEIGHT_Y

LABEL1X = 20 * WEIGHT_X
LABEL1Y = 10 * WEIGHT_Y
LABEL2X = 20 * WEIGHT_X
LABEL2Y = 60 * WEIGHT_Y
LABEL3X = 20 * WEIGHT_X
LABEL3Y = 100 * WEIGHT_Y
LABEL4X = 100 * WEIGHT_X
LABEL4Y = 95 * WEIGHT_Y
LABEL5X = 280 * WEIGHT_X
LABEL5Y = 140 * WEIGHT_Y
LABEL6X = 250 * WEIGHT_X
LABEL6Y = 0 * WEIGHT_Y

INPUT_BOX1X = 80 * WEIGHT_X
INPUT_BOX1Y = 10 * WEIGHT_Y
INPUT_BOX2X = 80 * WEIGHT_X
INPUT_BOX2Y = 60 * WEIGHT_Y
INPUT_BOX3X = 80 * WEIGHT_X
INPUT_BOX3Y = 5 * WEIGHT_Y

BUTTON1X = 420 * WEIGHT_X
BUTTON1Y = 10 * WEIGHT_Y
BUTTON2X = 420 * WEIGHT_X
BUTTON2Y = 60 * WEIGHT_Y
BUTTON3X = 340 * WEIGHT_X
BUTTON3Y = 130 * WEIGHT_Y
BUTTON4X = 420 * WEIGHT_X
BUTTON4Y = 130 * WEIGHT_Y
BUTTON5X = 80 * WEIGHT_X
BUTTON5Y = 60 * WEIGHT_Y

BG = 'white'
FG = 'red'
FONT = ('宋体', 40)
ICO_PATH = r'Files\1.ico'
TITLE = [['序号', '群昵称', '微信名(备注名)']]
TITLE1 = '聊天成员'
TITLE2 = '微信群昵称获取'
TITLE3 = '修改等待时间'
DEFAULTFILENAME = 'name.xlsx'
WECHAT_EXE_NAME = 'WeChat.exe'
CLASS_NAME = 'WeChatMainWndForPC'
INI_PATH = r'Files\settings.ini'
DESKTOP_PATH_IN_REGEDIT = r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'

SETTING1 = '路径1'
SETTING2 = '路径2'
SETTING3 = '等待时间'
ABOUT1 = '帮助'
ABOUT2 = '作者'
MENU1 = '配置'
MENU2 = '关于'
TEXT1 = '保存至：'
TEXT2 = '名单：'
TEXT3 = '时间还剩：'
TEXT4 = '秒'
TEXT5 = '等待时间：'
BUTTON1 = BUTTON2 = '浏览'
BUTTON3 = '获取'
BUTTON4 = '对比'
BUTTON5 = '确认修改'


def GetDesktopPath():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, DESKTOP_PATH_IN_REGEDIT)
    return winreg.QueryValueEx(key, 'Desktop')[0]


def GetIniByKey(key, section='settings'):
    path = common.GetFilePath(INI_PATH)
    config = configparser.ConfigParser()
    config.read(path)
    return config.get(section, key)


def WriteInIni(section, key, value):
    path = common.GetFilePath(INI_PATH)
    config = configparser.ConfigParser()
    config.add_section(section)
    config.set(section, key, value)
    config.write(open(path, 'w'))


def main():
    pass


if __name__ == '__main__':
    main()
