# -*- coding:UTF-8 -*-
import psutil
from pywinauto.application import Application

from Config import settings


def GetWechatPid():
    pidS = psutil.pids()
    for pid in pidS:
        p = psutil.Process(pid)
        if p.name() == settings.WECHAT_EXE_NAME:
            return pid
    return None


def GetNameList(pid: int):
    """
    通过pid获取 群名称 和 微信当前打开群的成员的群昵称和微信名

    :param pid: 微信进程ID
    :return: (group_title, [name, nickname])
    """
    name_nick = []

    app = Application(backend='uia').connect(process=pid)
    main_dialog = app.window(class_name=settings.CLASS_NAME)
    # group_name = main_dialog.Edit3.window_text()  # 通过 Edit3 标签获取群聊名称 放在这不可用，不知道为什么
    chat_list = main_dialog.child_window(control_type='List', title=settings.TITLE1)

    for i in chat_list.items():
        p = i.descendants()
        if p and len(p) > 5:
            if p[5].texts() and p[5].texts()[0].strip() != '' and \
                    (p[5].texts()[0].strip() != '添加' and p[5].texts()[0].strip() != '移出'):
                name_nick.append([p[5].texts()[0].strip(), p[3].texts()[0].strip()])

    # main_dialog.wait('visible')
    # main_dialog.print_control_identifiers(depth=None, filename=None)
    group_name = main_dialog.Edit3.window_text()    # 通过 Edit3 标签获取群聊名称

    return group_name, name_nick


def main():
    print(GetWechatPid())


if __name__ == '__main__':
    main()
