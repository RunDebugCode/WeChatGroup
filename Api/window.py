# -*- coding:UTF-8 -*-
import os
import tkinter
from tkinter import filedialog

import pywinauto
import win32api
import windnd

from Api import excel
from Api import wechat
from Config import settings
from Lib import common


def Center(width, height):
    screen_x = win32api.GetSystemMetrics(0)
    screen_y = win32api.GetSystemMetrics(1)
    return '%dx%d+%d+%d' % (width, height, (screen_x - width) / 2, (screen_y - height) / 2)


def GetPath(path):
    return str(path).split('\n')[0]


def change_input_box(target, value, edit=False):
    target.configure(state=tkinter.NORMAL)
    target.delete('1.0', tkinter.END)
    target.insert(tkinter.END, value)
    if not edit:
        target.configure(state=tkinter.DISABLED)


def check():
    msg = '抱歉，对比名单功能尚未实现！'
    common.warn_messagebox(msg)


class TK:
    def __init__(self):
        self.wait_time = int(settings.GetIniByKey('wait_time'))
        self.ico_path = common.GetFilePath(settings.ICO_PATH)

        self.root = tkinter.Tk()
        self.root.title(settings.TITLE2)
        self.root.geometry(Center(settings.WIDTH1, settings.HEIGHT1))
        self.root.resizable(width=False, height=False)
        self.root.iconbitmap(self.ico_path)

        # 绑定
        self.root.protocol('WM_DELETE_WINDOW', self.bye)
        self.root.bind('<Escape>', self.bye)
        self.root.bind('<F1>', common.help_message)
        self.root.bind('<F12>', common.author)

        MenuBar = tkinter.Menu(self.root)
        self.root.config(menu=MenuBar)

        SettingsMenu = tkinter.Menu(MenuBar, tearoff=False)
        SettingsMenu.add_command(label=settings.SETTING1, command=self.file_dialog1)
        SettingsMenu.add_command(label=settings.SETTING2, command=self.file_dialog2)
        SettingsMenu.add_command(label=settings.SETTING3, command=self.new_tk_to_change_time)

        AboutMenu = tkinter.Menu(MenuBar, tearoff=False)
        AboutMenu.add_command(label=settings.ABOUT1, command=common.help_message, accelerator='F1')
        AboutMenu.add_command(label=settings.ABOUT2, command=common.author, accelerator='F12')

        MenuBar.add_cascade(label=settings.MENU1, menu=SettingsMenu)
        MenuBar.add_cascade(label=settings.MENU2, menu=AboutMenu)

        self.label1 = tkinter.Label(self.root, text=settings.TEXT1)
        self.label1.place(x=settings.LABEL1X, y=settings.LABEL1Y)

        self.label2 = tkinter.Label(self.root, text=settings.TEXT2)
        self.label2.place(x=settings.LABEL2X, y=settings.LABEL2Y)

        self.label3 = tkinter.Label(self.root, text=settings.TEXT3)
        self.label3.place(x=settings.LABEL3X, y=settings.LABEL3Y)

        self.label4 = tkinter.Label(self.root, text=self.wait_time,
                                    bg=settings.BG, fg=settings.FG,
                                    width=int(5*settings.WEIGHT_X), height=1,
                                    font=settings.FONT,
                                    anchor='center')
        self.label4.place(x=settings.LABEL4X, y=settings.LABEL4Y)

        self.label5 = tkinter.Label(self.root, text=settings.TEXT4)
        self.label5.place(x=settings.LABEL5X, y=settings.LABEL5Y)

        self.input_box1 = tkinter.Text(self.root, width=int(35*settings.WEIGHT_X), height=1)
        change_input_box(self.input_box1, os.path.join(settings.GetDesktopPath(), settings.DEFAULTFILENAME))
        self.input_box1.place(x=settings.INPUT_BOX1X, y=settings.INPUT_BOX1Y)
        windnd.hook_dropfiles(self.input_box1, self.drag_file1)

        self.input_box2 = tkinter.Text(self.root, width=int(35*settings.WEIGHT_X), height=1)
        self.input_box2.configure(state=tkinter.DISABLED)
        self.input_box2.place(x=settings.INPUT_BOX2X, y=settings.INPUT_BOX2Y)
        windnd.hook_dropfiles(self.input_box2, self.drag_file2)

        self.button1 = tkinter.Button(self.root, text=settings.BUTTON1, command=self.file_dialog1)
        self.button1.place(x=settings.BUTTON1X, y=settings.BUTTON1Y)

        self.button2 = tkinter.Button(self.root, text=settings.BUTTON2, command=self.file_dialog2)
        self.button2.place(x=settings.BUTTON2X, y=settings.BUTTON2Y)

        self.button3 = tkinter.Button(self.root, text=settings.BUTTON3, command=self.get_name)
        self.button3.place(x=settings.BUTTON3X, y=settings.BUTTON3Y)

        self.button4 = tkinter.Button(self.root, text=settings.BUTTON4, command=check)
        self.button4.place(x=settings.BUTTON4X, y=settings.BUTTON4Y)

        self.root.mainloop()

    def drag_file1(self, files):
        msg = '\n'.join((item.decode('gbk') for item in files)).split('\n')
        if len(msg) == 1:
            path = msg[0]
            name, suffix = os.path.splitext(path)
            if suffix == '':
                path += '.xlsx'
                change_input_box(self.input_box1, path)
        else:
            common.warn_messagebox('一次只能拖入一个文件')

    def drag_file2(self, files):
        msg = '\n'.join((item.decode('gbk') for item in files)).split('\n')
        if len(msg) == 1:
            path = msg[0]
            name, suffix = os.path.splitext(path)
            if suffix == '':
                path += '.xlsx'
                change_input_box(self.input_box1, path)
        else:
            common.warn_messagebox('一次只能拖入一个文件')

    def file_dialog1(self):
        r = filedialog.asksaveasfilename(title='保存...',
                                         initialdir=settings.GetDesktopPath(),
                                         initialfile='name.xlsx',
                                         filetypes=[('Excel 文件', '*.xls *.xlsx'), ('All files', '*')])
        if r != '':
            name, suffix = os.path.splitext(r)
            if suffix == '':
                r += '.xlsx'
                change_input_box(self.input_box1, r)
            change_input_box(self.input_box1, r)

    def file_dialog2(self):
        r = filedialog.askopenfilename(title='打开',
                                       initialdir=settings.GetDesktopPath(),
                                       multiple=False,
                                       filetypes=[('Excel 文件', '*.xls *.xlsx'), ('All files', '*')])
        if r != '':
            name, suffix = os.path.splitext(r)
            if suffix == '':
                r += '.xlsx'
                change_input_box(self.input_box1, r)
            change_input_box(self.input_box1, r)

    def countdown(self, seconds):
        if seconds >= 0:
            self.label4.config(text=str(seconds))
            seconds -= 1
            self.root.after(1000, self.countdown, seconds)
        else:
            self.label4.config(text='处理中')
            self.root.attributes('-topmost', 0)
            self.root.update()

            wechat_pid = wechat.GetWechatPid()
            if wechat_pid:
                try:
                    name, data = wechat.GetNameList(wechat_pid)
                    path = GetPath(self.input_box1.get('1.0', tkinter.END))
                    excel.write(name, data, path)
                    self.root.attributes('-topmost', 0)
                    self.label4.config(text='完成')
                    msg = f'在群 {name} 共获得 {len(data)} 人的名单，\n' \
                          f'名单已经保存至：{path}。\n\n' \
                          '注：如果获得人数与实际人数不符，请确认是否打开查看更多。'
                    common.ok_messagebox(msg)
                except pywinauto.findwindows.ElementNotFoundError as e:
                    print(e)
                    common.warn_messagebox('检测到没有打开 聊天信息 窗口，请重新尝试！')
            else:
                common.warn_messagebox('检测到没有打开微信程序，请先登录微信！')

            self.button3.config(state=tkinter.NORMAL)

    def get_name(self):
        path = GetPath(self.input_box1.get('1.0', tkinter.END))
        if not os.path.exists(path):
            if not os.path.isdir(path):
                suffix = os.path.splitext(path)[1]
                if suffix in ['.xls', '.xlsx']:
                    count = self.wait_time
                    self.label4.config(text=count)      # 初始化 label4 为等待时间
                    self.button3.config(state=tkinter.DISABLED)     # 将按钮改为灰色
                    self.root.attributes('-topmost', 1)

                    self.countdown(count)
                else:
                    common.warn_messagebox(f'{suffix} 不是受支持的后缀名！')
            else:
                common.warn_messagebox(f'{path} 不是文件！')
        else:
            common.warn_messagebox(f'{path} 已经存在，再次操作会覆写数据，请重新点击浏览选择文件！')

    def new_tk_to_change_time(self):
        self.wait_time = int(settings.GetIniByKey('wait_time'))

        self.branch = tkinter.Tk()
        self.branch.title(settings.TITLE3)
        self.branch.geometry(Center(settings.WIDTH2, settings.HEIGHT2))
        self.branch.resizable(width=False, height=False)
        self.branch.iconbitmap(self.ico_path)

        label5 = tkinter.Label(self.branch, text=settings.TEXT5)
        label5.place(x=settings.LABEL5X, y=settings.LABEL5Y)

        label6 = tkinter.Label(self.branch, text=settings.TEXT4)
        label6.place(x=settings.LABEL6X, y=settings.LABEL6Y)

        self.input_box3 = tkinter.Text(self.branch, width=int(18*settings.WEIGHT_X), height=1)
        change_input_box(self.input_box3, self.wait_time, True)
        self.input_box3.place(x=settings.INPUT_BOX3X, y=settings.INPUT_BOX3Y)

        button5 = tkinter.Button(self.branch, text=settings.BUTTON5, command=self.save_time)
        button5.place(x=settings.BUTTON5X, y=settings.BUTTON5Y)

        self.root.mainloop()

    def save_time(self):
        time = GetPath(self.input_box3.get('1.0', tkinter.END))
        settings.WriteInIni('settings', 'wait_time', time)
        self.branch.destroy()
        self.wait_time = self.wait_time = int(settings.GetIniByKey('wait_time'))
        self.label4.config(text=self.wait_time)
        self.root.update()

    def bye(self, ev=None):
        print(ev)
        self.root.destroy()
        common.ok_messagebox('感谢使用')


def main():
    TK()


if __name__ == '__main__':
    main()
