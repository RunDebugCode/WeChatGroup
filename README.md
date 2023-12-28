# 项目介绍
    项目名称:   一键获取群成员昵称
    项目描述:   获取微信群聊中成员的群昵称
    项目环境:   pycharm & python & Windows 11
    作者   :   SOUL


# 环境依赖
    推荐使用以下配置：

    Python :    3.6.4
    Pycharm:    2021.1.3 (Community Edition)
    Windows:    11
    WeChat :    3.9.7.29


# 第三方库依赖
    openpyxl
    pywinauto
    pywin32
    windnd


# 使用说明
* 登录微信
* 点击第一个浏览选择获取的名单保存的位置
* 点击获取，然后在规定的秒数内依次打开：
*      （1）打开目标群聊
       （2）点击右上角聊天信息
       （3）点击 [查看更多 ⌵]（注：如果没有可忽略）
* 4、等待获取成功


# 项目结构
    PYCHARMPROJECTS\WECHATGROUP
    │  .gitignore
    │  LICENSE
    │  README.MD
    │  setup.py
    │  setup.spec
    │
    ├─Api
    │     excel.py
    │     wechat.py
    │     window.py
    │
    ├─Config
    │     settings.py
    │
    ├─Core
    │     src.py
    │
    ├─Files
    │     1.ico
    │     settings.ini
    │
    └─Lib
          common.py
          exe2icon.py


# 版本内容更新
###### v1.0.0
* 基础功能实现
###### v2.0.0
* 修复了已知 bug
* 实现了动态分辨率调整
