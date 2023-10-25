"""
use ctypes to get image from exe/dll
ref:
  https://www.zhihu.com/question/425053417
  https://www.cnblogs.com/ibingshan/p/11057390.html
  mss.windows (mss grab library)
"""
import platform, os

system = platform.system().lower()
if system != "windows" and os.environ["IGNORE_SYSTEMCHECK"] != "True":
    raise Exception("GETICON is only avaliable on Windows!")

import ctypes
from ctypes.wintypes import HICON, LPCSTR, UINT, INT

ExIcon = ctypes.windll.user32.PrivateExtractIconsA
DesIcon = ctypes.windll.user32.DestroyIcon

from ctypes import POINTER, Structure, WINFUNCTYPE, c_void_p
from ctypes.wintypes import (
    BOOL,
    DOUBLE,
    DWORD,
    HBITMAP,
    HDC,
    HGDIOBJ,
    HWND,
    INT,
    LONG,
    LPARAM,
    RECT,
    UINT,
    WORD,
)

SRCCOPY = 0x00CC0020

user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32


class BITMAPINFOHEADER(Structure):
    " From mss.windows "

    _fields_ = [
        ("biSize", DWORD),
        ("biWidth", LONG),
        ("biHeight", LONG),
        ("biPlanes", WORD),
        ("biBitCount", WORD),
        ("biCompression", DWORD),
        ("biSizeImage", DWORD),
        ("biXPelsPerMeter", LONG),
        ("biYPelsPerMeter", LONG),
        ("biClrUsed", DWORD),
        ("biClrImportant", DWORD),
    ]


class BITMAPINFO(Structure):
    " From mss.windows "

    _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", DWORD * 3)]


def rgb(raw, width=32, height=32):
    # From mss.screenshot
    rgb = bytearray(height * width * 3)
    rgb[0::3] = raw[2::4]
    rgb[1::3] = raw[1::4]
    rgb[2::3] = raw[0::4]
    return bytes(rgb)


def get_raw_data(path, index=0, size=32):
    path = os.path.abspath(path)
    # From zhihu
    width = height = size

    ExIcon.restype = UINT

    icon_total_count = ExIcon(path.encode(), 0, 0, 0, None, None, 0, 0)

    ExIcon.argtypes = [
        LPCSTR, INT, INT, INT,
        ctypes.POINTER(HICON * icon_total_count),
        ctypes.POINTER(UINT * icon_total_count), UINT,
        UINT
    ]

    hIconArray = HICON * icon_total_count
    hicons = hIconArray()

    p_hicons = ctypes.pointer(hicons)

    IDArray = UINT * icon_total_count
    ids = IDArray()
    p_ids = ctypes.pointer(ids)

    success_count = ExIcon(path.encode(), 0, 32, 32, p_hicons, p_ids, icon_total_count, 0)
    # From mss.windows
    srcdc = user32.GetWindowDC(0)
    memdc = gdi32.CreateCompatibleDC(srcdc)

    bmp = gdi32.CreateCompatibleBitmap(srcdc, width, height)
    gdi32.SelectObject(memdc, bmp)
    user32.DrawIcon(memdc, 0, 0, hicons[index])  # From cnblogs

    bmi = BITMAPINFO()
    bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    bmi.bmiHeader.biPlanes = 1  # Always 1
    bmi.bmiHeader.biBitCount = 32  # See grab.__doc__ [2]
    bmi.bmiHeader.biCompression = 0  # 0 = BI_RGB (no compression)
    bmi.bmiHeader.biClrUsed = 0  # See grab.__doc__ [3]
    bmi.bmiHeader.biClrImportant = 0  # See grab.__doc__ [3]

    bmi.bmiHeader.biWidth = width
    bmi.bmiHeader.biHeight = -height  # Why minus? [1]

    data = ctypes.create_string_buffer(width * height * 4)

    bits = gdi32.GetDIBits(
        memdc, bmp, 0, height, data, ctypes.byref(bmi), 0
    )
    gdi32.DeleteObject(bmp)

    if bits != height:
        raise Exception("gdi32.GetDIBits() failed.")

    return bytearray(data)


def get_rgb_data(path):
    return rgb(get_raw_data(path))


if __name__ == "__main__":
    exe = r'D:\Tencent\WeChat\WeChat.exe'
    data = rgb(get_raw_data(exe, 0, 32), 32, 32)
    import mss.tools

    mss.tools.to_png(data, (32, 32), 8, "explorer.png")
