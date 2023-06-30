import sys
import time
import pyautogui
import numpy as np
from math import sqrt

WEGO_X_LIST = [*range(382, 452),779,1018,1258,1378,1498,  *range(513, 568), 659, 899, 1138] #判斷x值範圍
WEGO_CHECKING_COLOR_A = '#262d39' #判斷威格上方顏色
WEGO_CHECKING_COLOR_B = '#99eeed' #判斷威格下方顏色
WEGO_CHECKING_Y_A = 544 #判斷威格上方位置Y值
WEGO_CHECKING_Y_B = 688 #判斷威格下方位置Y值
WEGO_XA_XB_distance = 786 - 779 #上方判斷點與下方判斷點X距離差距 (下減上)
WEGO_XA_COLOR_distance = 30 #色差範圍值
WEGO_XB_COLOR_distance = 5 #色差範圍值

#將RGB值轉換成16進位格式的色碼
def rgb_to_hex(r, g, b):
    return '#{:02x}{:02x}{:02x}'.format(r, g, b)

#將16進位色碼轉換成RGB值
def hex_to_rgb(value):
    value = value.lstrip('#')
    lv = len(value)
    return tuple(int(value[i:i + lv // 3], 16) for i in range(0, lv, lv // 3))

#用於截取屏幕某個位置的像素，然後將像素的RGB值轉換成16進位色碼
def get_position_color(xy):
    try:
        screen = pyautogui.screenshot()
        r, g, b = screen.getpixel(xy)
        return rgb_to_hex(r, g, b)
    except Exception as e:
        return None

#取得指定位置xy的顏色
def get_position_color_s(xy, screen):
    try:
        r, g, b = screen.getpixel(xy)
        return rgb_to_hex(r, g, b)
    except Exception as e:
        return None

#重複抓取定位為至顏色是否一致
def wait_until_position_color(xy, hex):
    i = 0
    while i < 1000:
        if get_position_color(xy) == hex:
            return
        time.sleep(0.05)
        i += 1

#舉例疊代382位置，還有抓取的螢幕畫面
def wego_in(x, screen):
    u = color_distance(get_position_color_s((x, WEGO_CHECKING_Y_A), screen), 
                       WEGO_CHECKING_COLOR_A)
    d = color_distance(get_position_color_s((x + WEGO_XA_XB_distance, WEGO_CHECKING_Y_B), 
                                            screen), WEGO_CHECKING_COLOR_B)
    if d <= WEGO_XB_COLOR_distance  and u <= WEGO_XA_COLOR_distance:
        print('we', x, u, d)
        return True
    return False

def move_circle(x, y, diameter, duration):
    steps = 50      # 圆形轨迹的分辨率，此处为 50 等分
    radius = diameter / 2
    interval = duration / steps
    for i in range(steps):
        angle = i / steps * math.pi * 2
        dx = math.cos(angle) * radius
        dy = math.sin(angle) * radius
        pyautogui.moveTo(x + dx, y + dy, interval=interval)

#將參數WEGO_X_LIST的值一直疊代下去
def is_wego_in():
    screen = pyautogui.screenshot()
    li = WEGO_X_LIST
    for i in li:
        if wego_in(i, screen):
            return True
    return False

#學習範本
def sample_wego_in(x, screen):
    u = color_distance(get_position_color_s((x, WEGO_CHECKING_Y_A), screen), 
                       WEGO_CHECKING_COLOR_A)
    d = color_distance(get_position_color_s((x + WEGO_XA_XB_distance, WEGO_CHECKING_Y_B), 
                                            screen), WEGO_CHECKING_COLOR_B)
    if d <= WEGO_XB_COLOR_distance  and u <= WEGO_XA_COLOR_distance:
        print('we', x, u, d)
        return True
    return False

#學習範本
def sample_is_wego_in():
    screen = pyautogui.screenshot()
    li = WEGO_X_LIST
    for i in li:
        if sample_wego_in(i, screen):
            return True
    return False

def color_distance(ha, hb):
    r, g, b = hex_to_rgb(ha)
    cr, cg, cb = hex_to_rgb(hb)
    color_diff = sqrt((r - cr) ** 2 + (g - cg) ** 2 + (b - cb) ** 2)
    return color_diff

#抓顏色
def get_color():
    mouse_now = None
    while True:
        if mouse_now != pyautogui.position():
            mouse_now = pyautogui.position()
            color_hex = get_position_color(mouse_now)

            print(mouse_now, color_hex)
        time.sleep(0.05)

def test():
    # 判斷sr劍及UR炸彈是否出現
    sample_is_wego_in():

    # 将鼠标移动到 (100, 100)
    pyautogui.moveTo(100, 100)
    # 滑动一个直径为 50 的圆圈，持续时间为 5 秒
    move_circle(100, 100, 50, 5)


if __name__ == '__main__':
    test()

'''
if __name__ == '__main__':
    # 抓顏色和位置用
    # get_color()
    # sys.exit()

    i = 0
    mouse_now = None
    while i < 500:
        # 判斷再抽一次按鈕是否出現
        wait_until_position_color((1274, 889), '#59c9ff')

        # 判斷sr劍及UR炸彈是否出現
        if is_wego_in():
            print('found')
            break

        pyautogui.click(1274, 889)
        time.sleep(0.1)

        # 判斷確認按鈕是否出現
        wait_until_position_color((1055, 641), '#404040')
        time.sleep(0.5)
        pyautogui.click(1055, 641)
        time.sleep(0.1)

        # 判斷skip按鈕是否出現
        wait_until_position_color((1598, 152), '#ffffff')
        pyautogui.click(1592, 154)
        time.sleep(0.1)

        i += 1
'''