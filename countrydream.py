# -*- coding: utf-8 -*-

import time
import numpy as np
import cv2 as cv
from matplotlib import pyplot as plt
import win32ui
import win32gui, win32api, win32con


def locate_simulator():  # 获得模拟器内部窗口的坐标
    hwnd = win32gui.FindWindow('Qt5QWindowIcon', '夜神模拟器')
    hwnd = win32gui.FindWindowEx(hwnd, 0, 'Qt5QWindowIcon', 'ScreenBoardClassWindow')
    #hwnd = win32gui.FindWindowEx(hwnd, 0, 'AnglePlayer_0', 'AnglePlayer_0')
    #hwnd = win32gui.FindWindowEx(hwnd, 0, 'Child_1', 'Child_1')
    #hwnd = win32gui.FindWindowEx(hwnd, 0, 'subWin', 'sub')
    #print(hwnd)
    #text = win32gui.GetWindowText(hwnd)
    #print(text)
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)  # 返回窗口左上角和右下角坐标
    #print(left, top, right, bottom)
    #wid = right - left
    #height = bottom - top
    #print(wid, height)
    return (left, top, right, bottom, hwnd)


def cur_prtsc(left, top, right, bottom, hwnd, text):  # 截图当前画面
    # 获取句柄窗口的大小信息
    width = int((right - left)* 1.25)
    height = int((bottom - top)* 1.25)
    #print(width, height)

    # 返回句柄窗口的设备环境，覆盖整个窗口，包括非客户区，标题栏，菜单，边框
    hWndDC = win32gui.GetDC(hwnd)
    # 创建设备描述表
    mfcDC = win32ui.CreateDCFromHandle(hWndDC)
    # 创建内存设备描述表
    saveDC = mfcDC.CreateCompatibleDC()
    # 创建位图对象准备保存图片
    saveBitMap = win32ui.CreateBitmap()
    # 为bitmap开辟存储空间
    saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
    # 将截图保存到saveBitMap中
    saveDC.SelectObject(saveBitMap)
    # 保存bitmap到内存设备描述表
    saveDC.BitBlt((0, 0), (width, height), mfcDC, (0, 0), win32con.SRCCOPY)
    # 保存为png文件
    saveBitMap.SaveBitmapFile(saveDC, text+".png")


def find_cargo():  # 定位货物的位置， 返回标记对模拟器窗口的相对坐标
    img = cv.imread('test1.png', 0)  # 读取截图
    template_img = cv.imread('x.png', 0)  # 读取货物标记
    width, height = template_img.shape[::-1]
    #print(width, height)

    # 在截图中匹配货物标记的位置，返回最佳匹配的坐标
    res = cv.matchTemplate(img, template_img, cv.TM_CCORR_NORMED)
    min_val, max_val, min_loc, max_loc = cv.minMaxLoc(res)
    #print(max_loc)
    top_left = max_loc
    #bottom_right = (top_left[0] + width, top_left[1] + height)
    '''
    cv.rectangle(img, top_left, bottom_right, 255, 2)
    cv.namedWindow('image', cv.WINDOW_NORMAL)
    cv.imshow('image', img)
    cv.waitKey(0)
    #cv.destroyAllwindows()
    '''
    # 如果匹配到的位置在小火车区域，返回货物标记的中间位置
    mid_loc = (int((top_left[0] + width/2)/1.25), int((top_left[1] + height/2)/1.25))
    #print(mid_loc)
    if mid_loc[0] < 290 and mid_loc[1] < 690:
        return None
    else:
        return mid_loc


def choose_cargo(left, top, loc, hwnd):  # 选中货物
    # 货物的坐标需要转换成在整个屏幕中的坐标，并绑定窗口句柄
    start_pos = win32gui.ScreenToClient(hwnd, (loc[0] + left, loc[1] + top))
    # test_pos = win32gui.ScreenToClient(hwnd, (left+210, top+424))
    tmp = win32api.MAKELONG(start_pos[0], start_pos[1])  # 将坐标转换成SenfMessage()需要的一维数据坐标
    win32gui.SendMessage(hwnd, win32con.WM_ACTIVATE, win32con.WM_ACTIVATE, 0)
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0000, tmp)  # 先将鼠标弹起，减少误触概率
    #print('up', loc)
    time.sleep(0.2)
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, tmp)  # 在tmp坐标处按下鼠标左键
    #print('down', loc)


def unload_cargo(left, top, loc, hwnd):  # 将货物移动到目标建筑
    target_loc = find_target()  # 获取目标建筑坐标
    #print('target:', target_loc)
    # 如果没找到坐标，返回None
    if target_loc is None:
        return None
    # 将鼠标移动到建筑上
    target_pos = win32gui.ScreenToClient(hwnd, (left+target_loc[0], top+target_loc[1]))
    tmp = win32api.MAKELONG(target_pos[0], target_pos[1])
    time.sleep(0.5)
    win32gui.SendMessage(hwnd, win32con.WM_MOUSEMOVE, win32con.MK_LBUTTON, tmp)
    #print('move', target_loc)
    time.sleep(0.5)
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0000, tmp)
    #print('up', target_loc)
    time.sleep(1)


def find_target():  # 寻找目标建筑坐标
    # 读取按下货物前后的截图
    img1 = cv.imread('test1.png')
    img2 = cv.imread('test2.png')
    #cv.imshow('img', img2)
    #cv.waitKey(0)
    #height, width = img1.shape[0:2]
    #print(height, width)
    #result_window = np.zeros((height, width), dtype=img1.dtype)

    # 将两张图片像素相减，再转换成灰度图，然后找出所有轮廓
    result_window = cv.absdiff(img1, img2)
    result_window = cv.cvtColor(result_window, cv.COLOR_BGR2GRAY)
    ret, thresh1 = cv.threshold(result_window, 40, 255, cv.THRESH_BINARY)
    contours, hierarchy = cv.findContours(thresh1, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE)
    for contour in contours:
        min = np.nanmin(contour, 0)
        max = np.nanmax(contour, 0)
        loc1 = (min[0][0], min[0][1])
        loc2 = (max[0][0], max[0][1])
        #cv.rectangle(thresh1, loc1, loc2, 255, 2)

        # 目标建筑的轮廓会特别大，返回其坐标
        if max[0][0]-min[0][0] >= 100 or max[0][1]-min[0][1] >= 100:
            #print('success:', loc1)
            #print(max[0][0] - min[0][0], max[0][1] - min[0][1])
            cv.rectangle(thresh1, loc1, loc2, 255, 2)

            #cv.imshow('img', thresh1)
            #cv.waitKey(0)

            target_loc = (int(loc1[0]/1.25+60), int(loc1[1]/1.25+40))
            return target_loc
    #cv.imshow('img', thresh1)
    #cv.waitKey(0)
    return None


sim_left, sim_top, sim_right, sim_bottom, sim_hwnd = locate_simulator()
while True:
    cur_prtsc(sim_left, sim_top, sim_right, sim_bottom, sim_hwnd, 'test1')  # 对点击货物前的画面截图
    cargo_loc = find_cargo()
    #print('cargo:', cargo_loc)
    # 没找到货物时休眠10
    if cargo_loc is None:
        time.sleep(10)
        continue
    choose_cargo(sim_left, sim_top, cargo_loc, sim_hwnd)
    time.sleep(0.5)
    cur_prtsc(sim_left, sim_top, sim_right, sim_bottom, sim_hwnd, 'test2')  # 对点击货物后的画面截图
    find_target()
    unload_cargo(sim_left, sim_top, cargo_loc, sim_hwnd)
    time.sleep(1)




