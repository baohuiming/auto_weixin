"""
@Description :   自动登录微信，需要安装pyautogui、pillow、psutil、pywin32、win32gui
@Author      :   Bao Huiming 
@Mail        :   m@baohuiming.net
@Time        :   2023/02/27 14:48:48
"""

import pyautogui
from PIL import Image
import psutil
import os
import time

exe_path = "E:\Program Files (x86)\Tencent\WeChat\WeChat.exe"

def is_weixin_running():
    # 判断微信进程是否存在
    pl = psutil.pids()
    for pid in pl:
        if psutil.Process(pid).name() == "WeChat.exe":
            print('微信已运行：', pid)
            return True
    else:
        print("微信未运行")
        return False
        

def is_weixin_logined():
    # 判断微信进程是否存在
    pl = psutil.pids()
    for pid in pl:
        if psutil.Process(pid).name() == "WeChatUtility.exe":
            print('微信已登录：', pid)
            return True
    else:
        print("微信未登录")
        return False
        

def click_login_button():
    #事先对按钮截图
    buttonImg= Image.open("login-button.png")
    #截图当前屏幕并找到之前加载的按钮截图
    msg = pyautogui.locateOnScreen(buttonImg, grayscale=True,confidence=.9)
    if msg==None: 
        print ("没找到微信登录窗口")
        put_forward()
        return False
    else:
        x,y,width,height=msg
        print ("登录按钮在屏幕中的位置是：X={},Y={}，宽{}像素,高{}像素".format(x,y,width,height))
        #左键点击屏幕上的这个位置
        pyautogui.click(x + width/2,y + height/2)
        print("点击登录成功")
        return True


def put_forward():
    import win32gui
    import win32con
    def get_all_hwnd(hwnd, mouse):
        if (win32gui.IsWindow(hwnd) and
            win32gui.IsWindowEnabled(hwnd) and
            win32gui.IsWindowVisible(hwnd)):
            hwnd_map.update({hwnd: win32gui.GetWindowText(hwnd)})
    
    hwnd_map = {}
    win32gui.EnumWindows(get_all_hwnd, 0)
    
    for h, t in hwnd_map.items():
        if t :
            if t == '微信':
                # h 为想要放到最前面的窗口句柄
                print("找到微信窗口：",h)
    
                win32gui.BringWindowToTop(h)
                
                # 被其他窗口遮挡，调用后放到最前面
                win32gui.SetForegroundWindow(h)
    
                # 解决被最小化的情况
                win32gui.ShowWindow(h, win32con.SW_RESTORE)


if __name__ == '__main__':
    # 判断微信是否运行
    running = is_weixin_running()
    if not running:
        # 启动微信
        os.startfile(exe_path)
        time.sleep(1)

    # 判断微信是否登录
    logined = is_weixin_logined()
    if not logined:
        for _ in range(100):
            # 尝试100次
            res = click_login_button()
            if res:
                break
            time.sleep(0.05)
