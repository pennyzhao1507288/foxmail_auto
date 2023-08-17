import pyautogui
import subprocess
import time
import pandas as pd  # 用于处理 Excel 表格
from PIL import Image
import cv2
import numpy
import sys


# 获取命令行参数，参数0是脚本文件名，参数1是 Excel 路径，参数2是应用路径
excel_file_path = sys.argv[1]
app_path = sys.argv[2]
shortcut_key = sys.argv[3]
search_box_type = sys.argv[4]

#启动应用
subprocess.Popen([app_path])
# 等待一段时间，确保 Foxmail 应用完全启动
time.sleep(1)
# 获取 Foxmail 应用窗口的位置和大小
window = pyautogui.getWindowsWithTitle("Foxmail")[0]
# 将窗口激活（确保它在前台）
window.activate()
# 最大化窗口
# pyautogui.hotkey('win', 'up')

df = pd.read_excel(excel_file_path)

# 假设重复内容高光颜色为黄色，RGB 值为 (255, 255, 0)
highlight_color = (255, 255, 0)

# 假设无内容的标识图片文件为 "no_content.png"
no_content_image = "image/no_content.png"


# 加载搜索框图片
subject_search_box_image = cv2.imread("image/subject_search_box.png")    #主题搜索框
full_text_search_box_image = cv2.imread("image/full_text_search_box.png")   #全文搜索框
# default_search_box_image = cv2.imread("default_search_box_image.png")
# 定义搜索框位置的偏移量，向右移动 50 个像素
search_box_offset_x = 100# 定义搜索框位置的偏移量，向右移动 50 个像素


# 加载滚轮区域图片
scroll_area_image = cv2.imread("image/scroll_area.png")
scroll_area_offset_x = 100

# 模拟按下快捷键
def press_shortcut_key():
    pyautogui.hotkey(*shortcut_key.split("+"))  # 使用 * 操作符展开列表作为参数

# 使用 OpenCV 进行图片识别，定位搜索框的位置
def locate_search_box(search_box_image):
    screenshot = pyautogui.screenshot()
    screenshot = cv2.cvtColor(numpy.array(screenshot), cv2.COLOR_RGB2BGR)  # 将 Pillow 图像转换为 OpenCV 格式
    result = cv2.matchTemplate(screenshot, search_box_image, cv2.TM_CCOEFF_NORMED)
    _, _, _, max_loc = cv2.minMaxLoc(result)
    # 计算搜索框的位置，加上偏移量
    search_box_x = max_loc[0] + search_box_offset_x
    search_box_y = max_loc[1] + search_box_image.shape[0] // 2
    return search_box_x, search_box_y


# 使用 OpenCV 进行图片识别，定位滚轮区域的位置
def locate_scroll_area():
    screenshot = pyautogui.screenshot()
    screenshot = cv2.cvtColor(numpy.array(screenshot), cv2.COLOR_RGB2BGR)  # 将 Pillow 图像转换为 OpenCV 格式

    result = cv2.matchTemplate(screenshot, scroll_area_image, cv2.TM_CCOEFF_NORMED)
    _, _, _, max_loc = cv2.minMaxLoc(result)
    # 返回滚轮区域的中心位置
    scroll_area_x = max_loc[0] + scroll_area_offset_x
    scroll_area_y = max_loc[1] + scroll_area_image.shape[0] // 2
    return scroll_area_x, scroll_area_y

# for keyword in df.iloc[:, 0]:  # 假设关键字在第一列
# 循环遍历 Excel 表格中的关键字，并在搜索框中搜索
for index, row in df.iterrows():
    # 提取前两个单词作为关键字
    words = row[0].split()
    if len(words) >= 2:
        keyword = f"{words[0]} {words[1]}"
    elif len(words) == 1:
        keyword = words[0]
    else:
        continue  # 跳过空行

    # 根据搜索框类型选择不同的图片
    if search_box_type == "主题搜索框":
        search_box_image = subject_search_box_image
    elif search_box_type == "全文搜索框":
        search_box_image = full_text_search_box_image
    else:
        # 默认图片（如果未选择搜索框类型）
        search_box_image = full_text_search_box_image

    # 调用函数定位搜索框的位置
    search_box_x, search_box_y = locate_search_box(search_box_image)
    pyautogui.click(search_box_x, search_box_y)
    # 清空搜索框内容
    pyautogui.hotkey("ctrl", "a")  # 选择全部文本
    pyautogui.press("delete")      # 删除选中的文本
    # 输入关键字
    pyautogui.write(keyword)
    # 模拟回车键，执行搜索
    pyautogui.press("enter")
    # 等待一段时间，给应用时间进行搜索
    pyautogui.sleep(2)  # 适当调整等待时间

    # 初始化一个标志，用于判断是否找到高光
    found_highlight = False

    # 尝试找到高光，或滚动页面并重试
    for _ in range(3):  # 最多尝试 3 次
        # 截取屏幕并打开为 Pillow 图像
        screenshot = pyautogui.screenshot()
        screenshot = screenshot.convert("RGB")
        # 查找高光颜色的像素位置
        width, height = screenshot.size
        highlight_positions = []  # 存储高光位置的列表
        for x in range(width):
            for y in range(height):
                pixel_color = screenshot.getpixel((x, y))
                if pixel_color == highlight_color:
                    highlight_positions.append((x, y))
        if highlight_positions:
            # 找到 x 坐标和 y 坐标最小的高光位置
            min_x, min_y = min(highlight_positions, key=lambda pos: pos[0])
            # 将鼠标移动到最小的高光位置并点击
            pyautogui.click(min_x, min_y)
            # 模拟 Ctrl+Shift+R
            pyautogui.hotkey("ctrl", "shift", "r")
            # 等待页面刷新
            time.sleep(1)  # 适当调整等待时间
            # 移动鼠标到新页面的位置并点击两次
            # pyautogui.click(new_page_x, new_page_y, clicks=2)

            #快速文本快捷键
            press_shortcut_key()
            # pyautogui.hotkey("alt", "0")

            # 模拟按下 Ctrl+enter发送邮件
            pyautogui.hotkey("ctrl", "enter")

            # 等待一段时间，给页面时间处理操作
            time.sleep(1)  # 适当调整等待时间

            # 返回到原搜索页面，继续输入下一个关键字
            pyautogui.click(search_box_x, search_box_y)

            # 将标志设置为 True，表示找到了高光
            found_highlight = True
            break  # 跳出循环，不再尝试滚动页面
        else:
            # 使用图片识别判断是否出现了无内容的标识
            no_content_position = pyautogui.locateOnScreen(no_content_image, confidence=0.8)
            if no_content_position:
                print("搜索结果区域出现无内容")
                break  # 无内容时，跳出循环，不再尝试滚动页面
            else:

                # 调用函数定位滚轮区域的位置
                scroll_area_x, scroll_area_y = locate_scroll_area()

                # 鼠标移动到滚轮区域并点击，模拟滚轮操作
                pyautogui.click(scroll_area_x, scroll_area_y)
                pyautogui.scroll(-3)  # 向下滚动3次

                time.sleep(1)  # 适当调整等待时间

    if not found_highlight and len(words) >= 1:
        # 如果前两次都没有高光，先清空搜索框，然后尝试使用前一个单词作为关键字
        pyautogui.click(search_box_x, search_box_y)
        pyautogui.hotkey("ctrl", "a")  # 选择全部文本
        pyautogui.press("delete")  # 删除选中的文本
        keyword = words[0]
        # 输入关键字
        pyautogui.write(keyword)

        # 模拟回车键，执行搜索
        pyautogui.press("enter")

        # 等待一段时间，给应用时间进行搜索
        time.sleep(1)  # 适当调整等待时间

        # 尝试找到高光，或滚动页面并重试
        for _ in range(3):  # 最多尝试 3 次
            # 截取屏幕并打开为 Pillow 图像
            screenshot = pyautogui.screenshot()
            screenshot = screenshot.convert("RGB")

            # 查找高光颜色的像素位置
            width, height = screenshot.size
            highlight_positions = []  # 存储高光位置的列表
            for x in range(width):
                for y in range(height):
                    pixel_color = screenshot.getpixel((x, y))
                    if pixel_color == highlight_color:
                        highlight_positions.append((x, y))

            if highlight_positions:
                # 找到 x 坐标和 y 坐标最小的高光位置
                min_x, min_y = min(highlight_positions, key=lambda pos: pos[0])

                # 将鼠标移动到最小的高光位置并点击
                pyautogui.click(min_x, min_y)

                # 模拟 Ctrl+Shift+R
                pyautogui.hotkey("ctrl", "shift", "r")

                # 等待页面刷新
                time.sleep(1)  # 适当调整等待时间

                #快速文本快捷键
                press_shortcut_key()
                # pyautogui.hotkey("alt", "0")

                # 模拟按下 Ctrl+enter
                pyautogui.hotkey("ctrl", "enter")

                # 等待一段时间，给页面时间处理操作
                time.sleep(1)  # 适当调整等待时间

                # 返回到原搜索页面，继续输入下一个关键字
                pyautogui.click(search_box_x, search_box_y)

                break  # 跳出循环，不再尝试滚动页面
            else:
            # 鼠标移动到可以滚的位置
                # 调用函数定位滚轮区域的位置
                scroll_area_x, scroll_area_y = locate_scroll_area()
                # 鼠标移动到滚轮区域并点击，模拟滚轮操作
                pyautogui.click(scroll_area_x, scroll_area_y)
                pyautogui.scroll(-3)  # 向下滚动3次
                time.sleep(1)  # 适当调整等待时间
        else:
            print("未找到高光")
        # pyautogui.alert('This is a message.', 'Important')
        a = pyautogui.confirm(text='send successfully,please press ok to continue',title='消息提醒',buttons=['ok','cancel'])
        if a == "ok":
            continue
        else:
            pyautogui.alert(text='不ok也得ok', title='important', button='OK')
