"""
pip install PySimpleGUI pyautogui pyscreenshot pillow baidu-aip chardet auto-py-to-exe pypiwin32 pyscreenshot
"""
import time
# 界面库
import PySimpleGUI as sg
# 发送图片到剪贴板
import win32clipboard
# 获取桌面尺寸
import pyautogui as pag
# -------------------------------------------------------
import random
# 发送图片到剪贴板
from io import BytesIO
import win32clipboard
from PIL import Image
# 识图api
from aip import AipOcr
# 剪贴板
import pyperclip
# 翻译用到的库
import http.client
import hashlib
import urllib
import random
import json
# 获取桌面尺寸
import pyautogui as pag
import PySimpleGUI as sg


# 识别图片文字设置
# 你的 APPID AK SK
APP_ID = ''
API_KEY = ''
SECRET_KEY = ''
# 百度翻译设置
APPID = ''  # 填写你的appid
SECRETKEY = ''  # 填写你的密钥


screen_width, screen_height = pag.size()

def findApp():
    try:
        app =  pag.locateCenterOnScreen('flag.png', grayscale=True)
        x, y = app
        pag.moveTo(x+30, y+30)
        pag.click()
    except:
        pass

def screenShot(path):
    """
    保存全屏截图
    """
    pag.screenshot(path)

def popUP(text, x=screen_width/2-100, y=screen_height/2-10):
    """
    弹出提示消息
    """
    sg.popup_timed(text,
    no_titlebar=True, 
    location=(x, y),
    auto_close_duration=1,
    button_type=5,
    text_color="yellow") 


def partScreenShot(start_x, start_y, end_x, end_y):
    """
    保存局部全屏截图
    """  
    try:
        if end_y <= start_y and end_x <= start_x:
            pag.screenshot("picture.png", region=(end_x+2, end_y+2, start_x-end_x-2, start_y-end_y-2))
        elif end_y <= start_y:
            pag.screenshot("picture.png", region=(start_x+2, end_y+2, end_x-start_x-2, start_y-end_y-2))
        elif end_x <= start_x:
            pag.screenshot("picture.png", region=(end_x+2, start_y+2, start_x-end_x-2, end_y-start_y-2))
        else:
            pag.screenshot("picture.png", region=(start_x+2, start_y+2, end_x-start_x-2, end_y-start_y-2))
        popUP("已保存区域截图")
    except:
        pass

def send_to_clipboard(clip_type, filepath):
    """
    向剪贴板发送内容
    """
    image = Image.open(filepath)
    output = BytesIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(clip_type, data)
    win32clipboard.CloseClipboard()
    popUP("已把图片发送剪切板")

def get_file_content(file_path):
    """
    获取文件内容
    """
    with open(file_path, "rb") as fp:
        return fp.read()


def recognize(APP_ID, API_KEY, SECRET_KEY):
    client = AipOcr(APP_ID, API_KEY, SECRET_KEY)
    image = get_file_content('picture.png')
    res_image = client.basicGeneral(image)
    result = '\n'.join([words['words']
                        for words in res_image['words_result']])
    pyperclip.copy(result)
    popUP('识别成功')
    return result


def translate(appid, secret_key, lang):
    """
    翻译
    """
    appid = appid
    secret_key = secret_key

    http_client = None
    myurl = '/api/trans/vip/translate'

    question = pyperclip.paste()
    from_lang = 'auto'  # 原文语种
    to_lang = lang  # 译文语种
    salt = random.randint(32768, 65536)
    sign = appid + question + str(salt) + secret_key
    sign = hashlib.md5(sign.encode()).hexdigest()
    myurl = myurl + '?appid=' + appid + '&q=' + urllib.parse.quote(  # type: ignore
        question) + '&from=' + from_lang + '&to=' + to_lang + '&salt=' + str(salt) + '&sign=' + sign

    try:
        http_client = http.client.HTTPConnection('api.fanyi.baidu.com')
        http_client.request('GET', myurl)

        # response是HTTPResponse对象
        response = http_client.getresponse()
        result_all = response.read().decode("utf-8")
        result = json.loads(result_all)
        result = '\n'.join([src['dst'] for src in result['trans_result']])
        pyperclip.copy(result)
        popUP(f'翻译成功:\n{result}')
        return  result
    except Exception as error:
        print(error)
    finally:
        if http_client:
            http_client.close()

class Pos:
    def __init__(self) -> None:
        self.start_x = self.start_y = 0
        self.end_x = screen_width
        self.end_y = screen_height
        self.start_point = None
        self.end_point = None
        self.prior_rect = None
sg.theme("DarkBlue3")
sg.set_options(font=("Microsoft Yahei UI", 10, 'bold'))

lay1 = [
        [
            sg.Button('在此页面截图', key="capture"),
            sg.Button('说明', key="help")
        ],
        [[sg.Text('识别的结果将会显示在这里', key="-OUTPUT-")]]
    ]
win1 = sg.Window(                                                           # win1 主窗口
            '截图识别翻译工具', 
            lay1,
            no_titlebar=True,
            location=(1500, 500),
            icon='grab.ico',
            finalize=True,
            keep_on_top=True,
            grab_anywhere=True
                )
win1.bind("<Alt_L><w>", "ALT-w")
win1.bind("<Alt_L><x>", "ALT-x")
win1.bind("<Alt_L><s>", "ALT-s")
win1.bind("<Alt_L><g>", "ALT-g")
win1.bind("<Alt_L><r>", "ALT-r")
win1.bind("<Alt_L><f>", "ALT-f")
win1.bind("<Alt_L><e>", "ALT-e")
win2_active = False
while True:                                                                 # win1 主窗口逻辑
    ev1, vals1 = win1.read()
    if ev1 == sg.WIN_CLOSED or ev1 == "ALT-g":
        break
    if ev1 == "ALT-r":
        outup = recognize(APP_ID, API_KEY, SECRET_KEY)
        win1['-OUTPUT-'].update(outup)
    if ev1 in ("ALT-e"):
        outup = translate(APPID, SECRETKEY, "en")
        win1['-OUTPUT-'].update(outup)
    if ev1 in ("ALT-f"):
        outup = translate(APPID, SECRETKEY, "zh")
        win1['-OUTPUT-'].update(outup)
    if ev1 == "ALT-s":
        send_to_clipboard(win32clipboard.CF_DIB, 'picture.png')
    if ev1 == "ALT-x":
        layout3 = [
                [
                    sg.Image('picture.png', pad=(0, 0), expand_x=True, expand_y=True),
                ]
            ]
        win3 = sg.Window(                                                   # win3 贴图窗口
                    '贴图',
                    layout3,
                    icon='grab.ico', 
                    finalize=True, 
                    keep_on_top=True,
                    grab_anywhere=True,
                    no_titlebar=True, 
                    margins=(0, 0)
                        )
        win3.bind("<Alt_L><z>", "ALT-z")
        while True:                                                         # win3 贴图窗口逻辑
            ev3, vals3 = win3.read()
            if ev3 == "ALT-z":
                popUP("退出贴图")
                findApp()
                break
        win3.close()
    if ev1 in ("help"):
        sg.popup(
            """快捷键
            alt+g 退出软件
            alt+w 进入区域截图界面
            alt+q 退出区域截图界面
            alt+c 区域截图
            alt+r 识别区域截图的内容（内容在剪贴板里）
            alt+f 英文翻译成中文（内容在剪贴板里）
            alt+e 中文翻译成英文（内容在剪贴板里）
            alt+s 把区域截图发送到剪切板
            alt+x 贴图（把区域截图固定在窗口上）
            alt+z 取消贴图
            """, 
        no_titlebar=True, 
        location=(screen_width/2-100, screen_height/2-500),
        text_color="yellow")

    if ev1 == 'capture' and not win2_active or ev1 in ("ALT-w") and not win2_active:
        win1.hide()
        # 保存全屏截图
        image_file = r'background.png'  # 全屏图片位置
        screenShot(image_file)
        win2_active = True  # 显示截图页面
        popUP('请选择要截图的区域')
        layout2 = [[sg.Graph(pad=(0, 0), canvas_size=(screen_width, screen_height),
                    graph_bottom_left=(0, 0),
                    graph_top_right=(screen_width, screen_height),
                    key="-GRAPH-",
                    change_submits=True,    # 鼠标点击事件
                    background_color='lightblue',
                    drag_submits=True), ]]
        win2 = sg.Window(                                                   # win2 区域截图窗口
                    "draw rect on image",
                    layout2,
                    size=(screen_width, screen_height), 
                    icon='grab.png',
                    margins=(0, 0),
                    finalize=True,
                    no_titlebar=True
                )
        win2.bind("<Alt_L><q>", "ALT-q")
        win2.bind("<Alt_L><c>", "ALT-c")
        # get the graph element for ease of use later
        graph = win2["-GRAPH-"]  # type: sg.Graph
        graph.draw_image(image_file, location=(
            0, screen_height), ) if image_file else None
        dragging = False
        pos = Pos()
        while True:                                                         # win2 区域截图窗口逻辑
            ev2, vals2 = win2.read()
            if ev2 in ("ALT-q") or ev2 == sg.WIN_CLOSED or ev2 == 'Exit':
                popUP("退出截取图片")
                win2.close()
                win2_active = False
                win1.normal()
                findApp()
                break
            if ev2 in ("ALT-c"):
                partScreenShot(pos.start_x, pos.start_y, pos.end_x, pos.end_y)
            if ev2 == "-GRAPH-":  # if there's a "Graph" event, then it's a mouse
                x, y = vals2["-GRAPH-"]
                if not dragging:
                    pos.start_point = (x, y)
                    dragging = True
                else:
                    pos.end_point = (x, y)
                if pos.prior_rect:
                    graph.delete_figure(pos.prior_rect)
                if None not in (pos.start_point, pos.end_point):
                    pos.prior_rect = graph.draw_rectangle(
                        pos.start_point, pos.end_point, line_color='red', line_width=1)
            elif ev2.endswith('+UP'):  # The drawing has ended because mouse up
                # 截图逻辑
                try:
                    pos.start_x, pos.start_y = pos.start_point  # type: ignore
                    pos.end_x, pos.end_y = pos.end_point  # type: ignore
                except:
                    pass
                pos.start_y = 1080 - pos.start_y
                pos.end_y = 1080 - pos.end_y
                pos.start_point, pos.end_point = None, None  # enable grabbing a new rect
                dragging = False
win1.close()
