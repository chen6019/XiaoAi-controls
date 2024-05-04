"""打包指令：
pyinstaller -F -n XiaoAi-controls --noconsole --hidden-import=wmi --hidden-import=win11toast --hidden-import=pystray --hidden-import=pillow --icon=icon.ico --add-data 'icon.ico;.' main.py
"""
# 导入各种必要的模块
import socket
import os
import wmi
from win11toast import notify
import json
import logging
from tkinter import messagebox
import sys
import pystray
from PIL import Image
import threading
import subprocess
import time
import ctypes

"""
执行系统命令，并在超时后终止命令。

参数:
- cmd: 要执行的命令（字符串或列表）
- timeout: 命令执行的超时时间（秒，默认为30秒）

返回值:
- 命令终止的返回码
"""

def execute_command(cmd, timeout=30):
    process = subprocess.Popen(cmd, shell=True)
    process.poll()
    if timeout:
        remaining = timeout
        while process.poll() is None and remaining > 0:
            time.sleep(1)
            remaining -= 1
        if remaining == 0 and process.poll() is None:
            process.kill()
    return process.wait()


"""
设置屏幕亮度。

参数:
- value: 亮度值（0-100）
"""

def set_brightness(value):
    try:
        wmi.WMI(namespace='wmi').WmiMonitorBrightnessMethods()[0].WmiSetBrightness(value, 0)
    except Exception as e:
        logging.error(f"无法设置亮度: {e}")

"""
根据接收到的命令和主题来处理相应的操作。

参数:
- command: 接收到的命令
- topic: 命令的主题
"""

def process_command(command, topic):
    # 根据主题和命令执行不同的操作
    if topic == topic1:
        # 电脑开关机控制
        if command == 'on':
            notify("电脑已经开着啦",icon=image_path)
        elif command == 'off':
            execute_command("shutdown -s -t 30")
            notify("电脑将在30秒后关机",icon=image_path)
    elif topic == topic2:
        # 屏幕亮度控制
        if command == 'off' or command == '0':
            set_brightness(0)
        elif command == 'on':
            set_brightness(100)
        else:
            try:
                brightness = int(command[3:])
                set_brightness(brightness)
            except ValueError:
                logging.error("亮度值无效")
    elif topic == topic3:
        # 应用程序的启动和关闭
        if command == 'off':
            subprocess.call(['taskkill', '/F', '/IM', app.split('\\')[-1]])
        elif command == 'on':
            subprocess.Popen(app)


"""
打开GUI界面。

无参数
无返回值
"""

def open_gui():
    if os.path.isfile("GUI.py"):
        subprocess.Popen(["python", "GUI.py"])
    elif os.path.isfile("GUI.exe"):
        subprocess.Popen(["GUI.exe"])
    else:
        logging.error("既没有找到 GUI.py 也没有找到GUI.exe")

"""
退出程序。

无参数
无返回值
"""

def exit_program():
    try:
        tcp_client_socket.close()
        icon.stop()
        logging.info("程序已停止")
        sys.exit(0)
    except SystemExit:
        pass

"""
判断当前程序是否以管理员权限运行。

返回值:
- True: 以管理员权限运行
- False: 不是以管理员权限运行
"""

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

"""
显示管理员权限提示。

无参数
无返回值
"""

def admin():
    def show_message():
        if is_admin():
            messagebox.showerror("信息","已经拥有管理员权限")
        else:
            messagebox.showerror("错误","没有管理员权限")
    threading.Thread(target=show_message).start()
    

def connTCP():
    global tcp_client_socket
    # 创建socket
    tcp_client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    # IP 和端口
    server_ip = broker
    server_port = port
    uid = secret_id
    topics = [topic for topic in [topic1, topic2, topic3] if topic]
    # 使用逗号连接主题
    topic = ','.join(topics)
    print(f'topic={topics}')
    print(f'uid={uid}&topic={topic}')
    try:
        # 连接服务器
        tcp_client_socket.connect((server_ip, server_port))
        #发送订阅指令
            
        substr = f'cmd=1&uid={uid}&topic={topic}\r\n'
        tcp_client_socket.send(substr.encode("utf-8"))
    except:
        time.sleep(2)
        connTCP()


#心跳
def Ping():
    # 发送心跳
    try:
        keeplive = 'ping\r\n'
        tcp_client_socket.send(keeplive.encode("utf-8"))
    except:
        time.sleep(2)
        connTCP()
    #开启定时，30秒发送一次心跳
    t = threading.Timer(30,Ping)
    t.start()



# 获取打包应用的根目录
if getattr(sys, 'frozen', False):
    application_path = sys._MEIPASS
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

image_path = os.path.join(application_path, 'icon.ico')

# 初始化系统托盘图标和菜单
icon = pystray.Icon("Ai-controls")
image = Image.open(image_path)
menu = (pystray.MenuItem("打开配置", open_gui), pystray.MenuItem("管理员权限查询",admin),pystray.MenuItem("退出", exit_program))
icon.menu = menu
icon.icon = image
threading.Thread(target=icon.run).start()

# 日志和配置文件路径处理
appdata_path = os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming', 'Ai-controls')
log_path = os.path.join(appdata_path, 'log.txt')
logging.basicConfig(filename=log_path, level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

config_path = os.path.join(appdata_path, 'TCP_config.json')
# 检查配置文件是否存在
if not os.path.exists(config_path):
    messagebox.showerror("Error", "配置文件不存在\n请先打开GUI配置文件")
    logging.error('TCP_config.json 文件不存在')
    icon.stop()
    open_gui()
    sys.exit(0)
else:
    with open(config_path, 'r') as f:
        TCP_config = json.load(f)

if TCP_config['test'] == 1:
    logging.info("开启测试模式:可以不启用任何主题")
else:
    if TCP_config['topic1_checked'] == 0 and TCP_config['topic2_checked'] == 0 and TCP_config['topic3_checked'] == 0 and TCP_config['topic4_checked'] == 0 and TCP_config['topic5_checked'] == 0:
        messagebox.showerror("Error", "主题不能一个都没有吧！\n（除了测试模式）")
        icon.stop()
        open_gui()
        sys.exit(0)

broker = TCP_config.get('broker')
secret_id = TCP_config.get('secret_id')
port = TCP_config.get('port')

topic1 = TCP_config.get('topic1') if TCP_config.get('topic1_checked') == 1 else None
if TCP_config.get('topic1_checked') == 1 and not topic1:
    messagebox.showerror("Error", "主题1不能为空")
    icon.stop()
    open_gui()
    sys.exit(0)
# 如果主题1不为空，将其记录到日志中
if topic1:
    logging.info(f'主题"{topic1}"')

topic2 = TCP_config.get('topic2') if TCP_config.get('topic2_checked') == 1 else None
if TCP_config.get('topic2_checked') == 1 and not topic2:
    messagebox.showerror("Error", "主题2不能为空")
    icon.stop()
    open_gui()
    sys.exit(0)
# 如果主题2不为空，将其记录到日志中
if topic2:
    logging.info(f'主题"{topic2}"')

topic3 = TCP_config.get('topic3') if TCP_config.get('topic3_checked') == 1 else None
app = TCP_config.get('app') if TCP_config.get('topic3_checked') == 1 else None
if TCP_config.get('topic3_checked') == 1 and (not topic3 or not app):
    messagebox.showerror("Error", "主题3和值不能为空")
    icon.stop()
    open_gui()
    sys.exit(0)
# 如果主题3不为空，将其记录到日志中
if topic3:
    logging.info(f'主题"{topic3}"，值："{app}"')




connTCP()
Ping()

while True:
    # 接收服务器发送过来的数据
    recvData = tcp_client_socket.recv(1024)
    if len(recvData) != 0:
        print('recv:', recvData.decode('utf-8'))
    else:
        print("conn err")
        connTCP()