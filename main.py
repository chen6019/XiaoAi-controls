# -*- coding: utf-8 -*-
# 以下代码在2021年10月21日 python3.10环境下运行通过

# import paho.mqtt.client as mqtt

# HOST = "bemfa.com"
# PORT = 9501
# client_id = ""                       
# #连接并订阅
# def on_connect(client, userdata, flags, rc):
#     print("Connected with result code "+str(rc))
#     client.subscribe("led00202")         # 订阅消息

# #消息接收
# def on_message(client, userdata, msg):
#     print("主题:"+msg.topic+" 消息:"+str(msg.payload.decode('utf-8')))

# #订阅成功
# def on_subscribe(client, userdata, mid, granted_qos):
#     print("On Subscribed: qos = %d" % granted_qos)

# # 失去连接
# def on_disconnect(client, userdata, rc):
#     if rc != 0:
#         print("Unexpected disconnection %s" % rc)


# client = mqtt.Client(client_id)
# client.username_pw_set("userName", "passwd")
# client.on_connect = on_connect
# client.on_message = on_message
# client.on_subscribe = on_subscribe
# client.on_disconnect = on_disconnect
# client.connect(HOST, PORT, 60)
# client.loop_forever()
# 这个为临时版本，将来可能会改成tcp连接的版本，因为手机端无法得知操作结果
"""打包指令：
pyinstaller -F -n XiaoAi-controls --noconsole --hidden-import=paho-mqtt --hidden-import=wmi --hidden-import=win11toast --hidden-import=pystray --hidden-import=pillow --icon=icon.ico --add-data 'icon.ico;.' main.py
"""
# 导入各种必要的模块
import io
import paho.mqtt.client as mqtt
import os
import pkg_resources
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
MQTT订阅成功时的回调函数。

参数:
- client: MQTT客户端实例
- userdata: 用户数据
- mid: 消息ID
- reason_code_list: 订阅结果的状态码列表
- properties: 属性
"""

def on_subscribe(client, userdata, mid, reason_code_list, properties):
    for sub_result in reason_code_list:
        if isinstance(sub_result, int) and sub_result >= 128:
            logging.error(f"订阅失败:{reason_code_list}")
        else:
            logging.info(f"使用代码订阅成功：{sub_result}")

"""
MQTT取消订阅时的回调函数。

参数:
- client: MQTT客户端实例
- userdata: 用户数据
- mid: 消息ID
- reason_code_list: 取消订阅结果的状态码列表
- properties: 属性
"""

def on_unsubscribe(client, userdata, mid, reason_code_list, properties):
    if len(reason_code_list) == 0 or not reason_code_list[0].is_failure:
        logging.info("退订成功")
    else:
        logging.error(f"{broker} 回复失败: {reason_code_list[0]}")
    client.disconnect()

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
            ctypes.windll.user32.LockWorkStation()
        elif command == 'off':
            execute_command("shutdown -s -t 60")
            notify("电脑将在60秒后关机")
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
    elif topic == topic4:
        if command == 'off':
            subprocess.call(['taskkill', '/F', '/IM', app2.split('\\')[-1]])
        elif command == 'on':
            subprocess.Popen(app2)
    elif topic == topic5:
        # 服务的启动和停止
        if command == 'off':
            result = subprocess.run(["sc", "stop", app3], shell=True)
            if result.returncode == 0:
                notify(f"成功关闭 {app3}")
            else:
                notify(f"关闭 {app3} 失败","可能是没有管理员权限")
        elif command == 'on':
            result = subprocess.run(["sc", "start", app3], shell=True)
            if result.returncode == 0:
                notify(f"成功启动 {app3}")
            else:
                notify(f"启动 {app3} 失败","可能是没有管理员权限")

"""
MQTT接收到消息时的回调函数。

参数:
- client: MQTT客户端实例
- userdata: 用户数据
- message: 接收到的消息
"""

def on_message(client, userdata, message):
    userdata.append(message.payload)
    command = message.payload.decode()
    logging.info(f"收到 '{command}' 从 '{message.topic}' 主题")
    process_command(command, message.topic)

"""
MQTT连接时的回调函数。

参数:
- client: MQTT客户端实例
- userdata: 用户数据
- flags: 连接标志
- reason_code: 连接结果的状态码
- properties: 属性
"""

def on_connect(client, userdata, flags, reason_code, properties):
    if reason_code.is_failure:
        notify(f"连接MQTT失败: {reason_code}. 重新连接中...")  # 连接失败时的提示
        logging.error(f"连接失败: {reason_code}. loop_forever() 将重试连接")
    else:
        notify(f"MQTT成功连接至{broker}")  # 连接成功时的提示
        logging.info(f"连接到 {broker}")
        if topic1:
            client.subscribe(topic1)
        if topic2:
            client.subscribe(topic2)
        if topic3:
            client.subscribe(topic3)
        if topic4:
            client.subscribe(topic4)
        if topic5:
            client.subscribe(topic5)

"""
打开GUI界面。

无参数
无返回值
"""

def open_gui():
    if os.path.isfile("GUI.py"):
        subprocess.Popen(["python", "GUI.py"])
        notify("正在打开配置窗口...")
    elif os.path.isfile("GUI.exe"):
        subprocess.Popen(["GUI.exe"])
        notify("正在打开配置窗口...")
    else:
        def show_message():
            current_path = os.getcwd()
            messagebox.showerror("Error", f"找不到GUI.py或GUI.exe\n当前工作路径{current_path}")
            logging.error(f"找不到GUI.py或GUI.exe\n当前工作路径{current_path}")
        threading.Thread(target=show_message).start()

"""
退出程序。

无参数
无返回值
"""

def exit_program():
    try:
        mqttc.disconnect()
        mqttc.loop_stop()
        icon.stop()
        logging.info("程序已停止")
        sys.exit(0)
    except SystemExit:
        pass

# 文件过大时截断文件
def truncate_large_file(file_path, max_size=1024*1024*50):
    if os.path.getsize(file_path) > max_size:
        with open(file_path, 'w') as f:
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
            messagebox.showinfo("信息","已经拥有管理员权限")
        else:
            messagebox.showerror("错误","没有管理员权限")
    threading.Thread(target=show_message).start()

# 获取应用程序的路径
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

# 改变当前的工作路径
os.chdir(application_path)


# 获取资源文件的路径
resource_path = 'icon.ico'

# 从资源文件中读取图像
image_data = pkg_resources.resource_string(__name__, resource_path)

# 初始化系统托盘图标和菜单
icon = pystray.Icon("Ai-controls")
image = Image.open(io.BytesIO(image_data))
menu = (pystray.MenuItem("打开配置", open_gui), pystray.MenuItem("管理员权限查询",admin),pystray.MenuItem("退出", exit_program))
icon.menu = menu
icon.icon = image
threading.Thread(target=icon.run).start()

# 日志和配置文件路径处理
appdata_path = os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming', 'Ai-controls')
# 确保目录存在
os.makedirs(appdata_path, exist_ok=True)

log_path = os.path.join(appdata_path, 'log.txt')
logging.basicConfig(filename=log_path, level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

truncate_large_file(log_path)

config_path = os.path.join(appdata_path, 'mqtt_config.json')
# 检查配置文件是否存在
if not os.path.exists(config_path):
    messagebox.showerror("Error", "配置文件不存在\n请先打开GUI配置文件")
    logging.error('mqtt_config.json 文件不存在')
    icon.stop()
    open_gui()
    sys.exit(0)
else:
    with open(config_path, 'r') as f:
        mqtt_config = json.load(f)

if mqtt_config['test'] == 1:
    logging.info("开启测试模式:可以不启用任何主题")
else:
    if mqtt_config['topic1_checked'] == 0 and mqtt_config['topic2_checked'] == 0 and mqtt_config['topic3_checked'] == 0 and mqtt_config['topic4_checked'] == 0 and mqtt_config['topic5_checked'] == 0:
        messagebox.showerror("Error", "主题不能一个都没有吧！\n（除了测试模式）")
        icon.stop()
        open_gui()
        sys.exit(0)

broker = mqtt_config.get('broker')
secret_id = mqtt_config.get('secret_id')
port = mqtt_config.get('port')

topic1 = mqtt_config.get('topic1') if mqtt_config.get('topic1_checked') == 1 else None
if mqtt_config.get('topic1_checked') == 1 and not topic1:
    messagebox.showerror("Error", "主题1不能为空")
    icon.stop()
    open_gui()
    sys.exit(0)
# 如果主题1不为空，将其记录到日志中
if topic1:
    logging.info(f'主题"{topic1}"')

topic2 = mqtt_config.get('topic2') if mqtt_config.get('topic2_checked') == 1 else None
if mqtt_config.get('topic2_checked') == 1 and not topic2:
    messagebox.showerror("Error", "主题2不能为空")
    icon.stop()
    open_gui()
    sys.exit(0)
# 如果主题2不为空，将其记录到日志中
if topic2:
    logging.info(f'主题"{topic2}"')

topic3 = mqtt_config.get('topic3') if mqtt_config.get('topic3_checked') == 1 else None
app = mqtt_config.get('app') if mqtt_config.get('topic3_checked') == 1 else None
if mqtt_config.get('topic3_checked') == 1 and (not topic3 or not app):
    messagebox.showerror("Error", "主题3和值不能为空")
    icon.stop()
    open_gui()
    sys.exit(0)
# 如果主题3不为空，将其记录到日志中
if topic3:
    logging.info(f'主题"{topic3}"，值："{app}"')

topic4 = mqtt_config.get('topic4') if mqtt_config.get('topic4_checked') == 1 else None
app2 = mqtt_config.get('app2') if mqtt_config.get('topic4_checked') == 1 else None
if mqtt_config.get('topic4_checked') == 1 and (not topic4 or not app2):
    messagebox.showerror("Error", "主题4和值不能为空")
    icon.stop()
    open_gui()
    sys.exit(0)
# 如果主题4不为空，将其记录到日志中
if topic4:
    logging.info(f'主题"{topic4}"，值："{app2}"')

topic5 = mqtt_config.get('topic5') if mqtt_config.get('topic5_checked') == 1 else None
app3 = mqtt_config.get('app3') if mqtt_config.get('topic5_checked') == 1 else None
if mqtt_config.get('topic5_checked') == 1 and (not topic5 or not app3):
    messagebox.showerror("Error", "主题5和值不能为空")
    icon.stop()
    open_gui()
    sys.exit(0)
# 如果主题5不为空，将其记录到日志中
if topic5:
    logging.info(f'主题"{topic5}"，值："{app3}"')


# 初始化MQTT客户端
mqttc = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2)
mqttc.on_connect = on_connect
mqttc.on_message = on_message
mqttc.on_subscribe = on_subscribe
mqttc.on_unsubscribe = on_unsubscribe

mqttc.user_data_set([])
mqttc._client_id = secret_id
mqttc.connect(broker, port)
try:
    mqttc.loop_forever()
except KeyboardInterrupt:
    notify("收到中断信号\n程序停止")
    logging.info("收到中断,程序停止")
    exit_program()
logging.info(f"总共收到以下消息: {mqttc.user_data_get()}")