import paho.mqtt.client as mqtt
import os
import wmi
from windows_toasts import Toast, WindowsToaster
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

# 新增用于执行系统命令的函数，并增加异常处理
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

def toast(info):
    newToast.text_fields = [info]
    toaster.show_toast(newToast)

def on_subscribe(client, userdata, mid, reason_code_list, properties):
    for sub_result in reason_code_list:
        if isinstance(sub_result, int) and sub_result >= 128:
            print("Subscribe failed")
        else:
            print(f"Subscription successful with code: {sub_result}")

def on_unsubscribe(client, userdata, mid, reason_code_list, properties):
    if len(reason_code_list) == 0 or not reason_code_list[0].is_failure:
        print("Unsubscribe succeeded")
    else:
        print(f"Broker replied with failure: {reason_code_list[0]}")
    client.disconnect()

def set_brightness(value):
    try:
        wmi.WMI(namespace='wmi').WmiMonitorBrightnessMethods()[0].WmiSetBrightness(value, 0)
    except Exception as e:
        print(f"Failed to set brightness: {e}")

def process_command(command, topic):
    if topic == topic1:
        if command == 'on':
            toast("电脑已经开着啦")
        elif command == 'off':
            execute_command("shutdown -s -t 30")
            toast("电脑将在30秒后关机")
    elif topic == topic2:
        if command == 'off' or command == '0':
            set_brightness(0)
        elif command == 'on':
            set_brightness(100)
        else:
            try:
                brightness = int(command[3:])
                set_brightness(brightness)
            except ValueError:
                print("Invalid brightness value")
    elif topic == topic3:
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
        if command == 'off':
            subprocess.run(["sc", "stop", app3], shell=True)
        elif command == 'on':
            subprocess.run(["sc", "start", app3], shell=True)

def on_message(client, userdata, message):
    userdata.append(message.payload)
    command = message.payload.decode()
    print(f"Received `{command}` from `{message.topic}` topic")
    process_command(command, message.topic)

def on_connect(client, userdata, flags, reason_code, properties):
    if reason_code.is_failure:
        toast(f"连接MQTT失败: {reason_code}. 重新连接中...")
        print(f"Failed to connect: {reason_code}. loop_forever() will retry connection")
    else:
        toast(f"MQTT成功连接至{broker}")
        print("Connected to", broker)
        client.subscribe(topic1)
        client.subscribe(topic2)
        client.subscribe(topic3)
        client.subscribe(topic4)
        client.subscribe(topic5)

def open_gui():
    subprocess.Popen(["python", "GUI.py"])

def exit_program():
    mqttc.disconnect()
    mqttc.loop_stop()
    icon.stop()
    sys.exit(0)

# 判断是否拥有管理员权限
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def admin():
    def show_message():
        if is_admin():
            messagebox.showerror("info","已经拥有管理员权限")
        else:
            messagebox.showerror("error","没有管理员权限")
    threading.Thread(target=show_message).start()


icon = pystray.Icon("Ai-controls")
image = Image.open("icon.png")
menu = (pystray.MenuItem("打开配置", open_gui), pystray.MenuItem("管理员权限查询",admin),pystray.MenuItem("退出", exit_program))
icon.menu = menu
icon.icon = image
threading.Thread(target=icon.run).start()

appdata_path = os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming', 'Ai-controls')
log_path = os.path.join(appdata_path, 'log.txt')
logging.basicConfig(filename=log_path, level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

config_path = os.path.join(appdata_path, 'mqtt_config.json')
if not os.path.exists(config_path):
    messagebox.showerror("Error", "配置文件不存在\n请先打开GUI配置文件")
    logging.error('mqtt_config.json 文件不存在')
    icon.stop()
    sys.exit(0)
else:
    with open(config_path, 'r') as f:
        mqtt_config = json.load(f)

# 从MQTT配置中获取值并赋值给变量
broker = mqtt_config.get('broker', '')
topic1 = mqtt_config.get('topic1', '')
topic2 = mqtt_config.get('topic2', '')
topic3 = mqtt_config.get('topic3', '')
app = mqtt_config.get('app', '')
topic4 = mqtt_config.get('topic4', '')
app2 = mqtt_config.get('app2', '')
topic5 = mqtt_config.get('topic5', '')
app3 = mqtt_config.get('app3', '')
secret_id = mqtt_config.get('secret_id', '')
port = mqtt_config.get('port', '')

toaster = WindowsToaster('Python')
newToast = Toast()
info = '加载中..'
newToast.text_fields = [info]
newToast.on_activated = lambda _: print('Toast clicked!')
toaster.show_toast(newToast)

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
    print("Interrupt received, stopping...")
    toast("收到中断信号\n程序停止")
    logging.info("收到中断信号,程序停止")
    exit_program()
print(f"Received the following message: {mqttc.user_data_get()}")