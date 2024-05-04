import os
import json
import tkinter as tk
from tkinter import messagebox
import sys
import subprocess
def TEST():
    save_config()
    if os.path.isfile("main.py"):
        subprocess.Popen(["python", "main.py"])
    elif os.path.isfile("XiaoAi-controls.exe"):
        subprocess.Popen(["XiaoAi-controls.exe"])
    else:
        messagebox.showerror("Error","既没有找到 main.py 也没有找到Ai-controls.exe")
    
def open_config():
    os.startfile(appdata_path)
def on_closing():
    root.quit()
    sys.exit()
def save_config():
    broker = broker_entry.get()
    secret_id = secret_id_entry.get()
    port = port_entry.get()
    topic1 = topic1_entry.get()
    topic2 = topic2_entry.get()
    topic3 = topic3_entry.get()
    app = app_entry.get()
    topic4 = topic4_entry.get()
    app2 = app2_entry.get()
    topic5 = topic5_entry.get()
    app3 = app3_entry.get()

    test = test_checkbutton_var.get()
    topic1_checked = topic1_checkbutton_var.get()
    topic2_checked = topic2_checkbutton_var.get()
    topic3_checked = topic3_checkbutton_var.get()
    topic4_checked = topic4_checkbutton_var.get()
    topic5_checked = topic5_checkbutton_var.get()

    # 尝试将port的值转换为整数
    try:
        port = int(port)
    except ValueError:
        messagebox.showerror("Error", "端口必须为整数")
        return

    if not broker or not topic1 or not topic2 or not topic3 or not topic4 or not topic5 or not secret_id or not port:
        messagebox.showerror("Error", "所有字段都必须填写")
        return

    mqtt_config = {
        'broker': broker,
        'secret_id': secret_id,
        'port': port,
        'test': test,
        'topic1': topic1,
        'topic1_checked': topic1_checked,
        'topic2': topic2,
        'topic2_checked': topic2_checked,
        'topic3': topic3,
        'topic3_checked': topic3_checked,
        'app': app,
        'topic4': topic4,
        'topic4_checked': topic4_checked,
        'app2':app2,
        'topic5': topic5,
        'topic5_checked': topic5_checked,
        'app3':app3
    }
    
    with open(config_path, 'w') as f:
        json.dump(mqtt_config, f, indent=4)

    messagebox.showinfo("Ok", "配置保存成功")

# 获取APPDATA目录的路径
appdata_path = os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming', 'Ai-controls')

# 检查目录是否存在，如果不存在则创建目录 
if not os.path.exists(appdata_path):
    os.makedirs(appdata_path)

# 创建mqtt_config.json文件的路径
config_path = os.path.join(appdata_path, 'mqtt_config.json')

# 检查文件是否存在，如果不存在则创建文件
if not os.path.exists(config_path):
    mqtt_config = {
        'broker': '',
        'secret_id': '',
        'port': '',
        'test': '',
        'topic1': '',
        'topic1_checked': 0,
        'topic2': '',
        'topic2_checked': 0,
        'topic3': '',
        'topic3_checked': 0,
        'app': '',
        'topic4': '',
        'topic4_checked': 0,
        'app2': '',
        'topic5': '',
        'topic5_checked': 0,
        'app3': ''
    }
    with open(config_path, 'w') as f:
        json.dump(mqtt_config, f)

# 从json文件中读取MQTT配置
with open(config_path, 'r') as f:
    mqtt_config = json.load(f)

# 从MQTT配置中获取值并赋值给变量
broker = mqtt_config.get('broker', '')
secret_id = mqtt_config.get('secret_id', '')
port = mqtt_config.get('port', '')
test_checkbutton_var= mqtt_config.get('test', 0)
topic1 = mqtt_config.get('topic1', '')
topic1_checked = mqtt_config.get('topic1_checked', 0)
topic2 = mqtt_config.get('topic2', '')
topic2_checked = mqtt_config.get('topic2_checked', 0)
topic3 = mqtt_config.get('topic3', '')
topic3_checked = mqtt_config.get('topic3_checked', 0)
app = mqtt_config.get('app', '')
topic4 = mqtt_config.get('topic4', '')
topic4_checked = mqtt_config.get('topic4_checked', 0)
app2 = mqtt_config.get('app2', '')
topic5 = mqtt_config.get('topic5', '')
topic5_checked = mqtt_config.get('topic5_checked', 0)
app3 = mqtt_config.get('app3', '')

root = tk.Tk()

root.title("MQTT配置")

root.protocol("WM_DELETE_WINDOW", on_closing)

broker_label = tk.Label(root, text="Broker （IOT平台）")
broker_label.grid(row=0, column=0, pady=5)
broker_entry = tk.Entry(root)
broker_entry.insert(0, broker)  
broker_entry.grid(row=0, column=1, pady=5, padx=10)

secret_id_label = tk.Label(root, text="key（私钥）：")
secret_id_label.grid(row=1, column=0, pady=5)
secret_id_entry = tk.Entry(root)
secret_id_entry.insert(0, secret_id)  
secret_id_entry.grid(row=1, column=1, pady=5, padx=10)

port_label = tk.Label(root, text="端口号 ：")
port_label.grid(row=2, column=0, pady=5)
port_entry = tk.Entry(root)
port_entry.insert(0, port)
port_entry.grid(row=2, column=1, pady=5, padx=10)

test_checkbutton_var = tk.IntVar(value=test_checkbutton_var)
test_checkbutton = tk.Checkbutton(root, text="Test模式（不知道是什么就别开）", variable=test_checkbutton_var)
test_checkbutton.grid(row=10, column=0, pady=5, padx=10)

title_label = tk.Label(root, text="主题配置：勾选为启用，不勾选为禁用")
title_label.grid(row=3, column=0, pady=5)

topic1_entry = tk.Entry(root)
topic1_entry.insert(0, topic1) 
topic1_entry.grid(row=3, column=2, pady=5, padx=10)
topic1_checkbutton_var = tk.IntVar(value=topic1_checked)
topic1_checkbutton = tk.Checkbutton(root,text="主题 1（电脑开关机）：", variable=topic1_checkbutton_var)
topic1_checkbutton.grid(row=3, column=1)

topic2_entry = tk.Entry(root)
topic2_entry.insert(0, topic2) 
topic2_entry.grid(row=4, column=2, pady=5, padx=10)
topic2_checkbutton_var = tk.IntVar(value=topic2_checked)
topic2_checkbutton = tk.Checkbutton(root,text="主题 2（电脑屏幕亮度）：", variable=topic2_checkbutton_var)
topic2_checkbutton.grid(row=4, column=1)

topic3_entry = tk.Entry(root)
topic3_entry.insert(0, topic3) 
topic3_entry.grid(row=5, column=2, pady=5, padx=10)
topic3_checkbutton_var = tk.IntVar(value=topic3_checked)
topic3_checkbutton = tk.Checkbutton(root,text="主题 3（启动应用程序或者可执行文件）：", variable=topic3_checkbutton_var)
topic3_checkbutton.grid(row=5, column=1)

app_label = tk.Label(root, text="应用程序或者可执行文件目录 ：")
app_label.grid(row=6, column=1, pady=5)
app_entry = tk.Entry(root)
app_entry.insert(0, app)  # 设置默认值
app_entry.grid(row=6, column=2, pady=5, padx=10)

topic4_entry = tk.Entry(root)
topic4_entry.insert(0, topic4) 
topic4_entry.grid(row=7, column=2, pady=5, padx=10)
topic4_checkbutton_var = tk.IntVar(value=topic4_checked)
topic4_checkbutton = tk.Checkbutton(root,text="主题 4（启动应用程序或者可执行文件）：", variable=topic4_checkbutton_var)
topic4_checkbutton.grid(row=7, column=1)

app2_label = tk.Label(root, text="应用程序或者可执行文件目录 ：")
app2_label.grid(row=8, column=1, pady=5)
app2_entry = tk.Entry(root)
app2_entry.insert(0, app2)  # 设置默认值
app2_entry.grid(row=8, column=2, pady=5, padx=10)

topic5_entry = tk.Entry(root)
topic5_entry.insert(0, topic5) 
topic5_entry.grid(row=9, column=2, pady=5, padx=10)
topic5_checkbutton_var = tk.IntVar(value=topic5_checked)
topic5_checkbutton = tk.Checkbutton(root,text="主题 5（服务（需要管理员权限运行，否则无效））：", variable=topic5_checkbutton_var)
topic5_checkbutton.grid(row=9, column=1)

app3_label = tk.Label(root, text="服务名称 ：")
app3_label.grid(row=10, column=1, pady=5)
app3_entry = tk.Entry(root)
app3_entry.insert(0, app3)  # 设置默认值
app3_entry.grid(row=10, column=2, pady=5, padx=10)

open_button = tk.Button(root, text="打开配置文件夹", command=open_config)
open_button.grid(row=11, column=1, pady=5)

save_button = tk.Button(root, text="保存配置", command=save_config)
save_button.grid(row=11, column=2, pady=5)

save_button = tk.Button(root, text="普通测试", command=TEST)
save_button.grid(row=11, column=0, pady=5)

root.mainloop()