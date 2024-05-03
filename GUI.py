import os
import json
import tkinter as tk
from tkinter import messagebox
import sys


def on_closing():
    root.quit()
    sys.exit()
def save_config():
    broker = broker_entry.get()
    topic1 = topic1_entry.get()
    topic2 = topic2_entry.get()
    topic3 = topic3_entry.get()
    secret_id = secret_id_entry.get()
    port = port_entry.get()

    # 尝试将port的值转换为整数
    try:
        port = int(port)
    except ValueError:
        messagebox.showerror("Error", "Port must be an integer")
        return

    if not broker or not topic1 or not topic2 or not topic3 or not secret_id or not port:
        messagebox.showerror("Error", "All fields must be filled")
        return

    mqtt_config = {
        'broker': broker,
        'topic1': topic1,
        'topic2': topic2,
        'topic3': topic3,
        'secret_id': secret_id,
        'port': port  # port现在是一个整数
    }

    with open(config_path, 'w') as f:
        json.dump(mqtt_config, f)

    messagebox.showinfo("Info", "Configuration saved successfully")

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
        'topic1': '',
        'topic2': '',
        'topic3': '',
        'secret_id': '',
        'port': ''
    }
    with open(config_path, 'w') as f:
        json.dump(mqtt_config, f)

# 从json文件中读取MQTT配置
with open(config_path, 'r') as f:
    mqtt_config = json.load(f)

# 从MQTT配置中获取值并赋值给变量
broker = mqtt_config['broker']
topic1 = mqtt_config['topic1']
topic2 = mqtt_config['topic2']
topic3 = mqtt_config['topic3']
secret_id = mqtt_config['secret_id']
port = mqtt_config['port']


root = tk.Tk()

root.protocol("WM_DELETE_WINDOW", on_closing)

broker_label = tk.Label(root, text="Broker")
broker_label.pack()
broker_entry = tk.Entry(root)
broker_entry.insert(0, broker)  # 设置默认值
broker_entry.pack()

topic1_label = tk.Label(root, text="Topic 1")
topic1_label.pack()
topic1_entry = tk.Entry(root)
topic1_entry.insert(0, topic1)  # 设置默认值
topic1_entry.pack()

topic2_label = tk.Label(root, text="Topic 2")
topic2_label.pack()
topic2_entry = tk.Entry(root)
topic2_entry.insert(0, topic2)  # 设置默认值
topic2_entry.pack()

topic3_label = tk.Label(root, text="Topic 3")
topic3_label.pack()
topic3_entry = tk.Entry(root)
topic3_entry.insert(0, topic3)  # 设置默认值
topic3_entry.pack()

secret_id_label = tk.Label(root, text="Secret ID")
secret_id_label.pack()
secret_id_entry = tk.Entry(root)
secret_id_entry.insert(0, secret_id)  # 设置默认值
secret_id_entry.pack()

port_label = tk.Label(root, text="Port")
port_label.pack()
port_entry = tk.Entry(root)
port_entry.insert(0, port)  # 设置默认值
port_entry.pack()

save_button = tk.Button(root, text="Save Configuration", command=save_config)
save_button.pack()

root.mainloop()
