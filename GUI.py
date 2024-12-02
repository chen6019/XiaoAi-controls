"""打包命令
pyinstaller -F -n GUI --noconsole --hidden-import=win11toast --hidden-import=pywin32 --icon=icon.ico GUI.py
"""

import ctypes
import os
import shlex
from win11toast import notify
import json
import tkinter as tk
from tkinter import messagebox
import sys
import subprocess
import win32com.client

# 创建一个命名的互斥体
mutex = ctypes.windll.kernel32.CreateMutexW(None, False, "xagui_mutex")

# 检查互斥体是否已经存在
if ctypes.windll.kernel32.GetLastError() == 183:
    messagebox.showerror("错误", "应用程序已在运行。")
    sys.exit()


def TEST():
    save_config()
    if os.path.isfile("main.py"):
        subprocess.Popen(["python", "main.py"])
    elif os.path.isfile("XiaoAi-controls.exe"):
        subprocess.Popen(["XiaoAi-controls.exe"])
    else:
        messagebox.showerror("Error", "既没有找到 main.py 也没有找到Ai-controls.exe")


def open_config():
    os.startfile(appdata_path)


def on_closing():
    root.quit()
    # 释放互斥体
    ctypes.windll.kernel32.ReleaseMutex(mutex)
    sys.exit()


def save_config():
    mqtt_config = {
        "broker": broker_entry.get(),
        "secret_id": secret_id_entry.get(),
        "port": port_entry.get(),
        "test_test": test_checkbutton_var.get(),
    }
    for i in range(len(topics)):
        topic = topics[i].get()
        checked = checkbuttons_vars[i].get()
        nickname = nicknames[i].get()
        mqtt_config[topic] = value
        mqtt_config[f"{topic}_checked"] = checked
        mqtt_config[f"{topic}_nickname"] = nickname

    with open(config_path, "w") as f:
        json.dump(mqtt_config, f, indent=4)

    notify("Ok,配置保存成功")


def check_task():
    if check_task_exists("小爱控制"):
        auto_start_button.config(text="关开机自启", command=remove_auto_start)
    else:
        auto_start_button.config(text="开机自启", command=set_auto_start)
    auto_start_button.update_idletasks()


def set_auto_start():
    # 获取当前目录
    current_dir = os.path.dirname(os.path.abspath(sys.argv[0]))

    if os.path.isfile("main.py"):
        # 构造文件路径
        exe_path = os.path.join(current_dir, "main.py")
    elif os.path.isfile("XiaoAi-controls.exe"):
        # 构造文件路径
        exe_path = os.path.join(current_dir, "XiaoAi-controls.exe")
    else:
        messagebox.showerror("Error", "既没有找到 main.py 也没有找到Ai-controls.exe")

    # 引用 exe_path
    quoted_exe_path = shlex.quote(exe_path)
    # 创建任务计划
    result = subprocess.call(
        f'schtasks /Create /SC ONLOGON /TN "小爱控制" /TR "{quoted_exe_path}" /F',
        shell=True,
    )

    scheduler = win32com.client.Dispatch("Schedule.Service")
    scheduler.Connect()

    root_folder = scheduler.GetFolder("\\")

    task_definition = root_folder.GetTask("小爱控制").Definition

    # 设置任务以最高权限运行
    principal = task_definition.Principal
    principal.RunLevel = 1

    settings = task_definition.Settings

    settings.DisallowStartIfOnBatteries = False
    settings.StopIfGoingOnBatteries = False

    # 关闭任务超时停止
    settings.ExecutionTimeLimit = "PT0S"

    # 保存修改后的任务计划
    root_folder.RegisterTaskDefinition("小爱控制", task_definition, 6, "", "", 3)
    if result == 0:
        login_info_label.config(text="创建开机自启动成功")
    else:
        login_info_label.config(text="创建开机自启动失败")
    messagebox.showinfo("提示！", "移动位置后要重新设置哦！！")
    check_task()


def Get_administrator_privileges():
    save_config()
    messagebox.showinfo("提示！", "将以管理员权限重新运行！！\n已自动保存配置")
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, f'"{__file__}" set_title', None, 0
    )
    on_closing()


def remove_auto_start():
    # 弹出询问框，询问是否要删除任务
    if messagebox.askyesno("确定？", "你确定要删除开机自启动任务吗？"):
        delete_result = subprocess.call(
            'schtasks /Delete /TN "小爱控制" /F', shell=True
        )

    if delete_result == 0:
        login_info_label.config(text="关闭开机自启动成功")
    else:
        login_info_label.config(text="关闭开机自启动失败")
    check_task()


# 判断是否拥有管理员权限
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


# 检查是否有计划
def check_task_exists(task_name):
    scheduler = win32com.client.Dispatch("Schedule.Service")
    scheduler.Connect()
    root_folder = scheduler.GetFolder("\\")
    for task in root_folder.GetTasks(0):
        if task.Name == task_name:
            return True
    return False


# 获取APPDATA目录的路径
appdata_path = os.path.join(
    os.path.expanduser("~"), "AppData", "Roaming", "Ai-controls"
)

# 检查目录是否存在，如果不存在则创建目录
if not os.path.exists(appdata_path):
    os.makedirs(appdata_path)

# 创建mqtt_config.json文件的路径
config_path = os.path.join(appdata_path, "mqtt_config.json")

# 检查文件是否存在，如果不存在则创建文件
if not os.path.exists(config_path):
    mqtt_config = {
        "broker": "",
        "secret_id": "",
        "port": "",
        "test": 0,
        "Computer": "",
        "Computer_checked": 0,
        "screen": "",
        "screen_checked": 0,
        "volume": "",
        "volume_checked": 0,
        "sleep": "",
        "sleep_checked": 0,
    }
    with open(config_path, "w") as f:
        json.dump(mqtt_config, f)

# 从json文件中读取MQTT配置
with open(config_path, "r") as f:
    mqtt_config = json.load(f)

# 从MQTT配置中获取值并赋值给变量
broker = mqtt_config.get("broker", "")
secret_id = mqtt_config.get("secret_id", "")
port = mqtt_config.get("port", "")
test_checkbutton_var = mqtt_config.get("test_test", 0)

# 内置主题
builtin_topics = ["Computer", "screen", "volume", "sleep"]

# 动态主题
dynamic_topics = {}
for key, value in mqtt_config.items():
    if key.endswith("_checked") and key.replace("_checked", "") not in builtin_topics:
        topic = key.replace("_checked", "")
        dynamic_topics[topic] = {
            "checked": value,
            "value": mqtt_config.get(topic, ""),
            "nickname": mqtt_config.get(f"{topic}_nickname", ""),
        }

# 内置主题赋值
Computer = mqtt_config.get("Computer", "")
Computer_checked = mqtt_config.get("Computer_checked", 0)
screen = mqtt_config.get("screen", "")
screen_checked = mqtt_config.get("screen_checked", 0)
volume = mqtt_config.get("volume", "")
volume_checked = mqtt_config.get("volume_checked", 0)
sleep = mqtt_config.get("sleep", "")
sleep_checked = mqtt_config.get("sleep_checked", 0)

# 动态主题赋值
for topic, config in dynamic_topics.items():
    globals()[topic] = config["value"]
    globals()[f"{topic}_checked"] = config["checked"]
    globals()[f"{topic}_nickname"] = config["nickname"]

root = tk.Tk()

if is_admin():
    root.title("小爱控制 V1.0.1- 管理员")
else:
    root.title("小爱控制 V1.0.1")

root.protocol("WM_DELETE_WINDOW", on_closing)
# 主要配置0-10
broker_label = tk.Label(root, text="Broker （IOT平台）")
broker_label.grid(row=0, column=0, pady=5)
broker_entry = tk.Entry(root)
broker_entry.insert(0, broker)
broker_entry.grid(row=0, column=1, pady=5, padx=10)

login_info_label = tk.Label(root, font=("微软雅黑", 12), wraplength=370)
login_info_label.grid(row=0, column=2, pady=5, padx=10)

if is_admin():
    login_info_label.config(
        text="当前拥有“管理员权限”请谨慎操作\n点击“开机自启”或“关闭开机自启”设置"
    )
else:
    login_info_label.config(
        text="需要“管理员权限”才能设置开机自启\n点击“获取权限”按钮或“手动获取”权限"
    )

auto_start_button = tk.Button(root, font=("微软雅黑", 14))
auto_start_button.grid(row=1, column=2)


if is_admin():
    check_task()
else:
    auto_start_button.config(text="获取权限", command=Get_administrator_privileges)

secret_id_label = tk.Label(root, text="key（私钥）：")
secret_id_label.grid(row=1, column=0, pady=5)
secret_id_entry = tk.Entry(root, show="*")
secret_id_entry.insert(0, secret_id)
secret_id_entry.grid(row=1, column=1, pady=5, padx=10)

port_label = tk.Label(root, text="端口号 ：")
port_label.grid(row=2, column=0, pady=5)
port_entry = tk.Entry(root)
port_entry.insert(0, port)
port_entry.grid(row=2, column=1, pady=5, padx=10)

title_label = tk.Label(root, text="主题配置：勾选为启用，不勾选为禁用")
title_label.grid(row=2, column=2, pady=5)

topics = []
checkbuttons_vars = []
values = []
nicknames = []


def add_topic(is_builtin=False):
    row = len(topics) + 11
    topic_var = tk.StringVar()
    checkbutton_var = tk.IntVar()
    value_var = tk.StringVar()
    nickname_var = tk.StringVar()

    checkbutton = tk.Checkbutton(root, variable=checkbutton_var)
    checkbutton.grid(row=row, column=0)
    nickname_entry = tk.Entry(
        root, textvariable=nickname_var, state="readonly" if is_builtin else "normal"
    )
    nickname_entry.grid(row=row, column=1, pady=5, padx=10)
    if not is_builtin:
        topic_entry = tk.Entry(root, textvariable=topic_var)
        topic_entry.grid(row=row, column=2, pady=5, padx=10)
    value_entry = tk.Entry(root, textvariable=value_var)
    value_entry.grid(row=row, column=3, pady=5, padx=10)

    topics.append(topic_var)
    checkbuttons_vars.append(checkbutton_var)
    values.append(value_var)
    nicknames.append(nickname_var)


add_topic_button = tk.Button(root, text="添加主题", command=add_topic)
add_topic_button.grid(row=10, column=0, pady=5)

# 从配置文件中读取现有的主题和值
for key, value in mqtt_config.items():
    if key.endswith("_checked"):
        topic = key.replace("_checked", "")
        is_builtin = topic in builtin_topics
        add_topic(is_builtin)
        topics[-1].set(topic)
        checkbuttons_vars[-1].set(value)
        values[-1].set(mqtt_config.get(topic, ""))
        nicknames[-1].set(mqtt_config.get(f"{topic}_nickname", ""))

# 功能按钮
test_checkbutton_var = tk.IntVar(value=test_checkbutton_var)
test_checkbutton = tk.Checkbutton(
    root, text="Test模式（不知道是什么就别开）", variable=test_checkbutton_var
)
test_checkbutton.grid(row=81, column=0, pady=5, padx=10)

open_button = tk.Button(root, text="打开配置文件夹", command=open_config)
open_button.grid(row=90, column=0, pady=5)

save_button = tk.Button(root, text="保存配置", command=save_config)
save_button.grid(row=90, column=1, pady=5)

save_button = tk.Button(root, text="普通测试", command=TEST)
save_button.grid(row=90, column=2, pady=5)

root.mainloop()
# 释放互斥体
ctypes.windll.kernel32.ReleaseMutex(mutex)
