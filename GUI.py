"""打包命令
pyinstaller -F -n GUI --noconsole --icon=icon.ico GUI.py
"""

import os
import tkinter as tk
from tkinter import messagebox, filedialog
import json
import ctypes
import sys
import shlex
import subprocess
import win32com.client

# 创建一个命名的互斥体
mutex = ctypes.windll.kernel32.CreateMutexW(None, False, "XiaoAi-controls-GUI")

# 检查互斥体是否已经存在
if ctypes.windll.kernel32.GetLastError() == 183:
    messagebox.showerror("错误", "应用程序已在运行。")
    sys.exit()


# 判断是否拥有管理员权限
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


# 获取管理员权限
def get_administrator_privileges():
    messagebox.showinfo("提示！", "将以管理员权限重新运行！！")
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, f'"{__file__}"', None, 0
    )
    sys.exit()


# 设置窗口居中
def center_window(window):
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")


# 检查任务计划是否存在
def check_task_exists(task_name):
    scheduler = win32com.client.Dispatch("Schedule.Service")
    scheduler.Connect()
    root_folder = scheduler.GetFolder("\\")
    for task in root_folder.GetTasks(0):
        if task.Name == task_name:
            return True
    return False


# 设置开机自启动
def set_auto_start():
    exe_path = os.path.join(
        os.path.dirname(os.path.abspath(sys.argv[0])), "XiaoAi-controls.exe"
    )

    # 检查文件是否存在
    if not os.path.exists(exe_path):
        messagebox.showerror(
            "错误", "未找到 XiaoAi-controls.exe 文件\n请检查文件是否存在"
        )
        return

    quoted_exe_path = shlex.quote(exe_path)
    result = subprocess.call(
        f'schtasks /Create /SC ONLOGON /TN "小爱控制" /TR "{quoted_exe_path}" /F',
        shell=True,
    )

    scheduler = win32com.client.Dispatch("Schedule.Service")
    scheduler.Connect()
    root_folder = scheduler.GetFolder("\\")
    task_definition = root_folder.GetTask("小爱控制").Definition

    principal = task_definition.Principal
    principal.RunLevel = 1

    settings = task_definition.Settings
    settings.DisallowStartIfOnBatteries = False
    settings.StopIfGoingOnBatteries = False
    settings.ExecutionTimeLimit = "PT0S"

    root_folder.RegisterTaskDefinition("小爱控制", task_definition, 6, "", "", 3)
    if result == 0:
        messagebox.showinfo("提示", "创建开机自启动成功")
        messagebox.showinfo("提示！", "移动位置后要重新设置哦！！")
        check_task()
    else:
        messagebox.showerror("错误", "创建开机自启动失败")
        check_task()


# 移除开机自启动
def remove_auto_start():
    if messagebox.askyesno("确定？", "你确定要删除开机自启动任务吗？"):
        delete_result = subprocess.call(
            'schtasks /Delete /TN "小爱控制" /F', shell=True
        )
        if delete_result == 0:
            messagebox.showinfo("提示", "关闭开机自启动成功")
            check_task()
        else:
            messagebox.showerror("错误", "关闭开机自启动失败")
            check_task()


# 检查是否有计划任务并更新按钮状态
def check_task():
    if check_task_exists("小爱控制"):
        auto_start_button.config(text="关闭开机自启", command=remove_auto_start)
    else:
        auto_start_button.config(text="设置开机自启", command=set_auto_start)
    auto_start_button.update_idletasks()


# 鼠标双击事件处理程序
def on_double_click(event):
    modify_custom_theme()


# 如果配置中有自定义主题，加载它们
def load_custom_themes():
    app_index = 1
    serve_index = 1
    while True:
        app_key = f"application{app_index}"
        if app_key in config:
            theme = {
                "type": "程序",
                "checked": config.get(f"{app_key}_checked", 0),
                "nickname": config.get(f"{app_key}_name", ""),
                "name": config.get(app_key, ""),
                "value": config.get(f"{app_key}_directory{app_index}", ""),
            }
            custom_themes.append(theme)
            status = "开" if theme["checked"] else "关"
            display_name = theme["nickname"] or theme["name"]
            item_text = f"[{status}] {display_name}"
            custom_theme_list.insert(tk.END, item_text)
            app_index += 1
        else:
            break
    while True:
        serve_key = f"serve{serve_index}"
        if serve_key in config:
            theme = {
                "type": "服务",
                "checked": config.get(f"{serve_key}_checked", 0),
                "nickname": config.get(f"{serve_key}_name", ""),
                "name": config.get(serve_key, ""),
                "value": config.get(f"{serve_key}_value", ""),
            }
            custom_themes.append(theme)
            status = "开" if theme["checked"] else "关"
            display_name = theme["nickname"] or theme["name"]
            item_text = f"[{status}] {display_name}"
            custom_theme_list.insert(tk.END, item_text)
            serve_index += 1
        else:
            break


# 修改自定义主题的函数
def modify_custom_theme():
    selected = custom_theme_list.curselection()
    if not selected:
        messagebox.showwarning("警告", "请先选择一个自定义主题")
        return

    index = selected[0]
    theme = custom_themes[index]

    theme_window = tk.Toplevel(root)
    theme_window.title("修改自定义主题")
    theme_window.resizable(False, False)  # 禁用窗口大小调整

    tk.Label(theme_window, text="类型：").grid(row=0, column=0, sticky="e")
    tk.Label(theme_window, text="服务类型需要").grid(row=0, column=1, sticky="e")
    tk.Label(theme_window, text="管理员权限").grid(row=0, column=2, sticky="w")
    theme_type_var = tk.StringVar(value=theme["type"])
    tk.OptionMenu(theme_window, theme_type_var, "程序", "服务").grid(
        row=0, column=1, sticky="w"
    )

    tk.Label(theme_window, text="状态：").grid(row=1, column=0, sticky="e")
    theme_checked_var = tk.IntVar(value=theme["checked"])
    tk.Checkbutton(theme_window, variable=theme_checked_var).grid(
        row=1, column=1, sticky="w"
    )

    tk.Label(theme_window, text="昵称：").grid(row=2, column=0, sticky="e")
    theme_nickname_entry = tk.Entry(theme_window)
    theme_nickname_entry.insert(0, theme["nickname"])
    theme_nickname_entry.grid(row=2, column=1, sticky="w")

    tk.Label(theme_window, text="主题：").grid(row=3, column=0, sticky="e")
    theme_name_entry = tk.Entry(theme_window)
    theme_name_entry.insert(0, theme["name"])
    theme_name_entry.grid(row=3, column=1, sticky="w")

    tk.Label(theme_window, text="值：").grid(row=4, column=0, sticky="e")
    theme_value_entry = tk.Entry(theme_window)
    theme_value_entry.insert(0, theme["value"])
    theme_value_entry.grid(row=4, column=1, sticky="w")

    def select_file():
        file_path = filedialog.askopenfilename()
        theme_value_entry.delete(0, tk.END)
        theme_value_entry.insert(0, file_path)

    tk.Button(theme_window, text="选择文件", command=select_file).grid(
        row=4, column=2, sticky="w", padx=15
    )

    def save_theme():
        theme["type"] = theme_type_var.get()
        theme["checked"] = theme_checked_var.get()
        theme["nickname"] = theme_nickname_entry.get()
        theme["name"] = theme_name_entry.get()
        theme["value"] = theme_value_entry.get()
        status = "开" if theme["checked"] else "关"
        display_name = theme["nickname"] or theme["name"]
        item_text = f"[{status}] {display_name}"
        custom_theme_list.delete(index)
        custom_theme_list.insert(index, item_text)
        theme_window.destroy()

    def delete_theme():
        if messagebox.askyesno(
            "确认删除", "确定要删除这个自定义主题吗？", parent=theme_window
        ):
            custom_themes.pop(index)
            custom_theme_list.delete(index)
            theme_window.destroy()
        else:
            theme_window.lift()

    tk.Button(theme_window, text="保存", command=save_theme).grid(
        row=5, column=0, pady=15, padx=15
    )
    tk.Button(theme_window, text="删除", command=delete_theme).grid(row=5, column=1)

    center_window(theme_window)


# 添加自定义主题的函数中，也要更新显示
def add_custom_theme(config):
    theme_window = tk.Toplevel(root)
    theme_window.title("添加自定义主题")
    theme_window.resizable(False, False)  # 禁用窗口大小调整

    tk.Label(theme_window, text="类型：").grid(row=0, column=0, sticky="e")
    tk.Label(theme_window, text="服务类型需要").grid(row=0, column=1, sticky="e")
    tk.Label(theme_window, text="管理员权限").grid(row=0, column=2, sticky="w")
    theme_type_var = tk.StringVar(value="程序")
    tk.OptionMenu(theme_window, theme_type_var, "程序", "服务").grid(
        row=0, column=1, sticky="w"
    )

    tk.Label(theme_window, text="状态：").grid(row=1, column=0, sticky="e")
    theme_checked_var = tk.IntVar()
    tk.Checkbutton(theme_window, variable=theme_checked_var).grid(
        row=1, column=1, sticky="w"
    )

    tk.Label(theme_window, text="昵称：").grid(row=2, column=0, sticky="e")
    theme_nickname_entry = tk.Entry(theme_window)
    theme_nickname_entry.grid(row=2, column=1, sticky="w")

    tk.Label(theme_window, text="主题：").grid(row=3, column=0, sticky="e")
    theme_name_entry = tk.Entry(theme_window)
    theme_name_entry.grid(row=3, column=1, sticky="w")

    tk.Label(theme_window, text="值：").grid(row=4, column=0, sticky="e")
    theme_value_entry = tk.Entry(theme_window)
    theme_value_entry.grid(row=4, column=1, sticky="w")

    def select_file():
        file_path = filedialog.askopenfilename()
        theme_value_entry.delete(0, tk.END)
        theme_value_entry.insert(0, file_path)

    tk.Button(theme_window, text="选择文件", command=select_file).grid(
        row=4, column=2, sticky="w", padx=15
    )

    def save_theme():
        theme = {
            "type": theme_type_var.get(),
            "checked": theme_checked_var.get(),
            "nickname": theme_nickname_entry.get(),
            "name": theme_name_entry.get(),
            "value": theme_value_entry.get(),
        }
        custom_themes.append(theme)
        status = "开" if theme["checked"] else "关"
        display_name = theme["nickname"] or theme["name"]
        item_text = f"[{status}] {display_name}"
        custom_theme_list.insert(tk.END, item_text)
        theme_window.destroy()

    tk.Button(theme_window, text="保存", command=save_theme).grid(
        row=5, column=0, pady=15, padx=15
    )
    tk.Button(theme_window, text="取消", command=theme_window.destroy).grid(row=5, column=1)

    center_window(theme_window)


def generate_config():
    config = {
        "broker": website_entry.get(),
        "secret_id": secret_entry.get(),
        "port": int(port_entry.get()),
        "test": test_var.get(),
    }

    # 内置主题配置
    for theme in builtin_themes:
        key = theme["key"]
        value = theme["name_var"].get()
        config[key] = value
        config[f"{key}_checked"] = theme["checked"].get()

    # 自定义主题配置
    app_index = 1
    serve_index = 1
    for theme in custom_themes:
        if theme["type"] == "程序":
            prefix = f"application{app_index}"
            config[prefix] = theme["name"]
            config[f"{prefix}_name"] = theme["nickname"]
            config[f"{prefix}_checked"] = theme["checked"]
            config[f"{prefix}_directory{app_index}"] = theme["value"]
            app_index += 1
        else:
            prefix = f"serve{serve_index}"
            config[prefix] = theme["name"]
            config[f"{prefix}_name"] = theme["nickname"]
            config[f"{prefix}_checked"] = theme["checked"]
            config[f"{prefix}_value"] = theme["value"]
            serve_index += 1

    # 保存为 JSON 文件
    with open(config_file_path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)

    messagebox.showinfo("提示", "配置文件已保存")


# 打开配置文件夹的函数
def open_config_folder():
    os.startfile(appdata_dir)


# 获取用户的 AppData\Roaming 目录
appdata_dir = os.path.join(os.getenv("APPDATA"), "Ai-controls")

# 如果目录不存在则创建
if not os.path.exists(appdata_dir):
    os.makedirs(appdata_dir)

# 配置文件路径
config_file_path = os.path.join(appdata_dir, "config.json")

# 尝试读取配置文件
config = {}
if os.path.exists(config_file_path):
    with open(config_file_path, "r", encoding="utf-8") as f:
        config = json.load(f)

# 创建主窗口
root = tk.Tk()
root.title("小爱控制V1.1.0")

# 设置窗口最小值
root.wm_minsize(600, 650)  # 将宽度设置为600，高度设置为600

# 设置根窗口的行列权重
root.rowconfigure(0, weight=1)
root.rowconfigure(1, weight=1)
root.rowconfigure(2, weight=0)
root.columnconfigure(0, weight=1)

# 系统配置部分
system_frame = tk.LabelFrame(root, text="系统配置")
system_frame.grid(row=0, column=0, padx=10, pady=5, sticky="nsew")

# 使用 grid 布局，并设置权重
system_frame.grid(row=0, column=0, sticky="nsew")
for i in range(4):
    system_frame.rowconfigure(i, weight=1)
for j in range(3):
    system_frame.columnconfigure(j, weight=1)

tk.Label(system_frame, text="网站：").grid(row=0, column=0, sticky="e")
website_entry = tk.Entry(system_frame)
website_entry.grid(row=0, column=1, sticky="ew")
website_entry.insert(0, config.get("broker", ""))

tk.Label(system_frame, text="密钥：").grid(row=1, column=0, pady=10,sticky="e")
secret_entry = tk.Entry(system_frame)
secret_entry.grid(row=1, column=1, sticky="ew")
secret_entry.insert(0, config.get("secret_id", ""))

tk.Label(system_frame, text="端口：").grid(row=2, column=0, sticky="e")
port_entry = tk.Entry(system_frame)
port_entry.grid(row=2, column=1, sticky="ew")
port_entry.insert(0, str(config.get("port", "")))

test_var = tk.IntVar(value=config.get("test", 0))
test_check = tk.Checkbutton(system_frame, text="test", variable=test_var)
test_check.grid(row=3, column=0, columnspan=2, sticky="w")

# 添加设置开机自启动按钮上面的提示
auto_start_label = tk.Label(
    system_frame,
    text="需要管理员权限才能设置",
)
auto_start_label.grid(row=0, column=2, sticky="n")
auto_start_label1 = tk.Label(
    system_frame,
    text="开机自启和服务类型主题",
)
auto_start_label1.grid(row=1, column=2, sticky="n")

# 添加设置开机自启动按钮
auto_start_button = tk.Button(system_frame, text="", command=set_auto_start)
auto_start_button.grid(row=2, column=2,  sticky="n")

# 程序标题栏
if is_admin():
    check_task()
    root.title("小爱控制V1.1.0(管理员)")
else:
    auto_start_button.config(text="获取权限", command=get_administrator_privileges)
    # 隐藏test
    test_check.grid_remove()

# 主题配置部分
theme_frame = tk.LabelFrame(root, text="主题配置")
theme_frame.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")

# 使用 grid 布局，并设置权重
theme_frame.grid(row=1, column=0, sticky="nsew")
for i in range(6):
    theme_frame.rowconfigure(i, weight=1)
for j in range(4):
    theme_frame.columnconfigure(j, weight=1)

# 内置主题
builtin_themes = [
    {
        "nickname": "计算机",
        "key": "Computer",
        "name_var": tk.StringVar(),
        "checked": tk.IntVar(),
    },
    {
        "nickname": "屏幕",
        "key": "screen",
        "name_var": tk.StringVar(),
        "checked": tk.IntVar(),
    },
    {
        "nickname": "音量",
        "key": "volume",
        "name_var": tk.StringVar(),
        "checked": tk.IntVar(),
    },
    {
        "nickname": "睡眠",
        "key": "sleep",
        "name_var": tk.StringVar(),
        "checked": tk.IntVar(),
    },
]

tk.Label(theme_frame, text="内置").grid(row=0, column=0, sticky="w", columnspan=2)
tk.Label(theme_frame, text="主题").grid(row=0, column=2, sticky="w")
tk.Label(theme_frame, text="自定义(服务需要管理员权限)").grid(
    row=0, column=3, sticky="w"
)

for idx, theme in enumerate(builtin_themes):
    theme_key = theme["key"]
    theme["name_var"].set(config.get(theme_key, ""))
    theme["checked"].set(config.get(f"{theme_key}_checked", 0))

    tk.Checkbutton(theme_frame, text=theme["nickname"], variable=theme["checked"]).grid(
        row=idx + 1, column=0, sticky="w", columnspan=2
    )
    tk.Entry(theme_frame, textvariable=theme["name_var"]).grid(
        row=idx + 1, column=2, sticky="ew"
    )

# 自定义主题列表
custom_themes = []

# 自定义主题列表组件
custom_theme_list = tk.Listbox(theme_frame)
custom_theme_list.grid(row=1, column=3, rowspan=4, pady=10, sticky="nsew")

# 添加提示
tk.Label(theme_frame, text="双击即可修改").grid(row=5, column=3, pady=15, sticky="n")
# 添加和修改按钮
tk.Button(theme_frame, text="添加", command=lambda: add_custom_theme(config)).grid(
    row=5, column=3, sticky="w"
)
tk.Button(theme_frame, text="修改", command=lambda: modify_custom_theme()).grid(
    row=5, column=3, sticky="e"
)

# 绑定鼠标双击事件到自定义主题列表
custom_theme_list.bind("<Double-Button-1>", on_double_click)

# 添加按钮到框架中
button_frame = tk.Frame(root)
button_frame.grid(row=2, column=0, pady=15, sticky="ew")
button_frame.grid_rowconfigure(0, weight=1)
button_frame.grid_columnconfigure(0, weight=1)
button_frame.grid_columnconfigure(1, weight=1)

tk.Button(button_frame, text="打开配置文件夹", command=open_config_folder).grid(
    row=0, column=0, padx=20, sticky="e"
)
tk.Button(button_frame, text="保存配置文件", command=generate_config).grid(
    row=0, column=1, padx=20, sticky="w"
)

# 设置窗口在窗口大小变化时，框架自动扩展
root.rowconfigure(0, weight=1)
root.rowconfigure(1, weight=1)
root.rowconfigure(2, weight=0)
root.columnconfigure(0, weight=1)

# 设置窗口居中
center_window(root)

# 调用加载自定义主题的函数
load_custom_themes()

root.mainloop()

# 释放互斥体
ctypes.windll.kernel32.ReleaseMutex(mutex)

"""
GUI程序用来生成配置文件(用于Windows系统)
系统配置的内容为:
1.网站
2.密钥
3.端口
4.test模式开关(用于记录是否开启测试模式)
主题配置的内容为:
注:2和4和5是自定义主题才有,内置主题和自定义主题要分开显示,第一次打开无自定义主题
1.向服务器订阅用的主题名称
2.主题昵称
3.主题开关状态
4.主题值(需要填写路径,或调用系统api选择文件),是一个程序或文件(绝对路径)
5.主题类型(程序或者服务)下拉框选择(默认为程序)
自定义主题部分有两个按钮:添加和修改
添加按钮会弹出一个窗口,选择主题类型,主题开关状态,填写主题昵称,主题名称,主题值
修改按钮会弹出一个窗口,修改主题,有保存和删除按钮
"""
