# 导入必要的库
import tkinter as tk
from tkinter import messagebox, filedialog
import json
import os

# 创建主窗口
root = tk.Tk()
root.title("配置修改")

# 尝试读取配置文件
config_file = "config.json"
config = {}
if os.path.exists(config_file):
    with open(config_file, "r", encoding="utf-8") as f:
        config = json.load(f)

# 系统配置部分
system_frame = tk.LabelFrame(root, text="系统配置")
system_frame.pack(padx=10, pady=5, fill="x")

tk.Label(system_frame, text="网站：").grid(row=0, column=0, sticky="e")
website_entry = tk.Entry(system_frame)
website_entry.grid(row=0, column=1, sticky="w")
website_entry.insert(0, config.get("broker", ""))

tk.Label(system_frame, text="密钥：").grid(row=1, column=0, sticky="e")
secret_entry = tk.Entry(system_frame)
secret_entry.grid(row=1, column=1, sticky="w")
secret_entry.insert(0, config.get("secret_id", ""))

tk.Label(system_frame, text="端口：").grid(row=2, column=0, sticky="e")
port_entry = tk.Entry(system_frame)
port_entry.grid(row=2, column=1, sticky="w")
port_entry.insert(0, str(config.get("port", "")))

test_var = tk.IntVar(value=config.get("test", 0))
test_check = tk.Checkbutton(system_frame, text="test 模式开关", variable=test_var)
test_check.grid(row=3, columnspan=2)

# 主题配置部分
theme_frame = tk.LabelFrame(root, text="主题配置")
theme_frame.pack(padx=10, pady=5, fill="x")

# 内置主题
builtin_themes = [
    {"nickname": "计算机", "key": "Computer", "name_var": tk.StringVar(), "checked": tk.IntVar()},
    {"nickname": "屏幕", "key": "screen", "name_var": tk.StringVar(), "checked": tk.IntVar()},
    {"nickname": "音量", "key": "volume", "name_var": tk.StringVar(), "checked": tk.IntVar()},
    {"nickname": "休眠", "key": "sleep", "name_var": tk.StringVar(), "checked": tk.IntVar()},
]

tk.Label(theme_frame, text="内置主题").grid(row=0, column=0, sticky="w", columnspan=2)
tk.Label(theme_frame, text="自定义主题").grid(row=0, column=3, sticky="w")

for idx, theme in enumerate(builtin_themes):
    theme_key = theme["key"]
    theme["name_var"].set(config.get(theme_key, ""))
    theme["checked"].set(config.get(f"{theme_key}_checked", 0))

    tk.Checkbutton(
        theme_frame,
        text=theme["nickname"],
        variable=theme["checked"]
    ).grid(row=idx+1, column=0, sticky="w", columnspan=2)
    tk.Entry(theme_frame, textvariable=theme["name_var"]).grid(row=idx+1, column=2, sticky="w")

# 自定义主题列表
custom_themes = []

# 自定义主题列表组件
custom_theme_list = tk.Listbox(theme_frame)
custom_theme_list.grid(row=1, column=3, rowspan=4, padx=10)

# 添加和修改按钮
tk.Button(theme_frame, text="添加", command=lambda: add_custom_theme(config)).grid(row=5, column=3, sticky="w")
tk.Button(theme_frame, text="修改", command=lambda: modify_custom_theme()).grid(
    row=5, column=3, sticky="e"
)

# 如果配置中有自定义主题，加载它们
def load_custom_themes():
    app_index = 1
    serve_index = 1
    # 加载程序类型的自定义主题
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
    # 加载服务类型的自定义主题
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

    tk.Label(theme_window, text="主题类型：").grid(row=0, column=0, sticky="e")
    theme_type_var = tk.StringVar(value=theme["type"])
    tk.OptionMenu(theme_window, theme_type_var, "程序", "服务").grid(row=0, column=1, sticky="w")

    tk.Label(theme_window, text="主题开关状态：").grid(row=1, column=0, sticky="e")
    theme_checked_var = tk.IntVar(value=theme["checked"])
    tk.Checkbutton(theme_window, variable=theme_checked_var).grid(row=1, column=1, sticky="w")

    tk.Label(theme_window, text="主题昵称：").grid(row=2, column=0, sticky="e")
    theme_nickname_entry = tk.Entry(theme_window)
    theme_nickname_entry.insert(0, theme["nickname"])
    theme_nickname_entry.grid(row=2, column=1, sticky="w")

    tk.Label(theme_window, text="主题名称：").grid(row=3, column=0, sticky="e")
    theme_name_entry = tk.Entry(theme_window)
    theme_name_entry.insert(0, theme["name"])
    theme_name_entry.grid(row=3, column=1, sticky="w")

    tk.Label(theme_window, text="主题值：").grid(row=4, column=0, sticky="e")
    theme_value_entry = tk.Entry(theme_window)
    theme_value_entry.insert(0, theme["value"])
    theme_value_entry.grid(row=4, column=1, sticky="w")

    def select_file():
        file_path = filedialog.askopenfilename()
        theme_value_entry.delete(0, tk.END)
        theme_value_entry.insert(0, file_path)

    tk.Button(theme_window, text="选择文件", command=select_file).grid(row=4, column=2, sticky="w")

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
        custom_themes.pop(index)
        custom_theme_list.delete(index)
        theme_window.destroy()

    tk.Button(theme_window, text="保存", command=save_theme).grid(row=5, column=0)
    tk.Button(theme_window, text="删除", command=delete_theme).grid(row=5, column=1)

# 添加自定义主题的函数中，也要更新显示
def add_custom_theme(config):
    theme_window = tk.Toplevel(root)
    theme_window.title("添加自定义主题")

    tk.Label(theme_window, text="主题类型：").grid(row=0, column=0, sticky="e")
    theme_type_var = tk.StringVar(value="程序")
    tk.OptionMenu(theme_window, theme_type_var, "程序", "服务").grid(row=0, column=1, sticky="w")

    tk.Label(theme_window, text="主题开关状态：").grid(row=1, column=0, sticky="e")
    theme_checked_var = tk.IntVar()
    tk.Checkbutton(theme_window, variable=theme_checked_var).grid(row=1, column=1, sticky="w")

    tk.Label(theme_window, text="主题昵称：").grid(row=2, column=0, sticky="e")
    theme_nickname_entry = tk.Entry(theme_window)
    theme_nickname_entry.grid(row=2, column=1, sticky="w")

    tk.Label(theme_window, text="主题名称：").grid(row=3, column=0, sticky="e")
    theme_name_entry = tk.Entry(theme_window)
    theme_name_entry.grid(row=3, column=1, sticky="w")

    tk.Label(theme_window, text="主题值：").grid(row=4, column=0, sticky="e")
    theme_value_entry = tk.Entry(theme_window)
    theme_value_entry.grid(row=4, column=1, sticky="w")

    def select_file():
        file_path = filedialog.askopenfilename()
        theme_value_entry.delete(0, tk.END)
        theme_value_entry.insert(0, file_path)

    tk.Button(theme_window, text="选择文件", command=select_file).grid(row=4, column=2, sticky="w")

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

    tk.Button(theme_window, text="保存", command=save_theme).grid(row=5, column=0, columnspan=2)

# 调用加载自定义主题的函数
load_custom_themes()
# 生成配置文件
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
    with open("config.json", "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)

    messagebox.showinfo("提示", "配置文件已保存")

tk.Button(root, text="保存配置文件", command=generate_config).pack(pady=10)
root.mainloop()