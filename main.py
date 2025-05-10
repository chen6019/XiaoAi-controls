"""
依赖安装指令:
pip install paho-mqtt wmi win11toast pillow pystray comtypes pycaw
pip install --upgrade setuptools

注意:
python 13 使用虚拟环境需要将 C:\\Program Files\\Python313\\tcl\\tcl8.6 和 C:\\Program Files\\Python313\\tcl\\tk8.6 两个文件夹
复制到
C:\\Program Files\\Python313\\Lib 文件夹下
否则报错: _tkinter.TclError: Can't find a usable init.tcl

打包指令:
pyinstaller -F -n Remote-Controls --windowed --icon=icon.ico --add-data "icon.ico;."  main.py
程序名：Remote-Controls.exe
运行用户：SYSTEM or 当前登录用户（可能有管理员权限）
"""


#导入各种必要的模块
import io
import paho.mqtt.client as mqtt
import os
import pystray
from PIL import Image
import wmi
from win11toast import notify
import json
import logging
from tkinter import messagebox
import sys
import threading
import subprocess
import time
import ctypes
from ctypes import wintypes
import socket
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


# 创建一个命名的互斥体
mutex = ctypes.windll.kernel32.CreateMutexW(None, False, "Remote-Controls-main")

# 检查互斥体是否已经存在
if ctypes.windll.kernel32.GetLastError() == 183:
    messagebox.showerror("错误", "应用程序已在运行。")
    sys.exit()

"""
执行系统命令，并在超时后终止命令。

参数:
- cmd: 要执行的命令（字符串或列表）
- timeout: 命令执行的超时时间（秒，默认为30秒）

返回值:
- 命令终止的返回码
"""


def execute_command(cmd: str, timeout: int = 30) -> int:
    """
    English: Executes a system command with an optional timeout, terminating the command if it exceeds the timeout
    中文: 使用可选超时时间执行系统命令，如果超时则终止命令
    """
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


def on_subscribe(client, userdata, mid, reason_code_list, properties=None):
    """
    English: Callback when MQTT subscription completes
    中文: MQTT成功订阅后回调函数
    """
    for sub_result in reason_code_list:
        if isinstance(sub_result, int) and sub_result >= 128:
            logging.error(f"订阅失败:{reason_code_list}")
        else:
            logging.info(f"使用代码发送订阅申请成功：{mid}")


"""
MQTT取消订阅时的回调函数。

参数:
- client: MQTT客户端实例
- userdata: 用户数据
- mid: 消息ID
- reason_code_list: 取消订阅结果的状态码列表
- properties: 属性
"""


def on_unsubscribe(client, userdata: list, mid: int, reason_code_list: list, properties) -> None:
    """
    English: Callback when MQTT unsubscription completes
    中文: MQTT取消订阅后回调函数
    """
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


def set_brightness(value: int) -> None:
    """
    English: Sets the screen brightness to the specified value (0-100)
    中文: 设置屏幕亮度，取值范围为 0-100
    """
    try:
        wmi.WMI(namespace="wmi").WmiMonitorBrightnessMethods()[0].WmiSetBrightness(
            value, 0
        )
    except Exception as e:
        logging.error(f"无法设置亮度: {e}")


"""
设置音量。

参数:
- value: 音量值（0-100）
"""


def set_volume(value: int) -> None:
    """
    English: Sets the system volume to the specified value (0-100)
    中文: 设置系统音量，取值范围为 0-100
    """
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    # 控制音量在 0.0 - 1.0 之间
    volume.SetMasterVolumeLevelScalar(value / 100, None)  # type: ignore


def notify_in_thread(message: str) -> None:
    """
    English: Displays a Windows toast notification in a separate thread
    中文: 在单独线程中显示 Windows toast 通知
    """
    def notify_message():
        notify(message)

    thread = threading.Thread(target=notify_message)
    thread.daemon = True
    thread.start()


"""
根据接收到的命令和主题来处理相应的操作。

参数:
- command: 接收到的命令
- topic: 命令的主题
"""


def process_command(command: str, topic: str) -> None:
    """
    English: Handles the command received from MQTT messages based on the given topic
    中文: 根据主题处理从 MQTT 消息接收到的命令
    """
    logging.info(f"处理命令: {command} 主题: {topic}")

    # 先判断应用程序或服务类型主题
    for application, directory in applications:
        if topic == application:
            if command == "off":
                process_name = os.path.basename(directory)
                logging.info(f"尝试终止进程: {process_name}")
                notify_in_thread(f"尝试终止进程: {process_name}")
                result = subprocess.run(
                    ["taskkill", "/F", "/IM", process_name],
                    capture_output=True,
                    text=True,
                )
                logging.info(result.stdout)
            elif command == "on":
                if not directory or not os.path.isfile(directory):
                    notify_in_thread(f"启动失败，文件不存在: {directory}")
                    logging.error(f"启动失败，文件不存在: {directory}")
                    return
                subprocess.Popen(directory)
                logging.info(f"启动: {directory}")
                notify_in_thread(f"启动: {directory}")
            return
    
    def check_service_status(service_name):
        result = subprocess.run(["sc", "query", service_name], capture_output=True, text=True)
        if "RUNNING" in result.stdout:
            return "running"
        elif "STOPPED" in result.stdout:
            return "stopped"
        else:
            logging.error(f"无法获取服务{serve_name}的状态:{result.stderr}")
            return "unknown"
    
    for serve, serve_name in serves:
        if topic == serve:
            if command == "off":
                status = check_service_status(serve_name)
                if status == "unknown":
                    logging.error(f"无法获取服务{serve_name}的状态")
                    notify_in_thread(f"无法获取 {serve_name} 的状态,详情请查看日志")
                if status == "stopped":
                    logging.info(f"{serve_name} 还没有运行")
                    notify_in_thread(f"{serve_name} 还没有运行")
                else:
                    result = subprocess.run(["sc", "stop", serve_name], shell=True)
                    if result.returncode == 0:
                        logging.info(f"成功关闭 {serve_name}")
                        notify_in_thread(f"成功关闭 {serve_name}")
                    else:
                        logging.error(f"关闭 {serve_name} 失败")
                        logging.error(result.stderr)
                        notify_in_thread(f"关闭 {serve_name} 失败")
            elif command == "on":
                status = check_service_status(serve_name)
                if status == "unknown":
                    logging.error(f"无法获取服务{serve_name}的状态")
                    notify_in_thread(f"无法获取 {serve_name} 的状态,详情请查看日志")
                if status == "running":
                    logging.info(f"{serve_name} 已经在运行")
                    notify_in_thread(f"{serve_name} 已经在运行")
                else:
                    result = subprocess.run(["sc", "start", serve_name], shell=True)
                    if result.returncode == 0:
                        logging.info(f"成功启动 {serve_name}")
                        notify_in_thread(f"成功启动 {serve_name}")
                    else:
                        logging.error(f"启动 {serve_name} 失败")
                        logging.error(result.stderr)
                        notify_in_thread(f"启动 {serve_name} 失败")
            return

    # 若不匹配应用程序或服务，再判断是否为内置主题
    if topic == Computer:
        # 电脑开关机控制
        if command == "on":
            ctypes.windll.user32.LockWorkStation()
        elif command == "off":
            execute_command("shutdown -r -t 15")
            notify_in_thread("电脑将在15秒后重启")
    elif topic == screen:
        # 屏幕亮度控制
        if command == "off" or command == "1":
            set_brightness(0)
        elif command == "on":
            set_brightness(100)
        else:
            try:
                brightness = int(command[3:])
                set_brightness(brightness)
            except ValueError:
                notify_in_thread("亮度值无效")
                logging.error("亮度值无效")
    elif topic == volume:
        # 音量控制
        if command == "off" or command == "1":
            set_volume(0)
        elif command == "on":
            set_volume(100)
        else:
            try:
                volume_value = int(command[3:])
                set_volume(volume_value)
            except ValueError:
                notify_in_thread("音量值无效")
                logging.error("音量值无效")
            except Exception as e:
                notify_in_thread(f"设置音量时发生未知错误，请查看日志")
                logging.error(f"设置音量时出错: {e}")
    elif topic == sleep:
        if command == "off":
            notify_in_thread("当前还没有进入睡眠模式哦！")
        elif command == "on":
            execute_command("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
    else:
        # 未知主题
        notify_in_thread(f"未知主题: {topic}")
        logging.error(f"未知主题: {topic}")


"""
MQTT接收到消息时的回调函数。

参数:
- client: MQTT客户端实例
- userdata: 用户数据
- message: 接收到的消息
"""


def on_message(client, userdata: list, message) -> None:
    """
    English: Callback when an MQTT message is received
    中文: MQTT接收到消息时的回调函数
    """
    userdata.append(message.payload)
    command = message.payload.decode()
    logging.info(f"'{message.topic}' 主题收到 '{command}'")
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


def on_connect(client, userdata: list, flags: dict, reason_code, properties=None) -> None:
    # 兼容 int 和 ReasonCode 类型
    try:
        is_fail = reason_code.is_failure
    except AttributeError:
        is_fail = reason_code != 0

    if is_fail:
        notify_in_thread(
            f"连接MQTT失败: {reason_code}. 重新连接中..."
        )
        logging.error(f"连接失败: {reason_code}. loop_forever() 将重试连接")
    else:
        notify_in_thread(f"MQTT成功连接至{broker}")
        logging.info(f"连接到 {broker}")
        for key, value in config.items():
            if key.endswith("_checked") and value == 1:
                topic_key = key.replace("_checked", "")
                topic = config.get(topic_key)
                if topic:
                    client.subscribe(topic)
                    logging.info(f'订阅主题: "{topic}"')


"""
判断当前程序是否以管理员权限运行。

返回值:
- True: 以管理员权限运行
- False: 不是以管理员权限运行
"""


def is_admin() -> bool:
    """
    English: Checks whether the current process is running with administrator privileges
    中文: 检查当前进程是否以管理员权限运行
    """
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


"""
显示管理员权限提示。

无参数
无返回值
"""


def admin() -> None:
    """
    English: Opens a messagebox to show whether the current process has admin privileges
    中文: 弹出信息框展示当前进程是否拥有管理员权限
    """
    def show_message():
        if is_admin():
            messagebox.showinfo("信息", "已经拥有管理员权限")
        else:
            messagebox.showerror("错误", "没有管理员权限")

    thread2 = threading.Thread(target=show_message)
    thread2.daemon = True
    thread2.start()

   
def open_gui() -> None:
    """
    English: Attempts to open GUI.py or RC-GUI.exe, else shows an error message
    中文: 尝试运行 GUI.py 或 RC-GUI.exe，如果找不到则弹出错误提示
    """
    if os.path.isfile("GUI.py"):
        subprocess.Popen([".venv\\Scripts\\python.exe", "GUI.py"])
        notify_in_thread("正在打开配置窗口...")
    elif os.path.isfile("RC-GUI.exe"):
        subprocess.Popen(["RC-GUI.exe"])
        notify_in_thread("正在打开配置窗口...")
    else:
        def show_message():
            current_path = os.getcwd()
            messagebox.showerror(
                "Error", f"找不到GUI.py或RC-GUI.exe\n当前工作路径{current_path}"
            )
            logging.error(f"找不到GUI.py或RC-GUI.exe\n当前工作路径{current_path}")

        thread = threading.Thread(target=show_message)
        thread.daemon = True
        thread.start()


"""
退出程序。

无参数
无返回值
"""


def exit_program() -> None:
    """
    English: Stops the MQTT loop and exits the program
    中文: 停止 MQTT 循环，并退出程序
    """
    try:
        mqttc.loop_stop()
        mqttc.disconnect()
    except Exception as e:
        logging.error(f"程序停止时出错: {e}")
    finally:
        # 释放互斥体
        try:
            ctypes.windll.kernel32.ReleaseMutex(mutex)
            ctypes.windll.kernel32.CloseHandle(mutex)
            logging.info("互斥体已释放")
        except Exception as e:
            logging.error(f"释放互斥体时出错: {e}")
        
        logging.info("程序已停止")
        threading.Timer(0.5, lambda: os._exit(0)).start()
        sys.exit(0)
        
"""
托盘图标

"""
def tray() -> None:
    try:
        # 初始化系统托盘图标和菜单
        icon_path = resource_path("icon.ico")
        # 从资源文件中读取图像
        with open(icon_path, "rb") as f:
            image_data = f.read()
        icon = pystray.Icon("Remote-Controls", title="远程控制 V1.2.1")
        image = Image.open(io.BytesIO(image_data))
        menu = (
            pystray.MenuItem("打开配置", open_gui),
            pystray.MenuItem("管理员权限查询", admin),
            pystray.MenuItem("退出", exit_program),
        )
        icon.menu = menu
        icon.icon = image
        icon_Thread = threading.Thread(target=icon.run)
        icon_Thread.daemon = True
        icon_Thread.start()
        logging.info("托盘图标已加载完成")
    except Exception as e:
        messagebox.showerror(
            "Error", f"加载托盘图标时出错\n详情请查看日志"
        )
        logging.error(f"加载托盘图标时出错: {e}")


"""
文件过大时轮转文件。

参数:
- file_path: 文件路径
- max_size: 文件最大大小（字节），默认1MB
- backup_count: 保留的备份数量，默认1个
"""
def rotate_large_file(file_path: str, max_size: int = 1024 * 1024 * 1, backup_count: int = 1) -> None:
    """
    English: Rotates file if it's larger than the specified max_size
    中文: 如果文件大小超过限制则进行轮转备份
    """
    try:
        if os.path.exists(file_path) and os.path.getsize(file_path) > max_size:
            # 创建备份文件 (RC.log.back, etc.)
            for i in range(backup_count, 0, -1):
                if i > 1 and os.path.exists(f"{file_path}.{i-1}"):
                    if os.path.exists(f"{file_path}.{i}"):
                        os.remove(f"{file_path}.{i}")
                    os.rename(f"{file_path}.{i-1}", f"{file_path}.{i}")
            
            # 最新的日志文件变成 .back 备份
            if os.path.exists(f"{file_path}.back"):
                os.remove(f"{file_path}.back")
            os.rename(file_path, f"{file_path}.back")
            
            # 创建新的空日志文件
            with open(file_path, "w") as f:
                f.write(f"# 新日志文件创建于 {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"# 上一个日志文件备份为 {os.path.basename(file_path)}.back\n\n")
    except Exception as e:
        # 如果轮转失败，先确保日志文件存在
        if not os.path.exists(file_path):
            with open(file_path, "w") as f:
                f.write(f"# 新日志文件创建于 {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"# 警告：尝试轮转日志时出错: {e}\n\n")

# 获取资源文件的路径
def resource_path(relative_path):
    """获取资源文件的绝对路径"""
    # PyInstaller 创建临时文件夹
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base_path, relative_path)

# 获取应用程序的路径
if getattr(sys, "frozen", False):
    application_path = os.path.dirname(sys.executable)
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

# 改变当前的工作路径
os.chdir(application_path)

# 配置文件和日志文件路径改为当前工作目录
appdata_path = os.path.abspath(os.path.dirname(sys.argv[0]))

log_path = os.path.join(appdata_path, "RC.log")
config_path = os.path.join(appdata_path, "config.json")

# 先进行日志文件轮转，再配置日志系统
rotate_large_file(log_path)

# 日志和配置文件路径处理
logging.basicConfig(
    filename=log_path,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s: %(message)s",
    datefmt="%m/%d/%Y %H:%M:%S",
    encoding='utf-8'
)

# 记录程序启动信息
logging.info("=" * 50)
logging.info("程序启动")
logging.info(f"当前工作目录: {os.getcwd()}")
logging.info(f"日志文件路径: {log_path}")
logging.info(f"配置文件路径: {config_path}")
logging.info(f"Python版本: {sys.version}")
logging.info("=" * 50)

# 检查配置文件是否存在
if os.path.exists(config_path):
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
    except json.JSONDecodeError:
        messagebox.showerror("Error", "配置文件格式错误\n请检查config.json文件")
        logging.error("config.json 文件格式错误")
        open_gui()
        threading.Timer(0.5, lambda: os._exit(0)).start()
        sys.exit()
else:
    messagebox.showerror("Error", "配置文件不存在\n请先打开RC-GUI配置文件")
    logging.error("config.json 文件不存在")
    open_gui()
    threading.Timer(0.5, lambda: os._exit(0)).start()
    sys.exit()


# 确保config已经定义后再继续
if config.get("test") == 1:
    logging.warning("开启测试模式:可以不启用任何主题")
else:
    if (
        all(
            config.get(f"{key}_checked", 0) == 0
            for key in ["Computer", "screen", "volume", "sleep"]
        )
        and all(
            config.get(f"application{index}_checked", 0) == 0
            for index in range(1, 100)
        )
        and all(
            config.get(f"serve{index}_checked", 0) == 0 for index in range(1, 100)
        )
    ):
        logging.error("没有启用任何主题，显示错误信息")
        messagebox.showerror("Error", "主题不能一个都没有吧！\n（除了测试模式）")
        open_gui()
        threading.Timer(0.5, lambda: os._exit(0)).start()
    else:
        logging.info("至少已有一个主题被启用")

broker = config.get("broker")
secret_id = config.get("secret_id")
port = int(config.get("port"))


# 动态加载主题
def load_theme(key):
    """
    English: Loads the theme from config if it is enabled
    中文: 如果主题被勾选启用，则从配置中加载该主题
    """
    return config.get(key) if config.get(f"{key}_checked") == 1 else None


Computer = load_theme("Computer")
screen = load_theme("screen")
volume = load_theme("volume")
sleep = load_theme("sleep")

# 加载应用程序主题到应用程序列表
applications = []
for i in range(1, 50):
    app_key = f"application{i}"
    directory_key = f"application{i}_directory{i}"
    application = load_theme(app_key)
    directory = config.get(directory_key) if application else None
    if application:
        logging.info(f"加载应用程序: {app_key}, 目录: {directory}")
        applications.append((application, directory))
logging.info(f"读取的应用程序列表: {applications}\n")

# 加载服务主题到服务列表
serves = []
for i in range(1, 50):
    serve_key = f"serve{i}"
    serve_name_key = f"serve{i}_value"
    serve = load_theme(serve_key)
    serve_name = config.get(serve_name_key) if serve else None
    if serve:
        logging.info(f"加载服务: {serve_key}, 名称: {serve_name}")
        serves.append((serve, serve_name))
logging.info(f"读取的服务列表: {serves}\n")

# 如果主题不为空，将其记录到日志中
for key in ["Computer", "screen", "volume", "sleep"]:
    if config.get(key):
        logging.info(f'主题"{config.get(key)}"')

for application, directory in applications:
    logging.info(f'主题"{application}"，值："{directory}"')

for serve, serve_name in serves:
    logging.info(f'主题"{serve}"，值："{serve_name}"')

# tray()  # 托盘图标

# 初始化MQTT客户端
mqttc = mqtt.Client(callback_api_version=mqtt.CallbackAPIVersion.VERSION2) # type: ignore
mqttc.on_connect = on_connect
mqttc.on_message = on_message
mqttc.on_subscribe = on_subscribe
mqttc.on_unsubscribe = on_unsubscribe

mqttc.user_data_set([])
mqttc._client_id = secret_id
try:
    mqttc.connect(broker, port)
except socket.timeout:
    messagebox.showerror(
        "Error", "连接到 MQTT 服务器超时，请检查网络连接或服务器地址，端口号！"
    )
    open_gui()
    threading.Timer(0.5, lambda: os._exit(0)).start()
except socket.gaierror:
    messagebox.showerror(
        "Error", "无法解析 MQTT 服务器地址，请重试或检查服务器地址是否正确！"
    )
    open_gui()
    threading.Timer(0.5, lambda: os._exit(0)).start()
try:
    mqttc.loop_forever()
except KeyboardInterrupt:
    logging.warning("收到中断,程序停止")
    notify_in_thread("收到中断信号\n程序停止")
    exit_program()
except Exception as e:
    logging.error(f"程序异常: {e}")
    exit_program()

logging.info(f"总共收到以下消息: {mqttc.user_data_get()}")

# 确保在程序退出前释放互斥体
try:
    ctypes.windll.kernel32.ReleaseMutex(mutex)
    ctypes.windll.kernel32.CloseHandle(mutex)
except Exception as e:
    logging.error(f"释放互斥体时出错: {e}")

