# tray.py
import os
import sys
import subprocess
import threading
import time
import ctypes
import pystray
from PIL import Image
import psutil

# 配置
MAIN_EXE = "Remote-Controls.exe" if getattr(sys, "frozen", False) else "main.py"
GUI_EXE = "RC-GUI.exe"
GUI_PY = "GUI.py"
ICON_FILE = "icon.ico"
MUTEX_NAME = "Remote-Controls-main"

def resource_path(relative_path):
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base_path, relative_path)

def is_main_running():
    # 检查互斥体是否存在
    mutex = ctypes.windll.kernel32.OpenMutexW(0x100000, False, MUTEX_NAME)
    if mutex:
        ctypes.windll.kernel32.CloseHandle(mutex)
        return True
    return False

def get_main_proc():
    # 查找主程序进程
    for p in psutil.process_iter(['pid', 'name', 'exe', 'cmdline']):
        try:
            if MAIN_EXE.endswith('.exe'):
                if p.info['name'] == MAIN_EXE or (p.info['exe'] and os.path.basename(p.info['exe']) == MAIN_EXE):
                    return p
            else:
                if p.info['cmdline'] and MAIN_EXE in p.info['cmdline']:
                    return p
        except Exception:
            continue
    return None

def is_main_admin():
    proc = get_main_proc()
    if not proc:
        return False
    try:
        return proc.username().endswith("SYSTEM") or ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False

main_process = None

def start_main():
    global main_process
    if is_main_running():
        notify("主程序已在运行")
        return
    if MAIN_EXE.endswith('.exe') and os.path.exists(MAIN_EXE):
        main_process = subprocess.Popen([MAIN_EXE])
    elif os.path.exists(MAIN_EXE):
        main_process = subprocess.Popen([sys.executable, MAIN_EXE])
    else:
        notify("未找到主程序")
        return
    notify("主程序已启动")

def stop_main():
    proc = get_main_proc()
    if not proc:
        notify("主程序未运行")
        return
    try:
        proc.terminate()
        proc.wait(5)
        notify("主程序已关闭")
    except Exception:
        notify("关闭主程序失败")

def restart_main():
    stop_main()
    time.sleep(1)
    start_main()

def restart_main_as_admin():
    """以管理员权限重启主程序"""
    if is_main_running():
        stop_main()
        time.sleep(1)
    
    # 确定要启动的程序
    args = None
    if MAIN_EXE.endswith('.exe') and os.path.exists(MAIN_EXE):
        program = os.path.abspath(MAIN_EXE)
    elif os.path.exists(MAIN_EXE):
        program = sys.executable
        args = os.path.abspath(MAIN_EXE)
    else:
        notify("未找到主程序")
        return
    
    try:
        # 使用 ShellExecute 以管理员权限启动程序
        if MAIN_EXE.endswith('.exe'):
            ctypes.windll.shell32.ShellExecuteW(None, "runas", program, None, None, 1)
        else:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", program, f'"{args}"', None, 1)
        notify("已以管理员权限启动主程序")
    except Exception as e:
        notify(f"启动失败: {e}")

def check_admin():
    if is_main_running() and is_main_admin():
        notify("主程序已获得管理员权限")
    elif is_main_running():
        notify("主程序未获得管理员权限")
    else:
        notify("主程序未运行")

def open_gui():
    if os.path.exists(GUI_EXE):
        subprocess.Popen([GUI_EXE])
    elif os.path.exists(GUI_PY):
        subprocess.Popen([sys.executable, GUI_PY])
    else:
        notify("未找到配置界面")

def notify(msg):
    # 简单的Windows通知
    try:
        from win11toast import notify as toast
        threading.Thread(target=lambda: toast(msg)).start()
    except Exception:
        print(msg)

def on_open(icon, item):
    start_main()

def on_restart(icon, item):
    restart_main()

def on_restart_as_admin(icon, item):
    restart_main_as_admin()

def on_stop(icon, item):
    stop_main()

def on_check_admin(icon, item):
    check_admin()

def on_open_gui(icon, item):
    open_gui()

def on_exit(icon, item):
    icon.stop()
    sys.exit(0)

def tray_main():
    icon_path = resource_path(ICON_FILE)
    image = Image.open(icon_path) if os.path.exists(icon_path) else None
    menu = pystray.Menu(
        pystray.MenuItem("打开主程序", on_open),
        pystray.MenuItem("重启主程序", on_restart),
        pystray.MenuItem("以管理员权限重启主程序", on_restart_as_admin),
        pystray.MenuItem("关闭主程序", on_stop),
        pystray.MenuItem("主程序管理员权限", on_check_admin),
        pystray.MenuItem("打开配置界面", on_open_gui),
        pystray.MenuItem("退出托盘", on_exit),
    )
    icon = pystray.Icon("Remote-Controls-Tray", image, "远程控制托盘", menu)
    icon.run()

if __name__ == "__main__":
    tray_main()