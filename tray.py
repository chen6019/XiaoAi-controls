"""
程序文件名RC-tray.exe
运行用户：当前登录用户（可能有管理员权限）
pyinstaller -F -n RC-tray --windowed --icon=icon.ico --add-data "icon.ico;."  tray.py
"""


from email import message
import os
import sys
import subprocess
import threading
import time
import ctypes
import logging
from tkinter import messagebox
import pystray
from PIL import Image
import psutil

# 配置
MAIN_EXE = "Remote-Controls.exe" if getattr(sys, "frozen", False) else "main.py"
GUI_EXE = "RC-GUI.exe"
GUI_PY = "GUI.py"
ICON_FILE = "icon.ico"
MUTEX_NAME = "Remote-Controls-main"
# 日志配置
appdata_dir = os.path.abspath(os.path.dirname(sys.argv[0]))
tray_log_path = os.path.join(appdata_dir, "tray.log")
logging.basicConfig(
    filename=tray_log_path,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s: %(message)s",
    datefmt="%m/%d/%Y %H:%M:%S",
    encoding='utf-8'
)

def resource_path(relative_path):
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base_path, relative_path)


def is_main_running():
    # 优先按进程检测主程序是否已运行
    if get_main_proc():
        return True
    # 回退到互斥体判断
    mutex = ctypes.windll.kernel32.OpenMutexW(0x100000, False, MUTEX_NAME)
    if mutex:
        ctypes.windll.kernel32.CloseHandle(mutex)
        return True
    return False

def get_main_proc():
    """查找主程序进程，同时处理普通权限和管理员权限的情况"""
    # 方法1：使用psutil尝试直接获取进程信息
    for p in psutil.process_iter(['pid', 'name', 'exe', 'cmdline']):
        try:
            if MAIN_EXE.endswith('.exe'):
                if p.info['name'] == MAIN_EXE or (p.info['exe'] and os.path.basename(p.info['exe']) == MAIN_EXE):
                    return p
            else:
                if p.info['cmdline'] and MAIN_EXE in ' '.join(p.info['cmdline']):
                    return p
        except Exception:
            continue
            
    # 方法2：使用wmic命令行工具查找进程，可以检测管理员权限运行的进程
    try:
        if MAIN_EXE.endswith('.exe'):
            process_name = MAIN_EXE
            cmd = f'wmic process where "name=\'{process_name}\'" get ProcessId /value'
            result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
            if "ProcessId=" in result.stdout:
                pid_lines = [line for line in result.stdout.strip().split('\n') if "ProcessId=" in line]
                for pid_line in pid_lines:
                    try:
                        pid = int(pid_line.strip().split("=")[1])
                        # 尝试通过pid获取进程对象
                        return psutil.Process(pid)
                    except (psutil.NoSuchProcess, ValueError):
                        continue
        else:
            # 如果是Python脚本，尝试查找所有python进程并检查命令行
            cmd = 'wmic process where "name=\'python.exe\' or name=\'pythonw.exe\'" get ProcessId,CommandLine /value'
            result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
            lines = result.stdout.strip().split('\n')
            
            current_pid = None
            current_cmd = None
            
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                    
                if line.startswith("CommandLine="):
                    current_cmd = line[12:]
                elif line.startswith("ProcessId="):
                    current_pid = line[10:]
                    
                    if current_cmd and current_pid and MAIN_EXE in current_cmd:
                        try:
                            return psutil.Process(int(current_pid))
                        except (psutil.NoSuchProcess, ValueError):
                            pass
                        
                    current_pid = None
                    current_cmd = None
            
    except Exception as e:
        logging.error(f"使用wmic查找进程失败: {e}")
        
    return None
        

def is_main_admin():
    """检查主程序是否以管理员权限运行"""
    proc = get_main_proc()
    if not proc:
        return False
    try:
        # 检查进程用户
        if hasattr(proc, 'username') and callable(proc.username):
            username = proc.username()
            if username.endswith("SYSTEM") or "Administrator" in username:
                return True
        
        # 如果无法确定用户，尝试其他方法
        # 使用WMIC检查进程权限
        if hasattr(proc, 'pid'):
            try:
                cmd = f'wmic process where ProcessId={proc.pid} get ExecutablePath /value'
                result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
                exec_path = ""
                if "ExecutablePath=" in result.stdout:
                    exec_path = result.stdout.strip().split("=")[1]
                    
                # 检查程序是否在系统目录或Program Files
                system_dirs = ["\\Windows\\", "\\Program Files\\", "\\Program Files (x86)\\"]
                if any(sys_dir in exec_path for sys_dir in system_dirs):
                    return True
            except Exception:
                pass
                
        # 如果上述方法都无法确定，检查互斥体
        mutex = ctypes.windll.kernel32.OpenMutexW(0x1F0001, False, MUTEX_NAME)
        if mutex:
            ctypes.windll.kernel32.CloseHandle(mutex)
            # 如果我们能打开互斥体但psutil无法完全访问进程，可能是权限问题
            if not (hasattr(proc, 'username') and callable(proc.username)):
                return True
    except Exception as e:
        logging.error(f"检查管理员权限时出错: {e}")
    
    return False

main_process = None

def clean_orphaned_mutex():
    """清理可能未被正确释放的主程序互斥体"""
    try:
        # 先尝试以普通权限清理
        mutex = ctypes.windll.kernel32.CreateMutexW(None, False, MUTEX_NAME)
        if mutex:
            # 释放并关闭互斥体
            ctypes.windll.kernel32.ReleaseMutex(mutex)
            ctypes.windll.kernel32.CloseHandle(mutex)
            logging.info("成功以普通权限清理互斥体")
            return True
            
        # 如果普通权限清理失败，可能是权限问题，尝试提权清理
        logging.warning("普通权限清理互斥体失败，尝试提权清理")
        
        # 创建清理脚本
        script_content = f"""
        @echo off
        echo 正在尝试清理互斥体...
        
        REM 使用PowerShell尝试清理互斥体
        powershell -Command "$mutex = New-Object System.Threading.Mutex($false, '{MUTEX_NAME}'); if ($mutex) {{ $mutex.Close(); $mutex.Dispose(); Write-Host 'PowerShell清理成功' }}"
        
        REM 使用系统API尝试重置事件
        handle -c -y "{MUTEX_NAME}"
        
        exit
        """
        
        temp_script = os.path.join(os.environ.get('TEMP', '.'), 'clean_mutex.bat')
        with open(temp_script, 'w') as f:
            f.write(script_content)
        
        # 以管理员权限运行清理脚本
        result = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", "cmd.exe", f"/c {temp_script}", None, 0
        )
        
        if result > 32:  # 成功启动
            time.sleep(2)  # 等待脚本执行
            # 再次检查互斥体是否已被清理
            mutex = ctypes.windll.kernel32.CreateMutexW(None, False, MUTEX_NAME)
            if mutex:
                ctypes.windll.kernel32.ReleaseMutex(mutex)
                ctypes.windll.kernel32.CloseHandle(mutex)
                logging.info("提权清理互斥体成功")
                return True
            else:
                # 如果仍然无法创建，可能互斥体仍然存在
                if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
                    logging.warning("提权清理后互斥体仍然存在")
                else:
                    logging.error(f"提权后检查互斥体状态失败: {ctypes.windll.kernel32.GetLastError()}")
                return False
        else:
            logging.error(f"提权清理互斥体失败: {result}")
            return False
    except Exception as e:
        logging.error(f"清理互斥体失败: {e}")
        notify(f"清理互斥体失败，详情请查看日志: {tray_log_path}")
    return False

def start_main():
    global main_process
    logging.info("开始启动主程序")
    # 检查进程是否已在运行
    if get_main_proc():
        notify("主程序已在运行")
        return
    
    # 如果进程未运行但之前检测到互斥体存在，则可能是互斥体未被正确释放
    # 尝试清理互斥体
    clean_orphaned_mutex()
    
    if MAIN_EXE.endswith('.exe') and os.path.exists(MAIN_EXE):
        main_process = subprocess.Popen([MAIN_EXE], creationflags=subprocess.CREATE_NO_WINDOW)
    elif os.path.exists(MAIN_EXE):
        main_process = subprocess.Popen([sys.executable, MAIN_EXE], creationflags=subprocess.CREATE_NO_WINDOW)
    else:
        logging.error("未找到“Remote-Controls”程序")
        messagebox.showerror("Error","未找到“Remote-Controls”程序")
        return
    logging.info("主程序启动成功")
    notify("主程序已启动")

def stop_main():
    """停止主程序：使用taskkill命令关闭，类似于批处理文件的实现"""
    logging.info("开始关闭主程序")
    proc = get_main_proc()
    if not proc:
        messagebox.showerror("Error","主程序未运行")
        return

    # 检查是否为管理员权限进程
    is_admin_proc = is_main_admin()
    
    try:
        if is_admin_proc:
            logging.info("主程序以管理员权限运行，尝试提权关闭...")
            
            # 创建与批处理文件类似的脚本
            script_content = f"""
            @echo off
            echo 正在尝试关闭主程序...
            taskkill /im "{MAIN_EXE}" /f
            if %errorlevel% equ 0 (
                echo 成功关闭进程 "{MAIN_EXE}".
            ) else (
                echo 进程 "{MAIN_EXE}" 未运行或关闭失败.
            )
            exit
            """
            
            temp_script = os.path.join(os.environ.get('TEMP', '.'), 'stop_main.bat')
            with open(temp_script, 'w') as f:
                f.write(script_content)
            
            # 以管理员权限运行批处理
            result = ctypes.windll.shell32.ShellExecuteW(
                None, "runas", "cmd.exe", f"/c {temp_script}", None, 0
            )
            
            if result > 32:  # 成功启动
                time.sleep(3)  # 等待批处理执行完成
                if not is_main_running():
                    notify("主程序已关闭")
                else:
                    # 如果仍在运行，尝试更直接的方法
                    notify("主程序仍在运行，请尝试手动关闭（任务管理器）")
                    # force_stop_main()
            else:
                logging.warning(f"提权失败，尝试使用常规方法关闭 (错误码: {result})")
                notify("无法获取管理员权限，尝试使用常规方法关闭")
                
                # 回退到常规方法
                subprocess.run(f'taskkill /im "{MAIN_EXE}" /f', shell=True)
        else:
            # 正常关闭方法
            subprocess.run(f'taskkill /im "{MAIN_EXE}" /f', shell=True)
            notify("主程序已关闭")
    except Exception as e:
        logging.error(f"关闭主程序失败: {e}")
        messagebox.showerror("Error",f"关闭主程序失败,详情请查看日志{tray_log_path}")
    
    # 不管成功与否，都尝试清理互斥体
    time.sleep(1)
    clean_orphaned_mutex()

def restart_main():
    logging.info("重启主程序")
    stop_main()
    time.sleep(1)
    start_main()

def restart_main_as_admin():
    """以管理员权限重启主程序"""
    logging.info("以管理员权限重启主程序")
    # 检查进程是否在运行，如果在运行则关闭
    if get_main_proc():
        stop_main()
        time.sleep(1)
    
    # 无论进程是否在运行，都尝试清理可能残留的互斥体
    clean_orphaned_mutex()
    
    # 确定要启动的程序
    args = None
    if MAIN_EXE.endswith('.exe') and os.path.exists(MAIN_EXE):
        program = os.path.abspath(MAIN_EXE)
    elif os.path.exists(MAIN_EXE):
        program = sys.executable
        args = os.path.abspath(MAIN_EXE)
    else:
        messagebox.showerror("Error","未找到“Remote-Controls”程序")
        return
    
    try:
        # 使用 ShellExecute 以管理员权限启动程序
        if MAIN_EXE.endswith('.exe'):
            ctypes.windll.shell32.ShellExecuteW(None, "runas", program, None, None, 0)
        else:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", program, f'"{args}"', None, 0)
        notify("已以管理员权限启动主程序")
    except Exception as e:
        logging.error(f"启动失败: {e}")
        messagebox.showerror("Error",f"启动失败，详情请查看日志: {tray_log_path}")

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
    logging.info(f"通知: {msg}")
    try:
        from win11toast import notify as toast
        threading.Thread(target=lambda: toast(msg)).start()
    except Exception as e:
        logging.error(f"通知失败: {e}")
        print(msg)

def stop_tray():
    """关闭托盘程序自身"""
    logging.info("关闭托盘程序自身")
    notify("正在关闭托盘程序...")
    
    # 检查是否以管理员权限运行
    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
        if not is_admin:
            # 尝试创建一个批处理文件并以管理员权限运行
            tray_name = os.path.basename(sys.executable) if getattr(sys, "frozen", False) else "python.exe"
            script_content = f"""
            @echo off
            echo 正在关闭托盘程序...
            taskkill /im "{tray_name}" /f
            if %errorlevel% equ 0 (
                echo 成功关闭进程 "{tray_name}".
            ) else (
                echo 进程 "{tray_name}" 未运行或关闭失败.
            )
            exit
            """
            temp_script = os.path.join(os.environ.get('TEMP', '.'), 'stop_tray.bat')
            with open(temp_script, 'w') as f:
                f.write(script_content)
            
            # 以管理员权限运行批处理
            result = ctypes.windll.shell32.ShellExecuteW(
                None, "runas", "cmd.exe", f"/c {temp_script}", None, 0
            )
        # 退出当前实例
        threading.Timer(1.0, lambda: os._exit(0)).start()
    except Exception as e:
        logging.error(f"关闭托盘时出错: {e}")
        # 如果出错，仍然尝试正常退出
        threading.Timer(1.0, lambda: os._exit(0)).start()

def is_tray_admin():
    """检查当前托盘程序是否以管理员权限运行"""
    try:
        # 直接检查当前进程是否有管理员权限
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception as e:
        logging.error(f"检查托盘程序管理员权限时出错: {e}")
        return False

def check_tray_admin():
    """检查并通知托盘程序的管理员权限状态"""
    if is_tray_admin():
        notify("托盘程序已获得管理员权限")
    else:
        notify("托盘程序未获得管理员权限")

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

def on_check_tray_admin(icon, item):
    check_tray_admin()

def on_open_gui(icon, item):
    open_gui()

def on_stop_tray(icon, item):
    stop_tray()

def restart_tray_as_admin(icon, item):
    """以管理员权限重启托盘程序"""
    logging.info("以管理员权限重启托盘程序")
    # 提权重启当前脚本或exe
    try:
        if getattr(sys, "frozen", False):
            prog = sys.executable
            args = None
        else:
            prog = sys.executable
            args = os.path.abspath(__file__)
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", prog, f'"{args}"' if args else None, None, 0
        )
    except Exception as e:
        logging.error(f"重启托盘失败: {e}")
        messagebox.showerror("Error",f"重启托盘失败,请手动重启，详情请查看日志: {tray_log_path}")
        return
    # 退出当前实例
    threading.Timer(0.5, lambda: os._exit(0)).start()

# 检查主程序状态
main_status = "未运行"
if is_main_running():
    main_status = "以管理员权限运行" if is_main_admin() else "以普通权限运行"

# 检查托盘程序状态
tray_status = "以管理员权限运行" if is_tray_admin() else "以普通权限运行"

# 托盘图标设置
icon_path = resource_path(ICON_FILE)
image = Image.open(icon_path) if os.path.exists(icon_path) else None
menu = pystray.Menu(
        pystray.MenuItem("检查主程序管理员权限", on_check_admin),
        pystray.MenuItem("检查托盘程序管理员权限", on_check_tray_admin),
        pystray.MenuItem("打开配置界面", on_open_gui),
        pystray.MenuItem("打开主程序", on_open),
        pystray.MenuItem("以管理员权限重启主程序", on_restart_as_admin),
        pystray.MenuItem("以管理员权限重启托盘", restart_tray_as_admin),
        pystray.MenuItem("重启主程序", on_restart),
        pystray.MenuItem("关闭主程序", on_stop),
        pystray.MenuItem("退出托盘", on_stop_tray),
    )

# 启动前通知
notify(f"远程控制托盘程序已启动，主程序状态: {main_status}，托盘状态: {tray_status}")
logging.info(f"托盘程序启动，主程序状态: {main_status}，托盘状态: {tray_status}")

# 启动托盘图标
icon = pystray.Icon("Remote-Controls-Tray", image, "远程控制托盘", menu)
icon.run()