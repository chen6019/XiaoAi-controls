"""
程序文件名RC-tray.exe
运行用户：当前登录用户（可能有管理员权限）
pyinstaller -F -n RC-tray --windowed --icon=icon.ico --add-data "icon.ico;."  tray.py
"""


import os
import sys
import subprocess
import threading
import time
import ctypes
import logging
import pystray
from PIL import Image
import psutil

# 配置
MAIN_EXE = "Remote-Controls.exe" if getattr(sys, "frozen", False) else "main.py"
GUI_EXE = "RC-GUI.exe"
GUI_PY = "GUI.py"
ICON_FILE = "icon.ico"
MUTEX_NAME = "Remote-Controls-main"
# TRAY_MUTEX_NAME = "Remote-Controls-tray"
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
        notify(f"清理互斥体失败: {e}")
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
        notify("未找到主程序")
        return
    notify("主程序已启动")

def stop_main():
    """停止主程序：智能选择终止方法"""
    logging.info("开始关闭主程序")
    proc = get_main_proc()
    if not proc:
        notify("主程序未运行")
        return

    # 检查是否为管理员权限进程
    is_admin_proc = False
    try:
        if hasattr(proc, 'username') and callable(proc.username):
            username = proc.username()
            if username.endswith("SYSTEM") or "Administrator" in username:
                is_admin_proc = True
                logging.info(f"检测到主程序以管理员权限运行 ({username})")
    except Exception:
        # 如果无法获取用户名，通过其他方式判断
        is_admin_proc = is_main_admin()
        
    if is_admin_proc:
        logging.info("主程序以管理员权限运行，尝试提权关闭...")
        try:
            # 尝试提权后关闭
            script_content = f"""
            taskkill /F /IM {MAIN_EXE if MAIN_EXE.endswith('.exe') else 'python.exe'}
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
                time.sleep(2)  # 等待批处理执行完成
                if not is_main_running():
                    notify("主程序已关闭")
                else:
                    # 如果仍在运行，尝试更直接的方法
                    force_stop_main()
            else:
                logging.warning(f"提权失败，尝试使用常规方法关闭 (错误码: {result})")
                notify("无法获取管理员权限，尝试使用常规方法关闭")
                
                # 回退到常规方法
                if MAIN_EXE.endswith('.exe'):
                    subprocess.run(f"taskkill /F /IM {MAIN_EXE}", shell=True)
                elif hasattr(proc, 'pid'):
                    subprocess.run(f"taskkill /F /PID {proc.pid}", shell=True)
        except Exception as e:
            logging.error(f"提权关闭主程序失败: {e}")
            notify(f"提权关闭失败，尝试常规方法")
            # 回退到常规方法
            try:
                if MAIN_EXE.endswith('.exe'):
                    subprocess.run(f"taskkill /F /IM {MAIN_EXE}", shell=True)
                elif hasattr(proc, 'pid'):
                    subprocess.run(f"taskkill /F /PID {proc.pid}", shell=True)
            except Exception as e2:
                logging.error(f"常规方法也失败: {e2}")
                notify("所有关闭方法都失败，请手动关闭主程序")
    else:
        # 正常关闭方法
        try:
            if MAIN_EXE.endswith('.exe'):
                subprocess.run(f"taskkill /F /IM {MAIN_EXE}", shell=True)
            elif hasattr(proc, 'pid'):
                subprocess.run(f"taskkill /F /PID {proc.pid}", shell=True)
            notify("主程序已关闭")
        except Exception as e:
            logging.error(f"关闭主程序失败: {e}")
            notify(f"关闭主程序失败: {e}")
    
    # 不管成功与否，都尝试清理互斥体
    time.sleep(1)
    clean_orphaned_mutex()

def force_stop_main():
    """强制关闭主程序并清理互斥体"""
    logging.info("开始强制关闭主程序")
    # 先尝试常规方式关闭
    proc = get_main_proc()
    if proc:
        try:
            # 检查是否为管理员权限进程
            is_admin_proc = False
            try:
                if hasattr(proc, 'username') and callable(proc.username):
                    username = proc.username()
                    if username.endswith("SYSTEM") or "Administrator" in username:
                        is_admin_proc = True
                        logging.info(f"检测到主程序以管理员权限运行 ({username})")
            except Exception:
                # 如果无法获取用户名，通过其他方式判断
                is_admin_proc = is_main_admin()
            
            if is_admin_proc:
                logging.info("主程序以管理员权限运行，尝试提权强制关闭...")
                # 创建更强力的批处理脚本
                script_content = f"""
                taskkill /F /IM {MAIN_EXE if MAIN_EXE.endswith('.exe') else 'python.exe'} /T
                timeout /t 1
                taskkill /F /IM {MAIN_EXE if MAIN_EXE.endswith('.exe') else 'python.exe'} /T
                exit
                """
                temp_script = os.path.join(os.environ.get('TEMP', '.'), 'force_stop_main.bat')
                with open(temp_script, 'w') as f:
                    f.write(script_content)
                
                # 以管理员权限运行批处理
                result = ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", "cmd.exe", f"/c {temp_script}", None, 0
                )
                
                if result > 32:  # 成功启动
                    time.sleep(3)  # 等待批处理执行完成
                    # 无论成功与否，都继续尝试常规方法作为备份
            
            # 尝试常规方法终止
            try:
                proc.terminate()
                proc.wait(2)
            except Exception:
                # 如果常规关闭失败，使用强制结束
                try:
                    proc.kill()
                except Exception:
                    pass
        except Exception as e:
            logging.error(f"强制关闭进程失败: {e}")
            # 尝试使用命令行方式终止
            try:
                if MAIN_EXE.endswith('.exe'):
                    subprocess.run(f"taskkill /F /IM {MAIN_EXE} /T", shell=True)
                elif hasattr(proc, 'pid'):
                    subprocess.run(f"taskkill /F /PID {proc.pid} /T", shell=True)
            except Exception:
                pass
    
    # 无论进程是否存在，都尝试清理可能残留的互斥体
    if clean_orphaned_mutex():
        notify("已强制清理主程序")
    else:
        logging.warning("没有检测到主程序或互斥体")
        notify("没有发现主程序或互斥体")
        return False
    
    # 最后检查一次是否完全终止
    if is_main_running():
        notify("无法完全终止主程序，请尝试手动关闭或重启电脑")
        return False
    return True

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
        notify("未找到主程序")
        return
    
    try:
        # 使用 ShellExecute 以管理员权限启动程序
        # SW_HIDE = 0 用于隐藏窗口
        if MAIN_EXE.endswith('.exe'):
            ctypes.windll.shell32.ShellExecuteW(None, "runas", program, None, None, 0)
        else:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", program, f'"{args}"', None, 0)
        notify("已以管理员权限启动主程序")
    except Exception as e:
        logging.error(f"启动失败: {e}")
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
    logging.info(f"通知: {msg}")
    try:
        from win11toast import notify as toast
        threading.Thread(target=lambda: toast(msg)).start()
    except Exception as e:
        logging.error(f"通知失败: {e}")
        print(msg)

def on_open(icon, item):
    start_main()

def on_restart(icon, item):
    restart_main()

def on_restart_as_admin(icon, item):
    restart_main_as_admin()

def on_stop(icon, item):
    stop_main()

def on_force_stop(icon, item):
    force_stop_main()

def on_check_admin(icon, item):
    check_admin()

def on_open_gui(icon, item):
    open_gui()

def on_exit(icon, item):
    icon.stop()
    # 不直接调用 sys.exit()，因为这会在 pystray 的事件处理中引发异常
    # 而是安排在事件处理完成后退出
    threading.Timer(0.5, lambda: os._exit(0)).start()

def restart_tray_as_admin(icon, item):
    """以管理员权限重启托盘程序"""
    logging.info("以管理员权限重启托盘程序")
    # 先释放已创建的托盘互斥体
    # try:
    #     mu = ctypes.windll.kernel32.OpenMutexW(0x100000, False, TRAY_MUTEX_NAME)
    #     if mu:
    #         ctypes.windll.kernel32.ReleaseMutex(mu)
    #         ctypes.windll.kernel32.CloseHandle(mu)
    # except Exception:
    #     logging.error("释放托盘互斥体失败")
    #     notify("释放托盘互斥体失败")
    #     pass

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
        notify(f"重启托盘失败: {e}")
        return

    # 退出当前实例
    threading.Timer(0.5, lambda: os._exit(0)).start()

# def tray_main():
    # 创建托盘程序互斥体，检查是否已经运行
# tray_mutex = ctypes.windll.kernel32.CreateMutexW(None, False, TRAY_MUTEX_NAME)
# if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
#     notify("托盘程序已在运行")
#     sys.exit(0)
        # 简单检查桌面环境，但不要求必须可用
    # logging.info("尝试检测桌面环境...")
    # if is_desktop_available():
    #     logging.info("桌面环境检测成功，继续启动托盘程序")
    # else:
    #     logging.warning("桌面环境尚未完全可用，但仍将继续启动托盘程序")
    #     # 给桌面环境留出一些加载时间
    #     time.sleep(2)
    
    # 检查主程序状态
main_status = "未运行"
if is_main_running():
    main_status = "以管理员权限运行" if is_main_admin() else "以普通权限运行"
    
notify(f"远程控制托盘程序已启动，主程序状态: {main_status}")
logging.info(f"托盘程序启动，主程序状态: {main_status}")
        
icon_path = resource_path(ICON_FILE)
image = Image.open(icon_path) if os.path.exists(icon_path) else None
menu = pystray.Menu(
        pystray.MenuItem("检查主程序管理员权限", on_check_admin),
        pystray.MenuItem("打开配置界面", on_open_gui),
        pystray.MenuItem("打开主程序", on_open),
        pystray.MenuItem("重启主程序", on_restart),
        pystray.MenuItem("以管理员权限重启主程序", on_restart_as_admin),
        pystray.MenuItem("关闭主程序", on_stop),
        pystray.MenuItem("强制关闭主程序", on_force_stop),
        pystray.MenuItem("以管理员权限重启托盘", restart_tray_as_admin),
        pystray.MenuItem("退出托盘", on_exit),
    )
icon = pystray.Icon("Remote-Controls-Tray", image, "远程控制托盘", menu)
icon.run()

# if __name__ == "__main__":
#     tray_main()