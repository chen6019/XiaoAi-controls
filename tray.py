# tray.py
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
TRAY_MUTEX_NAME = "Remote-Controls-tray"  # 托盘程序的互斥体名称

# 日志配置
appdata_dir = os.path.abspath(os.path.dirname(sys.argv[0]))
tray_log_path = os.path.join(appdata_dir, "tray.log")
logging.basicConfig(
    filename=tray_log_path,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s: %(message)s",
)

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
    """查找主程序进程，同时处理普通权限和管理员权限的情况"""
    # 方法1：使用psutil尝试直接获取进程信息
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
            
    # 方法2：使用wmic命令行工具查找进程，可以检测管理员权限运行的进程
    try:
        if MAIN_EXE.endswith('.exe'):
            process_name = MAIN_EXE
            cmd = f'wmic process where "name=\'{process_name}\'" get ProcessId /value'
            result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
            if "ProcessId=" in result.stdout:
                pid = int(result.stdout.strip().split("=")[1])
                # 尝试通过pid获取进程对象
                try:
                    return psutil.Process(pid)
                except psutil.NoSuchProcess:
                    pass
    except Exception:
        pass
        
    # 还可以检查是否存在互斥体，这可以作为进程存在的另一种判断方式
    if is_main_running():
        # 如果互斥体存在但找不到进程，创建一个简单的对象来模拟进程存在
        class DummyProcess:
            def __init__(self):
                self.pid = -1  # 给一个默认的pid值
                
            def terminate(self):
                pass
                
            def wait(self, timeout=None):
                pass
                
            def kill(self):
                pass
                
            def username(self):
                return "SYSTEM"  # 假设为管理员权限
                
        return DummyProcess()
            
    return None

def is_main_admin():
    """检查主程序是否以管理员权限运行"""
    proc = get_main_proc()
    if not proc:
        return False
    try:
        # 如果是DummyProcess，则已假定为管理员权限
        if hasattr(proc, 'username') and callable(proc.username):
            return proc.username().endswith("SYSTEM") or ctypes.windll.shell32.IsUserAnAdmin()
        return True  # 如果无法确定，假设为管理员权限
    except Exception:
        # 检查互斥体是否存在，如果存在，可能是管理员权限进程创建的
        return is_main_running()

main_process = None

def clean_orphaned_mutex():
    """清理可能未被正确释放的主程序互斥体"""
    try:
        # 尝试创建与主程序相同名称的互斥体
        mutex = ctypes.windll.kernel32.CreateMutexW(None, False, MUTEX_NAME)
        if mutex:
            # 释放并关闭互斥体
            ctypes.windll.kernel32.ReleaseMutex(mutex)
            ctypes.windll.kernel32.CloseHandle(mutex)
            return True
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
    """停止主程序，增强对管理员权限进程的处理能力"""
    logging.info("开始关闭主程序")
    proc = get_main_proc()
    if not proc:
        notify("主程序未运行")
        return
        
    # 检查是否是管理员权限进程
    is_admin_process = False
    try:
        if hasattr(proc, 'username') and callable(proc.username):
            from __main__ import DummyProcess
            is_admin_process = proc.username().endswith("SYSTEM") or isinstance(proc, DummyProcess)
    except:
        is_admin_process = is_main_admin()
    
    if is_admin_process:
        try:
            # 尝试常规 taskkill
            if MAIN_EXE.endswith('.exe'):
                result = subprocess.run(f"taskkill /F /IM {MAIN_EXE}", shell=True, capture_output=True)
                if result.returncode == 0:
                    notify("管理员权限主程序已关闭")
                    clean_orphaned_mutex()
                    return
                # 提升权限重试
                ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", "taskkill", f"/F /IM {MAIN_EXE}", None, 0
                )
                notify("以管理员权限尝试关闭主程序")
                clean_orphaned_mutex()
                return

            # 针对脚本模式，按 PID 提权终止
            if hasattr(proc, 'pid') and proc.pid > 0:
                ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", "taskkill", f"/F /PID {proc.pid}", None, 0
                )
                notify("以管理员权限尝试终止主程序进程")
                clean_orphaned_mutex()
                return

            notify("无法关闭管理员权限主程序，请手动结束")
        except Exception as e:
            logging.error(f"关闭管理员权限主程序失败: {e}")
            notify(f"关闭管理员权限主程序失败: {e}")
    else:
        # 对普通权限进程使用标准方法
        try:
            # 尝试通过terminate终止
            proc.terminate()
            try:
                proc.wait(5)
                notify("主程序已关闭")
            except Exception:
                # 如果wait失败，尝试使用命令行强制终止进程
                if MAIN_EXE.endswith('.exe'):
                    subprocess.run(f"taskkill /F /IM {MAIN_EXE}", shell=True)
                    notify("主程序已强制关闭")
                else:
                    # 对于Python脚本，需要找到对应的python进程
                    if hasattr(proc, 'pid'):
                        subprocess.run(f"taskkill /F /PID {proc.pid}", shell=True)
                        notify("主程序已强制关闭")
        except Exception as e:
            logging.error(f"关闭主程序失败: {str(e)}")
            notify(f"关闭主程序失败: {str(e)}")
            # 最后尝试使用taskkill命令
            if MAIN_EXE.endswith('.exe'):
                subprocess.run(f"taskkill /F /IM {MAIN_EXE}", shell=True)
                notify("尝试通过taskkill关闭主程序")
    
    # 无论结果如何，尝试清理互斥体
    time.sleep(1)
    clean_orphaned_mutex()

def force_stop_main():
    """强制关闭主程序并清理互斥体"""
    logging.info("开始强制关闭主程序")
    # 先尝试常规方式关闭
    proc = get_main_proc()
    if proc:
        try:
            proc.terminate()
            proc.wait(2)
        except Exception:
            # 如果常规关闭失败，使用强制结束
            try:
                proc.kill()
            except Exception:
                pass
    
    # 无论进程是否存在，都尝试清理可能残留的互斥体
    if clean_orphaned_mutex():
        notify("已强制清理主程序")
    else:
        logging.warning("没有检测到主程序或互斥体")
        notify("没有发现主程序或互斥体")
        return False

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
    try:
        mu = ctypes.windll.kernel32.OpenMutexW(0x100000, False, TRAY_MUTEX_NAME)
        if mu:
            ctypes.windll.kernel32.ReleaseMutex(mu)
            ctypes.windll.kernel32.CloseHandle(mu)
    except Exception:
        logging.error("释放托盘互斥体失败")
        notify("释放托盘互斥体失败")
        pass

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

def tray_main():
    # 创建托盘程序互斥体，检查是否已经运行
    tray_mutex = ctypes.windll.kernel32.CreateMutexW(None, False, TRAY_MUTEX_NAME)
    if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
        notify("托盘程序已在运行")
        sys.exit(0)
        
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

if __name__ == "__main__":
    tray_main()