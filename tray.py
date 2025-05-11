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
from logging.handlers import RotatingFileHandler
import traceback
from tkinter import N, messagebox
import pystray
from PIL import Image
import psutil
# try:
#     import utils  # 尝试导入工具模块
# except ImportError:
#     utils = None

# 线程装饰器
def run_in_thread(func):
    """
    装饰器：将被装饰的函数在单独的线程中执行
    确保菜单函数不会阻塞主UI线程

    参数：
    - func: 要在线程中执行的函数
    
    返回：
    - wrapper: 包装后的函数，它会在新线程中执行原函数
    """
    def wrapper(*args, **kwargs):
        logging.info(f"在单独线程中执行: {func.__name__}")
        thread = threading.Thread(target=func, args=args, kwargs=kwargs)
        thread.daemon = True
        thread.start()
    return wrapper

# 配置
MAIN_EXE = "Remote-Controls.exe" if getattr(sys, "frozen", False) else "main.py"
GUI_EXE = "RC-GUI.exe"
GUI_PY = "GUI.py"
ICON_FILE = "icon.ico"
MUTEX_NAME = "Remote-Controls-main"

# 日志配置
appdata_dir = os.path.abspath(os.path.dirname(sys.argv[0]))
logs_dir = os.path.join(appdata_dir, "logs")
# 确保日志目录存在
if not os.path.exists(logs_dir):
    try:
        os.makedirs(logs_dir)
    except Exception as e:
        print(f"创建日志目录失败: {e}")
        logs_dir = appdata_dir

tray_log_path = os.path.join(logs_dir, "tray.log")

# 配置日志处理器，启用日志轮转
log_handler = RotatingFileHandler(
    tray_log_path, 
    maxBytes=1*1024*1024,  # 1MB
    backupCount=1,          # 保留1个备份
    encoding='utf-8'
)

# 设置日志格式
log_formatter = logging.Formatter(
    '%(asctime)s [%(levelname)s] %(filename)s:%(lineno)d - %(message)s',
    datefmt="%Y-%m-%d %H:%M:%S"
)
log_handler.setFormatter(log_formatter)

# 获取根日志记录器并设置
logger = logging.getLogger()
logger.setLevel(logging.INFO)
logger.addHandler(log_handler)

# 记录程序启动信息
logging.info("="*50)
logging.info("远程控制托盘程序启动")
logging.info(f"程序路径: {os.path.abspath(__file__)}")
logging.info(f"工作目录: {os.getcwd()}")
logging.info(f"Python版本: {sys.version}")
logging.info(f"系统信息: {sys.platform}")
logging.info("="*50)

# # 记录详细的系统信息
# if utils:
#     utils.log_system_info(detailed=True)
#     # 启动定期状态记录（每小时记录一次系统状态）
#     status_thread = utils.log_periodic_status(interval=3600)
#     logging.info("已启用系统状态定期记录")
# else:
#     logging.warning("未能加载utils模块，系统信息记录功能不可用")

def resource_path(relative_path):
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base_path, relative_path)


def is_main_running():
    # 优先按进程检测主程序是否已运行
    main_proc = get_main_proc()
    if main_proc:
        return True
    # 回退到互斥体判断
    mutex = ctypes.windll.kernel32.OpenMutexW(0x100000, False, MUTEX_NAME)
    if mutex:
        ctypes.windll.kernel32.CloseHandle(mutex)
        logging.info(f"互斥体存在，主程序正在运行")
        return True
    else:
        logging.info(f"互斥体不存在，主程序未运行")
        return False

def run_as_admin(executable_path, parameters=None, working_dir=None, show_cmd=0):
    """
    以管理员权限运行指定程序

    参数：
    executable_path (str): 要执行的可执行文件路径
    parameters (str, optional): 传递给程序的参数，默认为None
    working_dir (str, optional): 工作目录，默认为None（当前目录）
    show_cmd (int, optional): 窗口显示方式，默认为0（1正常显示，0隐藏）

    返回：
    int: ShellExecute的返回值，若小于等于32表示出错
    """
    if parameters is None:
        parameters = ''
    if working_dir is None:
        working_dir = ''
    # 调用ShellExecuteW，设置动词为'runas'
    result = ctypes.windll.shell32.ShellExecuteW(
        None,                  # 父窗口句柄
        'runas',               # 操作：请求管理员权限
        executable_path,       # 要执行的文件路径
        parameters,            # 参数
        working_dir,           # 工作目录
        show_cmd               # 窗口显示方式
    )
    
    return result

def run_py_in_venv_as_admin_hidden(python_exe_path, script_path, script_args=None):
    """
    使用指定的 Python 解释器（如虚拟环境中的 python.exe）以管理员权限静默运行脚本
    
    参数：
    python_exe_path (str): Python 解释器路径（如 "D:/Code/Python/Remote-Controls/.venv/Scripts/python.exe"）
    script_path (str): 要运行的 Python 脚本路径
    script_args (list): 传递给脚本的参数（可选）
    """
    if not os.path.exists(python_exe_path):
        raise FileNotFoundError(f"Python 解释器未找到: {python_exe_path}")

    if script_args is None:
        script_args = []

    # 构造命令（确保路径带引号，防止空格问题）
    command = f'"{python_exe_path}" "{script_path}" {" ".join(script_args)}'
    logging.info(f"构造的命令: {command}")

    # 使用 ShellExecuteW 以管理员权限静默运行
    result = ctypes.windll.shell32.ShellExecuteW(
        None,               # 父窗口句柄
        'runas',            # 请求管理员权限
        'cmd.exe',          # 通过 cmd 执行（但隐藏窗口）
        f'/c {command}',    # /c 执行后关闭窗口
        None,               # 工作目录
        0                   # 窗口模式：0=隐藏
    )
    return result
def get_main_proc():
    """查找主程序进程是否存在"""
    process_name = MAIN_EXE
    
    # 如果不是管理员权限运行，可能无法查看所有进程，记录警告
    if not is_tray_admin():
        logging.warning("托盘程序未以管理员权限运行,可能无法查看所有进程")
    if MAIN_EXE.endswith('.exe'):
        logging.info(f"查找主程序可执行文件: {process_name}")
        # 可执行文件查找方式
        target_user=None
        process_name = process_name.lower()
        target_user = target_user.lower() if target_user else None

        for proc in psutil.process_iter(['name', 'username']):
            try:
                # 获取进程信息
                proc_info = proc.info
                current_name = proc_info['name'].lower()
                current_user = proc_info['username']

                # 匹配进程名
                if current_name == process_name:
                    # 未指定用户则直接返回True
                    if target_user is None:
                        return True
                    # 指定用户时提取用户名部分比较
                    if current_user:
                        user_part = current_user.split('\\')[-1].lower()
                        if user_part == target_user:
                            return True
            except (psutil.AccessDenied, psutil.NoSuchProcess):
                continue
        # 如果找不到进程，记录信息
        logging.info(f"未找到主程序进程: {process_name}")
        return None
    else:
        # Python脚本查找方式
        logging.info(f"查找Python脚本主程序: {process_name}")
        # 尝试查找命令行中包含脚本名的Python进程
        for proc in psutil.process_iter(['pid', 'name', 'cmdline', 'username']):
            try:
                proc_info = proc.info
                if proc_info['name'] and proc_info['name'].lower() in ('python.exe', 'pythonw.exe'):
                    cmdline = ' '.join(proc_info['cmdline']) if proc_info['cmdline'] else ""
                    if process_name in cmdline:
                        logging.info(f"找到主程序Python进程: {proc.pid}, 命令行: {cmdline}")
                        return proc
            except (psutil.AccessDenied, psutil.NoSuchProcess, Exception) as e:
                logging.error(f"获取Python进程信息失败: {e}")
                continue
        
        # 如果常规方法找不到，尝试使用wmic命令行工具
        logging.info("常规方法未找到Python进程，尝试使用wmic命令行工具")
        try:
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
                    
                    if current_cmd and current_pid and process_name in current_cmd:
                        try:
                            logging.info(f"wmic找到主程序Python进程: {current_pid}, 命令行: {current_cmd}")
                            return psutil.Process(int(current_pid))
                        except (psutil.NoSuchProcess, ValueError) as e:
                            logging.error(f"获取进程对象失败，PID: {current_pid}, 错误: {e}")
                    
                    current_pid = None
                    current_cmd = None
        except Exception as e:
            logging.error(f"使用wmic查找Python进程失败: {e}")
    
    logging.info("未找到主程序进程")
    return None
        

def is_main_admin():
    """检查主程序是否以管理员权限运行"""
    # 通过读取主程序写入的状态文件来判断管理员权限
    status_file = os.path.join(logs_dir, "admin_status.txt")
    
    # 首先检查文件是否存在
    if os.path.exists(status_file):
        try:
            with open(status_file, "r", encoding="utf-8") as f:
                content = f.read().strip()
                logging.info(f"读取主程序权限状态文件: {content}")
                # 检查文件内容
                if content == "admin=1":
                    return True
                elif content == "admin=0":
                    return False
                else:
                    logging.warning(f"权限状态文件内容格式错误: {content}")
        except Exception as e:
            logging.error(f"读取权限状态文件时出错: {e}")
    else:
        logging.warning(f"权限状态文件不存在: {status_file}")
        return False
    

def clean_orphaned_mutex():
    """清理可能未被正确释放的主程序互斥体"""
    logging.info(f"尝试清理互斥体: {MUTEX_NAME}")
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
        error_code = ctypes.windll.kernel32.GetLastError()
        logging.warning(f"普通权限清理互斥体失败，错误码: {error_code}，尝试提权清理")
        
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
        logging.info(f"创建临时清理脚本: {temp_script}")
        result = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", "cmd.exe", f"/c {temp_script}", None, 0
        )
        
        if result > 32:  # 成功启动
            logging.info("清理脚本启动成功，等待执行")
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
                error_code = ctypes.windll.kernel32.GetLastError()
                if error_code == 183:  # ERROR_ALREADY_EXISTS
                    logging.warning(f"提权清理后互斥体仍然存在，错误码: {error_code}")
                else:
                    logging.error(f"提权后检查互斥体状态失败，错误码: {error_code}")
                return False
        else:
            logging.error(f"提权清理互斥体失败，启动脚本失败，错误码: {result}")
            return False
    except Exception as e:
        logging.error(f"清理互斥体出现异常: {e}")
        # logging.error(f"异常详情: {traceback.format_exc()}")
        notify(f"清理互斥体失败，详情请查看日志: {tray_log_path}", level="error", show_error=True)
    return False

@run_in_thread
def start_main(icon=None, item=None):
    """启动主程序"""
    global main_process
    logging.info("="*30)
    logging.info("开始启动主程序")
    
    # 检查进程是否已在运行
    if get_main_proc():
        logging.info("主程序已在运行，不需要重复启动")
        notify("主程序已在运行", level="warning")
        return
    
    # 如果进程未运行但之前检测到互斥体存在，则可能是互斥体未被正确释放
    logging.info("主程序未运行，检查并清理可能残留的互斥体")
    clean_orphaned_mutex()
    
    # 选择启动方式
    cmd_line = None
    try:
        if MAIN_EXE.endswith('.exe') and os.path.exists(MAIN_EXE):
            logging.info(f"以可执行文件方式启动: {MAIN_EXE}")
            cmd_line = [os.path.abspath(MAIN_EXE)]
            main_process = subprocess.Popen(cmd_line, creationflags=subprocess.CREATE_NO_WINDOW)
            logging.info(f"启动进程ID: {main_process.pid}")
        elif os.path.exists(MAIN_EXE):
            logging.info(f"以Python脚本方式启动: {sys.executable} {MAIN_EXE}")
            cmd_line = [sys.executable, os.path.abspath(MAIN_EXE)]
            main_process = subprocess.Popen(cmd_line, creationflags=subprocess.CREATE_NO_WINDOW)
            logging.info(f"启动进程ID: {main_process.pid}")
        else:
            error_msg = f"未找到主程序文件: {MAIN_EXE}"
            logging.error(error_msg)
            notify(error_msg, level="error", show_error=True)
            messagebox.showerror("错误", f"未找到\"Remote-Controls\"程序\n路径: {os.path.abspath(MAIN_EXE)}")
            return
        
        # 检查进程是否成功启动
        time.sleep(1)
        if main_process.poll() is None:  # 如果进程仍在运行
            logging.info(f"主程序启动成功，进程ID: {main_process.pid}")
            # notify("主程序已启动", level="info")
        else:
            exit_code = main_process.returncode
            error_msg = f"主程序启动失败，退出码: {exit_code}"
            logging.error(error_msg)
            notify(error_msg, level="error", show_error=True)
    
    except Exception as e:
        cmd = " ".join(cmd_line) if cmd_line else "未知"
        logging.error(f"启动主程序时出现异常: {e}")
        logging.info(f"启动命令: {cmd}")
        # logging.error(f"异常详情: {traceback.format_exc()}")
        notify(f"启动主程序失败，详情请查看日志: {tray_log_path}", level="error", show_error=True)

@run_in_thread
def is_admin_start_main():
    """管理员权限运行主程序"""
    if MAIN_EXE.endswith('.exe') and os.path.exists(MAIN_EXE):
        close_exe(MAIN_EXE)
        logging.info(f"以可执行文件方式启动: {MAIN_EXE}")
        rest=run_as_admin(MAIN_EXE)
        if rest > 32:
            logging.info(f"成功以管理员权限启动主程序，PID: {rest}")
        else:
            logging.error(f"以管理员权限启动主程序失败，错误码: {rest}")
            notify(f"以管理员权限启动主程序失败，错误码: {rest}", level="error", show_error=True)
    elif os.path.exists(MAIN_EXE):
        close_script(MAIN_EXE)
        logging.info(f"以Python脚本方式启动: {sys.executable} {MAIN_EXE}")
        rest=run_py_in_venv_as_admin_hidden(sys.executable, MAIN_EXE)
        if rest > 32:
            logging.info(f"成功以管理员权限启动主程序，PID: {rest}")
        else:
            logging.error(f"以管理员权限启动主程序失败，错误码: {rest}")
            notify(f"以管理员权限启动主程序失败，错误码: {rest}", level="error", show_error=True)

@run_in_thread
def check_admin(icon=None, item=None):
    if is_main_running() and is_main_admin():
        logging.info("主程序已获得管理员权限")
        notify("主程序已获得管理员权限")
    elif is_main_running():
        logging.info("主程序未获得管理员权限")
        notify("主程序未获得管理员权限")
    else:
        logging.info("主程序未运行")
        notify("主程序未运行")

@run_in_thread
def open_gui(icon=None, item=None):
    if os.path.exists(GUI_EXE):
        subprocess.Popen([GUI_EXE])
        logging.info(f"打开配置界面: {GUI_EXE}")
    elif os.path.exists(GUI_PY):
        logging.info(f"打开配置界面: {GUI_PY}")
        subprocess.Popen([sys.executable, GUI_PY])
    else:
        logging.error(f"未找到配置界面: {GUI_EXE} 或 {GUI_PY}")
        notify("未找到配置界面")

@run_in_thread
def notify(msg, level="info", show_error=False):
    """
    发送通知并记录日志
    
    参数:
    - msg: 通知消息
    - level: 日志级别 ( "info", "warning", "error", "critical")
    - show_error: 是否在通知失败时显示错误对话框
    """
    # 根据级别记录日志
    log_func = getattr(logging, level.lower())
    log_func(f"通知: {msg}")
    
    # 发送Win11通知
    try:
        from win11toast import notify as toast
        threading.Thread(target=lambda: toast(msg)).start()
    except Exception as e:
        logging.error(f"发送通知失败: {e}")
        # logging.error(f"通知失败详情: {traceback.format_exc()}")
        
        if show_error:
            try:
                messagebox.showinfo("通知", msg)
            except Exception as e2:
                logging.error(f"显示消息框也失败: {e2}")
                print(msg)  # 最后的后备方案，打印到控制台
        else:
            print(msg)

def close_exe(name:str):
    """关闭指定名称的进程"""
    """关闭程序函数"""
    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
        if is_admin:
            # 尝试创建一个批处理文件并以管理员权限运行
            # tray_name = os.path.basename(sys.executable) if getattr(sys, "frozen", False) else "python.exe"
            logging.info(f"尝试关闭进程: {name}")
            script_content = f"""
            @echo off
            echo 正在关闭托盘程序...
            taskkill /im "{name}" /f
            if %errorlevel% equ 0 (
                echo 成功关闭进程 "{name}".
            ) else (
                echo 进程 "{name}" 未运行或关闭失败.
            )
            exit
            """
            temp_script = os.path.join(os.environ.get('TEMP', '.'), 'stop_tray.bat')
            with open(temp_script, 'w') as f:
                f.write(script_content)
            
            # 以管理员权限运行批处理
            rest=ctypes.windll.shell32.ShellExecuteW(
                None, "runas", "cmd.exe", f"/c {temp_script}", None, 0
            )
            if rest > 32:
                logging.info(f"成功关闭进程，PID: {rest}")
            else:
                logging.error(f"关闭进程失败，错误码: {rest}")
                notify(f"关闭进程失败，错误码: {rest}", level="error", show_error=True)
                return
        else:
            logging.warning(f"当前用户没有管理员权限，无法关闭进程{name}")
            notify(f"当前用户没有管理员权限，无法关闭进程{name}", level="warning")
    except FileNotFoundError:
        logging.error(f"未找到进程文件: {name}")
        notify(f"未找到进程文件: {name}", level="error", show_error=True)
    except Exception as e:
        logging.error(f"关闭{name}时出错: {e}")
        # 如果出错，仍然尝试正常退出
        threading.Timer(1.0, lambda: os._exit(0)).start()

def close_script(script_name):
    """关闭脚本"""
    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
        if is_admin:
            # 通过名称查找进程（模糊匹配）
            logging.info(f"尝试关闭脚本: {script_name}")
            cmd_find = f'tasklist /FI "IMAGENAME eq python.exe" /FO CSV /NH'
            output = subprocess.check_output(cmd_find, shell=True).decode('utf-8')
            
            # 解析输出，找到目标脚本的PID
            target_pids = []
            for line in output.splitlines():
                if script_name in line:
                    parts = line.replace('"', '').split(',')
                    pid = parts[1].strip()
                    target_pids.append(pid)
            
            # 终止所有匹配的进程
            for pid in target_pids:
                try:
                    # 以管理员权限调用 taskkill
                    subprocess.run(
                        f'taskkill /F /PID {pid}',
                        shell=True,
                        check=True,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    print(f"已终止进程 PID: {pid}")
                except subprocess.CalledProcessError:
                    print(f"无法终止进程 PID: {pid}（可能需要管理员权限）")
        else:
            logging.warning(f"当前用户没有管理员权限，无法关闭脚本{script_name}")
            notify(f"当前用户没有管理员权限，无法关闭脚本{script_name}", level="warning")
    except FileNotFoundError:
        logging.error(f"未找到脚本文件: {script_name}")
        notify(f"未找到脚本文件: {script_name}", level="error", show_error=True)
    except Exception as e:
        logging.error(f"关闭脚本{script_name}时出错: {e}")

@run_in_thread
def stop_tray():
    """关闭托盘程序和主程序"""
    logging.info("="*30)
    notify("正在关闭托盘程序...")
    
    # 安全停止托盘图标
    if 'icon' in globals() and icon:
        try:
            icon.stop()
            logging.info("托盘图标已停止")
        except Exception as e:
            logging.error(f"停止托盘图标时出错: {e}")
    tray_name = os.path.basename(sys.executable) if getattr(sys, "frozen", False) else "python.exe"
    if MAIN_EXE.endswith('.exe') and os.path.exists(MAIN_EXE):
        close_exe(MAIN_EXE)
        close_exe(tray_name)
    else:
        close_script(MAIN_EXE)
        close_script(tray_name)

        
def is_tray_admin():
    """检查当前托盘程序是否以管理员权限运行"""
    try:
        # 直接检查当前进程是否有管理员权限
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception as e:
        logging.error(f"检查托盘程序管理员权限时出错: {e}")
        return False

@run_in_thread
def check_tray_admin(icon=None, item=None):
    """检查并通知托盘程序的管理员权限状态"""
    if is_tray_admin():
        notify("托盘程序已获得管理员权限")
    else:
        notify("托盘程序未获得管理员权限")
        
def close_main():
    """关闭主程序"""
    main_proc = get_main_proc()
    if main_proc:
        if MAIN_EXE.endswith('.exe') and os.path.exists(MAIN_EXE):
            close_exe(MAIN_EXE)
        elif os.path.exists(MAIN_EXE):
            close_script(MAIN_EXE)
    else:
        logging.info("主程序未运行")

# 检查主程序状态
main_status = "未运行"
if is_main_running():
    main_status = "以管理员权限运行" if is_main_admin() else "以普通权限运行"

# 检查托盘程序状态
tray_status = "以管理员权限运行" if is_tray_admin() else "以普通权限运行"

# 权限提示
admin_tip = ""
if not is_tray_admin():
    admin_tip = "，无法查看开机自启的主程序状态"

# 托盘图标设置
icon_path = resource_path(ICON_FILE)
image = Image.open(icon_path) if os.path.exists(icon_path) else None
menu = pystray.Menu(
        pystray.MenuItem("检查主程序管理员权限", check_admin),
        pystray.MenuItem("检查托盘程序管理员权限", check_tray_admin),
        pystray.MenuItem("打开配置界面", open_gui),
        pystray.MenuItem("管理员权限启动主程序",is_admin_start_main),
        pystray.MenuItem("打开主程序", start_main),
        pystray.MenuItem("关闭主程序", close_main),
        pystray.MenuItem("退出托盘", stop_tray),
    )

# 信号处理函数，用于捕获CTRL+C等中断信号
def signal_handler(signum, frame):
    logging.info(f"接收到信号: {signum}，正在优雅退出")
    stop_tray()

# 注册信号处理器
try:
    import signal
    signal.signal(signal.SIGINT, signal_handler)  # 处理 Ctrl+C
    signal.signal(signal.SIGTERM, signal_handler)  # 处理终止信号
    logging.info("已注册信号处理器")
except (ImportError, AttributeError) as e:
    logging.warning(f"无法注册信号处理器: {e}")

# 启动前通知
notify(f"远程控制托盘程序已启动\n主程序状态: {main_status}\n托盘状态: {tray_status}{admin_tip}")
logging.info(f"托盘程序启动，主程序状态: {main_status}，托盘状态: {tray_status}{admin_tip}")

# 创建托盘图标
icon = pystray.Icon("Remote-Controls-Tray", image, "远程控制托盘", menu)

# 在带异常处理的环境中运行托盘程序
try:
    logging.info("开始运行托盘图标")
    icon.run()
except KeyboardInterrupt:
    logging.info("检测到键盘中断，正在退出")
    stop_tray()
except Exception as e:
    logging.error(f"托盘图标运行时出错: {e}")
    logging.error(traceback.format_exc())
finally:
    logging.info("托盘程序正在退出")
    os._exit(0)