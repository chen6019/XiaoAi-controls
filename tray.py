"""
程序文件名RC-tray.exe
运行用户：当前登录用户（可能有管理员权限）
pyinstaller -F -n RC-tray --windowed --icon=icon.ico --add-data "icon.ico;."  tray.py
"""


from email import message
from math import log
import os
import sys
import subprocess
import threading
import time
import ctypes
import logging
from logging.handlers import RotatingFileHandler
import traceback
from tkinter import messagebox
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
        thread.daemon = True  # 设置为守护线程，这样主程序退出时线程也会退出
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
        logs_dir = appdata_dir  # 如果创建失败，回退到应用目录

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
    if get_main_proc():
        return True
    # 回退到互斥体判断
    mutex = ctypes.windll.kernel32.OpenMutexW(0x100000, False, MUTEX_NAME)
    if mutex:
        ctypes.windll.kernel32.CloseHandle(mutex)
        return True
    return False

def run_as_admin(executable_path, parameters=None, working_dir=None, show_cmd=0):
    """
    以管理员权限运行指定程序

    参数：
    executable_path (str): 要执行的可执行文件路径
    parameters (str, optional): 传递给程序的参数，默认为None
    working_dir (str, optional): 工作目录，默认为None（当前目录）
    show_cmd (int, optional): 窗口显示方式，默认为1（1正常显示，0隐藏）

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
    # 方法1：使用psutil尝试直接获取进程信息
    for p in psutil.process_iter(['pid', 'name', 'exe', 'cmdline']):
        try:
            if MAIN_EXE.endswith('.exe'):
                if p.info['name'] == MAIN_EXE or (p.info['exe'] and os.path.basename(p.info['exe']) == MAIN_EXE):
                    logging.info(f"找到主程序进程: {p.info}")
                    return p
                else:
                    logging.info(f"未找到主程序进程")
                    return None
            else:
                if p.info['cmdline'] and MAIN_EXE in ' '.join(p.info['cmdline']):
                    logging.info(f"找到主程序进程: {p.info}")
                    return p
                else:
                    logging.info(f"未找到主程序进程")
                    return None
        except Exception as e:
            logging.error(f"获取进程信息失败: {e}")
            continue
            
    # 方法2：使用wmic命令行工具查找进程，可以检测管理员权限运行的进程
    try:
        if MAIN_EXE.endswith('.exe'):
            logging.info("主程序是可执行文件，尝试使用wmic查找")
            process_name = MAIN_EXE
            cmd = f'wmic process where "name=\'{process_name}\'" get ProcessId /value'
            result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
            if "ProcessId=" in result.stdout:
                pid_lines = [line for line in result.stdout.strip().split('\n') if "ProcessId=" in line]
                for pid_line in pid_lines:
                    try:
                        pid = int(pid_line.strip().split("=")[1])
                        # 尝试通过pid获取进程对象
                        logging.info(f"找到主程序进程: {pid}")
                        return psutil.Process(pid)
                    except (psutil.NoSuchProcess, ValueError):
                        logging.error(f"获取进程对象失败，PID: {pid_line}")
                        continue
        else:
            logging.info("主程序是Python脚本，尝试查找所有python进程")
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
                    logging.info(f"检查命令行: {line}")
                    current_cmd = line[12:]
                elif line.startswith("ProcessId="):
                    logging.info(f"检查进程ID: {line}")
                    current_pid = line[10:]
                    
                    if current_cmd and current_pid and MAIN_EXE in current_cmd:
                        try:
                            logging.info(f"找到主程序进程: {current_pid}, {current_cmd}")
                            return psutil.Process(int(current_pid))
                        except (psutil.NoSuchProcess, ValueError):
                            logging.error(f"获取进程对象失败，PID: {current_pid}")
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
def stop_main(icon=None, item=None):
    """停止主程序：使用taskkill命令关闭，类似于批处理文件的实现"""
    logging.info("="*30)
    logging.info("开始关闭主程序")
    
    # 获取主程序进程
    proc = get_main_proc()
    if not proc:
        logging.warning("主程序未运行，无需关闭")
        messagebox.showerror("Error","主程序未运行")
        return None

    # 检查是否为管理员权限进程
    is_admin_proc = is_main_admin()
    logging.info(f"主程序权限状态: {'管理员权限' if is_admin_proc else '普通权限'}")
    
    start_time = time.time()
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
                exit /b 1
            )
            exit /b 0
            """
            
            temp_script = os.path.join(os.environ.get('TEMP', '.'), 'stop_main.bat')
            logging.info(f"创建临时批处理脚本: {temp_script}")
            with open(temp_script, 'w') as f:
                f.write(script_content)
            
            # 以管理员权限运行批处理
            logging.info("尝试以管理员权限执行批处理脚本")
            result = ctypes.windll.shell32.ShellExecuteW(
                None, "runas", "cmd.exe", f"/c {temp_script}", None, 0
            )
            
            if result > 32:  # 成功启动
                logging.info(f"批处理脚本启动成功，等待执行完成...")
                time.sleep(3)  # 等待批处理执行完成
                if not is_main_running():
                    logging.info(f"主程序已成功关闭，耗时: {time.time()-start_time:.2f}秒")
                    notify("主程序已关闭", level="info")
                else:
                    # 如果仍在运行，尝试更直接的方法
                    logging.warning("提权关闭后主程序仍在运行")
                    notify("主程序仍在运行，请尝试手动关闭（任务管理器）", level="warning", show_error=True)
                    # force_stop_main()
            else:
                error_code = result
                logging.warning(f"提权失败，错误码: {error_code}，尝试使用常规方法关闭")
                notify("无法获取管理员权限，尝试使用常规方法关闭", level="warning")
                
                # 回退到常规方法
                logging.info("使用常规方法关闭进程")
                subprocess.run(f'taskkill /im "{MAIN_EXE}" /f', shell=True)
                
                # 检查是否成功关闭
                time.sleep(1)
                if not is_main_running():
                    logging.info(f"主程序已成功关闭(常规方法)，耗时: {time.time()-start_time:.2f}秒")
                    notify("主程序已关闭", level="info")
                else:
                    logging.error("常规方法关闭失败，主程序仍在运行")
                    notify("无法关闭主程序，请尝试手动关闭（任务管理器）", level="error", show_error=True)
        else:
            # 正常关闭方法
            logging.info("使用常规方法关闭普通权限进程")
            result = subprocess.run(f'taskkill /im "{MAIN_EXE}" /f', shell=True, capture_output=True, text=True)
            
            # 检查是否成功关闭
            if result.returncode == 0:
                logging.info(f"主程序已成功关闭，耗时: {time.time()-start_time:.2f}秒")
                notify("主程序已关闭", level="info")
            else:
                logging.error(f"关闭主程序失败: {result.stderr}")
                notify("关闭主程序失败，详情请查看日志", level="error", show_error=True)
    except Exception as e:
        logging.error(f"关闭主程序时出现异常: {e}")
        # logging.error(f"异常详情: {traceback.format_exc()}")
        notify(f"关闭主程序失败，详情请查看日志: {tray_log_path}", level="error", show_error=True)
    
    # 不管成功与否，都尝试清理互斥体
    logging.info("尝试清理可能残留的互斥体")
    time.sleep(1)
    clean_orphaned_mutex()

@run_in_thread
def restart_main(icon=None, item=None):
    """重启主程序"""
    logging.info("="*30)
    logging.info("重启主程序")
    
    # 检查主程序是否正在运行
    was_running = False
    proc = get_main_proc()
    if proc:
        was_running = True
        logging.info(f"主程序当前在运行，PID: {proc.pid}，准备关闭")
    else:
        logging.info("主程序当前未运行，将直接启动")
    
    # 如果正在运行，先停止
    if was_running:
        stop_main()
        time.sleep(2)  # 确保有足够时间关闭
        
        # 二次确认进程已关闭
        if is_main_running():
            logging.warning("主程序未能完全关闭，重启可能不完全")
            notify("主程序未能完全关闭，重启可能不完全", level="warning", show_error=True)
    
    # 启动主程序
    logging.info("开始启动主程序")
    start_main()
    
    # 确认重启结果
    time.sleep(1)
    if is_main_running():
        proc = get_main_proc()
        if proc:
            logging.info(f"主程序重启成功，新PID: {proc.pid}")
            notify("主程序已成功重启", level="info")
        else:
            # 理论上不应该运行到这里
            logging.warning("检测到主程序正在运行，但无法获取进程详情")
    else:
        logging.error("主程序重启失败，未检测到进程")
        notify("主程序重启失败，详情请查看日志", level="error", show_error=True)

@run_in_thread
def restart_main_as_admin(icon=None, item=None):
    """以管理员权限重启主程序"""
    logging.info("="*30)
    logging.info("以管理员权限重启主程序")
    
    # 记录当前状态
    was_running = False
    proc = get_main_proc()
    if proc:
        was_running = True
        logging.info(f"主程序当前在运行，PID: {proc.pid}，准备关闭")
        stop_main()
        time.sleep(2)  # 确保有足够时间关闭
        
        # 二次确认进程已关闭
        if is_main_running():
            logging.warning("主程序仍在运行，可能无法完全关闭")
            notify("无法完全关闭主程序，重启可能不完全", level="warning", show_error=True)
            return None
    else:
        logging.info("主程序当前未运行")
    
    logging.info("清理可能残留的互斥体")
    clean_orphaned_mutex()
    
    # 确定要启动的程序
    args = None
    try:
        if MAIN_EXE.endswith('.exe') and os.path.exists(MAIN_EXE):
            program = os.path.abspath(MAIN_EXE)
            logging.info(f"准备以管理员权限启动可执行文件: {program}")
        elif os.path.exists(MAIN_EXE):
            program = sys.executable
            args = os.path.abspath(MAIN_EXE)
            logging.info(f"准备以管理员权限启动Python脚本: {program} {args}")
        else:
            error_msg = f"未找到主程序文件: {MAIN_EXE}"
            logging.error(error_msg)
            messagebox.showerror("Error", f"未找到\"Remote-Controls\"程序\n路径: {os.path.abspath(MAIN_EXE)}")
            return
        
        # 使用 ShellExecute 以管理员权限启动程序
        logging.info("尝试获取管理员权限启动程序")
        logging.info(f"ShellExecuteW参数: {program}, {args}")
        if MAIN_EXE.endswith('.exe'):
            logging.info("以可执行文件方式启动")
            result = run_as_admin(program)
        else:
            logging.info("以Python脚本方式启动")
            result = run_py_in_venv_as_admin_hidden(program, args)
        # 处理错误（返回值 <= 32 表示错误）
        if result <= 32:
            error_messages = {
                0: "系统内存或资源不足",
                2: "文件未找到",
                3: "路径未找到",
                5: "拒绝访问",
                8: "内存不足",
                11: "错误的格式",
                26: "共享冲突",
                27: "关联不完整",
                30: "DDE繁忙",
                31: "DDE失败",
                32: "DDE超时或DLL未找到",
                1223: "用户取消操作"
            }
            error_msg = error_messages.get(result, f"未知错误，代码：{result}")
            logging.error(f"以管理员权限启动失败，错误码: {result}，{error_msg}")
            raise RuntimeError(f"无法以管理员权限启动程序：{error_msg}")
        elif result > 32:  # 成功启动
            logging.info("成功以管理员权限启动主程序")
            notify("已以管理员权限启动主程序", level="info")
            
            # 等待一段时间，检查程序是否真的启动了
            time.sleep(3)
            if is_main_running():
                proc = get_main_proc()
                if proc:
                    logging.info(f"确认主程序已启动，PID: {proc.pid}")
                    # 检查是否真的以管理员身份运行
                    if is_main_admin():
                        logging.info("确认主程序已获得管理员权限")
                    else:
                        logging.warning("主程序已启动，但似乎未获得管理员权限")
                        notify("主程序已启动，但可能未获得管理员权限", level="warning")
            else:
                logging.warning("主程序可能启动失败，未检测到进程")
                notify("主程序可能启动失败，请检查日志", level="warning")
        else:
            logging.error(f"以管理员权限启动失败，错误码: {result}")
            notify("以管理员权限启动失败，可能被用户取消", level="error", show_error=True)
    except Exception as e:
        logging.error(f"重启主程序时出现异常: {e}")
        # logging.error(f"异常详情: {traceback.format_exc()}")
        notify(f"重启主程序失败，详情请查看日志: {tray_log_path}", level="error", show_error=True)

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

@run_in_thread
def check_tray_admin(icon=None, item=None):
    """检查并通知托盘程序的管理员权限状态"""
    if is_tray_admin():
        notify("托盘程序已获得管理员权限")
    else:
        notify("托盘程序未获得管理员权限")

@run_in_thread
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
        pystray.MenuItem("检查主程序管理员权限", check_admin),
        pystray.MenuItem("检查托盘程序管理员权限", check_tray_admin),
        pystray.MenuItem("打开配置界面", open_gui),
        pystray.MenuItem("打开主程序", start_main),
        pystray.MenuItem("以管理员权限重启主程序", restart_main_as_admin),
        pystray.MenuItem("以管理员权限重启托盘", restart_tray_as_admin),
        pystray.MenuItem("重启主程序", restart_main),
        pystray.MenuItem("关闭主程序", stop_main),
        pystray.MenuItem("退出托盘", stop_tray),
    )

# 启动前通知
notify(f"远程控制托盘程序已启动，主程序状态: {main_status}，托盘状态: {tray_status}")
logging.info(f"托盘程序启动，主程序状态: {main_status}，托盘状态: {tray_status}")

# 启动托盘图标
icon = pystray.Icon("Remote-Controls-Tray", image, "远程控制托盘", menu)
icon.run()