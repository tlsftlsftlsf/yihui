import win32gui
import win32con
import win32api
import time
import subprocess
import os
import sys
import datetime
import psutil
from win32com.client import Dispatch

# 全局时间备份
original_hour = None


def kill_process(process_name):
    """强制结束指定进程"""
    try:
        killed = False
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] == process_name:
                proc.kill()
                print(f"[进程终止] 已结束进程: {process_name}")
                killed = True
        return killed
    except Exception as e:
        print(f"[异常] 结束进程出错: {e}")
        return False


def backup_system_hour():
    """备份当前系统小时（仅小时部分）"""
    global original_hour
    original_hour = datetime.datetime.now().hour
    print(f"[时间备份] 原始小时已保存: {original_hour:02d}时")
    return original_hour


def set_system_hour(new_hour):
    """修改系统时间（仅小时部分）"""
    try:
        now = datetime.datetime.now()
        time_str = f"{new_hour:02d}:{now.minute:02d}:{now.second:02d}"
        result = os.system(f'time {time_str}')
        if result == 0:  # 返回0表示成功
            print(f"[时间修改] 已设置为: {time_str}")
            return True
        print("[错误] 时间修改失败（可能需要管理员权限）")
        return False
    except Exception as e:
        print(f"[异常] 修改时间出错: {e}")
        return False


def restore_system_hour():
    """恢复原始系统小时"""
    if original_hour is not None:
        if set_system_hour(original_hour):
            print(f"[时间恢复] 已成功还原到{original_hour:02d}时")
        else:
            print("[警告] 时间恢复失败，请手动检查系统时间！")


def open_app_shortcut(lnk_path, wait_time=5):
    """通过快捷方式启动应用"""
    try:
        if not os.path.exists(lnk_path):
            raise FileNotFoundError(f"快捷方式不存在: {lnk_path}")

        shell = Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(lnk_path)

        if not os.path.exists(shortcut.TargetPath):
            raise FileNotFoundError(f"目标程序不存在: {shortcut.TargetPath}")

        process = subprocess.Popen(
            shortcut.TargetPath,
            cwd=shortcut.WorkingDirectory,
            shell=True
        )
        print(f"[应用启动] 成功启动: {os.path.basename(shortcut.TargetPath)}")
        time.sleep(wait_time)
        return process
    except Exception as e:
        print(f"[错误] 启动失败: {e}")
        sys.exit(1)


def find_window(title, timeout=10):
    """查找窗口，支持超时等待"""
    start_time = time.time()
    while time.time() - start_time < timeout:
        hwnd = win32gui.FindWindow(None, title)
        if hwnd:
            print(f"[窗口查找] 找到窗口: {title} (句柄: {hwnd})")
            return hwnd
        time.sleep(1)
    print(f"[错误] 未找到窗口: {title} (等待超时)")
    return None


def click_button(parent_hwnd, button_text):
    """安全点击按钮"""
    try:
        button_hwnd = win32gui.FindWindowEx(parent_hwnd, 0, "Button", button_text)
        if button_hwnd:
            # 先确保窗口在前台
            win32gui.SetForegroundWindow(parent_hwnd)
            time.sleep(0.3)

            # 发送点击消息
            win32gui.SendMessage(button_hwnd, win32con.BM_CLICK, 0, 0)
            print(f"[按钮点击] 已触发: {button_text}")
            return True
        print(f"[错误] 找不到按钮: {button_text}")
        return False
    except Exception as e:
        print(f"[异常] 点击按钮出错: {e}")
        return False


def click_at_position(x, y):
    """在指定屏幕坐标位置点击鼠标"""
    try:
        # 移动鼠标到指定位置
        win32api.SetCursorPos((x, y))
        time.sleep(0.1)

        # 模拟鼠标左键按下和释放
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
        time.sleep(0.1)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)

        print(f"[坐标点击] 已点击位置: ({x}, {y})")
        return True
    except Exception as e:
        print(f"[异常] 点击坐标出错: {e}")
        return False


if __name__ == "__main__":
    # 配置区（请根据实际情况修改）
    CONFIG = {
        "shortcut_path": r"C:\Users\Administrator\Desktop\0\1.lnk",
        "window_title": "驿辉查件机器人v2.5.4",
        "button_text": "开始运行",
        "modified_hour": 22,  # 要设置的小时数
        "timeout": 15,  # 窗口等待超时(秒)
        "click_position": (1516, 38),  # 要点击的坐标
        "click_delay": 10,  # 点击延迟(秒)
        "process_to_kill": "YihuiRobotUI.exe"  # 需要关闭的进程名
    }

    try:
        # === 0. 关闭已有进程 ===
        print("=" * 40)
        print(f"正在检查并关闭 {CONFIG['process_to_kill']} 进程...")
        kill_process(CONFIG['process_to_kill'])
        time.sleep(1)  # 等待进程完全关闭

        # === 1. 启动应用 ===
        print("=" * 40)
        print("正在启动应用程序...")
        process = open_app_shortcut(CONFIG["shortcut_path"], wait_time=3)

        # === 2. 备份并修改时间 ===
        print("=" * 40)
        backup_system_hour()
        if not set_system_hour(CONFIG["modified_hour"]):
            sys.exit(1)
        time.sleep(1)  # 确保时间生效

        # === 3. 查找主窗口 ===
        print("=" * 40)
        main_hwnd = find_window(CONFIG["window_title"], timeout=CONFIG["timeout"])
        if not main_hwnd:
            sys.exit(1)

        # === 4. 点击按钮 ===
        print("=" * 40)
        if click_button(main_hwnd, CONFIG["button_text"]):
            print("[操作成功] 按钮点击完成，等待1秒后恢复时间...")
            time.sleep(1)  # 新增的1秒延迟
        else:
            print("[操作失败] 请检查按钮状态")

    except KeyboardInterrupt:
        print("\n[用户中断] 检测到Ctrl+C，正在恢复时间...")
    except Exception as e:
        print(f"[主程序异常] {e}")
    finally:
        # === 5. 确保时间恢复 ===
        print("=" * 40)
        restore_system_hour()
        print("=" * 40)
        print("程序执行完毕")

        # === 6. 延迟后点击指定坐标 ===
        print(f"等待 {CONFIG['click_delay']} 秒后点击坐标 {CONFIG['click_position']}")
        time.sleep(CONFIG["click_delay"])
        click_at_position(*CONFIG["click_position"])
