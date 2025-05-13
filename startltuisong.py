import os
import time
import psutil
from pywinauto import Application


def kill_process(process_name):
    """强制结束指定进程"""
    try:
        killed = False
        for proc in psutil.process_iter(['name']):
            if proc.info['name'].lower() == process_name.lower():
                proc.kill()
                print(f"[进程终止] 已结束进程: {process_name}")
                killed = True
        if not killed:
            print(f"[进程状态] 未找到运行中的 {process_name} 进程")
        return killed
    except Exception as e:
        print(f"[异常] 结束进程出错: {e}")
        return False


def open_app_shortcut(shortcut_path):
    """打开指定的快捷方式(.lnk)"""
    try:
        if not os.path.exists(shortcut_path):
            print(f"[文件错误] 快捷方式不存在: {shortcut_path}")
            return False

        os.startfile(shortcut_path)
        print(f"[程序启动] 已打开快捷方式: {os.path.basename(shortcut_path)}")
        return True
    except Exception as e:
        print(f"[异常] 打开快捷方式出错: {e}")
        return False


def press_key(key='enter'):
    """模拟按键"""
    try:
        import pyautogui
        pyautogui.press(key)
        print(f"[键盘输入] 已模拟按下 {key} 键")
        return True
    except ImportError:
        try:
            import keyboard
            keyboard.press_and_release(key)
            print(f"[键盘输入] 已模拟按下 {key} 键")
            return True
        except ImportError:
            print("[错误] 需要安装 pyautogui 或 keyboard 模块")
            return False
    except Exception as e:
        print(f"[异常] 模拟按键出错: {e}")
        return False


def bring_process_to_front(process_name, window_title):
    max_retries = 5
    retry_delay = 1
    for attempt in range(max_retries):
        try:
            for proc in psutil.process_iter(['name', 'pid']):
                if proc.info['name'].lower() == process_name.lower():
                    app = Application().connect(process=proc.info['pid'])
                    try:
                        window = app.window(title=window_title)
                        window.set_focus()
                        print(f"[窗口前置] {process_name} 已置于最前台")
                        return True
                    except Exception as win_e:
                        print(f"[尝试 {attempt + 1}] 查找窗口出错: {win_e}")
        except Exception as e:
            print(f"[尝试 {attempt + 1}] 将进程置于前台出错: {e}")
        time.sleep(retry_delay)
    print(f"[进程状态] 多次尝试后仍未找到运行中的 {process_name} 窗口")
    return False


if __name__ == "__main__":
    # 1. 关闭目标进程
    kill_process("Notify.exe")

    # 2. 打开快捷方式
    shortcut = r"C:\Users\Administrator\Desktop\0\n.lnk"
    open_app_shortcut(shortcut)

    time.sleep(10)
    # 3. 将Notify.exe进程置于最前台
    bring_process_to_front("Notify.exe", "n")

    # 4. 延迟5秒
    time.sleep(5)

    # 5. 回车
    press_key()
