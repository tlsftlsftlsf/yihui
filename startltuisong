import os
import time
import subprocess
import psutil  # 需要安装：pip install psutil

def kill_process(process_name):
    """关闭指定名称的进程"""
    try:
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'] == process_name:
                proc.kill()
                print(f"成功关闭进程: {process_name}")
                return True
        print(f"未找到进程: {process_name}")
        return False
    except Exception as e:
        print(f"关闭进程 {process_name} 时出错: {e}")
        return False

def open_app_shortcut(shortcut_path):
    """打开指定的快捷方式(.lnk)"""
    try:
        if not os.path.exists(shortcut_path):
            print(f"快捷方式不存在: {shortcut_path}")
            return False
            
        # 使用系统默认方式打开快捷方式
        os.startfile(shortcut_path)
        print(f"成功打开快捷方式: {shortcut_path}")
        return True
    except Exception as e:
        print(f"打开快捷方式 {shortcut_path} 时出错: {e}")
        return False

def press_enter():
    """模拟按下回车键"""
    try:
        import keyboard  # 需要安装：pip install keyboard
        keyboard.press_and_release('enter')
        print("已模拟按下回车键")
        return True
    except ImportError:
        print("未安装keyboard模块，无法模拟按键")
        return False
    except Exception as e:
        print(f"模拟按键失败: {e}")
        return False

# 主程序
if __name__ == "__main__":
    # 1. 关闭进程Notify.exe
    kill_process("Notify.exe")
    
    # 2. 打开快捷方式
    shortcut_path = r"C:\Users\Administrator\Desktop\1\n.lnk"
    open_app_shortcut(shortcut_path)
    
    # 3. 延迟5秒后输入回车键
    time.sleep(5)
    press_enter()
