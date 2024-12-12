import win32com.client
import pythoncom
import win32gui
import psutil
import keyboard
import tkinter as tk
from tkinter import ttk
import time

class CursorPositionApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("WPS Word 光标全局位置")
        
        # 设置窗口大小和位置
        self.root.geometry("300x150")
        self.root.resizable(False, False)
        
        # 创建标签框架
        frame = ttk.LabelFrame(self.root, text="光标全局位置信息", padding="10")
        frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 创建显示标签
        self.position_label = ttk.Label(frame, text="未获取位置信息")
        self.position_label.pack(pady=20)
        
        # 创建提示标签
        hint_label = ttk.Label(frame, text="按 Ctrl+Alt+1 获取光标全局位置")
        hint_label.pack(pady=5)
        
        # 注册快捷键
        keyboard.add_hotkey('ctrl+alt+1', self.show_cursor_position)
        
        # 窗口置顶
        self.root.attributes('-topmost', True)
        
    def get_wps_window(self):
        """获取 WPS 窗口句柄"""
        def callback(hwnd, hwnds):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if 'wps' in title.lower() and '.doc' in title.lower():
                    hwnds.append(hwnd)
            return True
        
        hwnds = []
        win32gui.EnumWindows(callback, hwnds)
        return hwnds[0] if hwnds else None

    def is_wps_running(self):
        """检查 WPS 是否在运行"""
        try:
            win32com.client.GetActiveObject('kwps.Application')
            return True
        except:
            return False

    def get_cursor_position(self):
        try:
            if not self.is_wps_running():
                return "WPS未运行"

            # 初始化 COM 环境
            pythoncom.CoInitialize()
            
            # 获取 WPS 窗口
            hwnd = self.get_wps_window()
            if not hwnd:
                return "未找到WPS窗口"

            # 确保 WPS 窗口在前台
            win32gui.SetForegroundWindow(hwnd)
            
            # 连接 WPS
            try:
                wps = win32com.client.GetActiveObject('kwps.Application')
            except:
                try:
                    wps = win32com.client.GetActiveObject('wps.Application')
                except:
                    return "无法连接到WPS"

            if not wps.Documents.Count:
                return "没有打开的文档"
                
            # 获取文档和选区
            doc = wps.ActiveDocument
            selection = wps.Selection
            
            # 获取位置信息
            start_pos = selection.Start
            text_before = doc.Range(0, start_pos).Text
            
            # 计算行号和列号
            line_number = text_before.count('\r') + 1
            last_newline = text_before.rfind('\r')
            if last_newline == -1:
                column_number = len(text_before) + 1
            else:
                column_number = len(text_before) - last_newline
            
            return f"行: {line_number}, 列: {column_number}"
            
        except Exception as e:
            return f"发生错误: {str(e)}"
        finally:
            pythoncom.CoUninitialize()

    def show_cursor_position(self):
        """显示光标位置"""
        position = self.get_cursor_position()
        self.position_label.config(text=position)
        
    def run(self):
        """运行应用程序"""
        self.root.mainloop()

def get_wps_cursor_position():
    try:
        # 检查 WPS 是否运行
        if not is_wps_running():
            print("WPS 未运行")
            return None, None

        # 初始化 COM 环境
        pythoncom.CoInitialize()
        
        # 获取 WPS 窗口
        hwnd = get_wps_window()
        if not hwnd:
            print("未找到 WPS 窗口")
            return None, None

        # 确保 WPS 窗口在前台
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.1)
        
        # 尝试连接 WPS
        try:
            wps = win32com.client.GetActiveObject('kwps.Application')
        except:
            try:
                wps = win32com.client.GetActiveObject('wps.Application')
            except:
                print("无法连接到 WPS")
                return None, None

        # 确保有活动文档
        if not wps.Documents.Count:
            print("没有打开的文档")
            return None, None
            
        # 获取当前活动文档和选区
        doc = wps.ActiveDocument
        selection = wps.Selection
        
        # 获取光标位置信息
        start_pos = selection.Start
        text_before = doc.Range(0, start_pos).Text
        
        # 计算行号（通过换行符计算）
        line_number = text_before.count('\r') + 1
        
        # 计算列号（从最后一个换行符到光标位置的字符数）
        last_newline = text_before.rfind('\r')
        if last_newline == -1:
            column_number = len(text_before) + 1
        else:
            column_number = len(text_before) - last_newline
        
        return line_number, column_number
        
    except Exception as e:
        print(f"发生错误: {str(e)}")
        return None, None
    finally:
        pythoncom.CoUninitialize()

def main():
    print("开始监控WPS光标位置...")
    print("按 Ctrl+C 退出程序")
    
    while True:
        try:
            position = get_wps_cursor_position()
            if position:
                line_num, col_num = position
                print(f"当前位置 - 行: {line_num}, 列: {col_num}")
            time.sleep(1)
        except KeyboardInterrupt:
            print("\n程序已退出")
            break
        except Exception as e:
            print(f"发生错误: {str(e)}")
            time.sleep(2)

if __name__ == "__main__":
    app = CursorPositionApp()
    app.run()