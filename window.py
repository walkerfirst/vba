""" 创建一个隐藏的窗口，用于创建对话框 """
import tkinter as tk

def create_window():
    # 创建确认对话框
    window = tk.Tk()
    window.overrideredirect(1)  # 完全隐藏窗口装饰
    window.withdraw()  # 立即隐藏主窗口
    
    # 提高DPI感知
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception as dpi_error:
        print(f"DPI设置失败: {str(dpi_error)}")
        
    # 设置窗口缩放
    window.tk.call('tk', 'scaling', 2.0)
    
    # 配置字体
    default_font = ('Microsoft YaHei', 10)  # 使用更清晰的字体
    window.option_add('*Font', default_font)
    
    # 确保窗口完全隐藏
    window.update_idletasks()
    window.update()
    return window
