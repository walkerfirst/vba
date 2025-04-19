import tkinter as tk
from tkinter import messagebox,Toplevel,ttk

"""
显示的对话框中字体不能设置颜色
"""
def window_askyesno(parent, title, message):
    """自定义清晰字体的确认对话框（屏幕居中）"""
    result = [None]  # 用于存储结果的闭包变量
    
    def on_confirm():
        nonlocal result
        result[0] = True
        dialog.destroy()
    
    def on_cancel():
        nonlocal result
        result[0] = False
        dialog.destroy()
    
    # 创建对话框窗口
    dialog = Toplevel(parent)
    dialog.title(title)
    dialog.transient(parent)  # 设为父窗口的临时窗口
    dialog.grab_set()         # 独占焦点
    
    # 设置字体样式
    style = ttk.Style()
    style.configure("Bold.TLabel", font=("Microsoft YaHei", 11, "bold"))
    style.configure("Normal.TLabel", font=("Microsoft YaHei", 11))
    style.configure("Large.TButton", font=("Microsoft YaHei", 10))
    
    # 添加内容
    ttk.Label(dialog, text=message, style="Normal.TLabel", 
             wraplength=400, justify="left").pack(pady=10, padx=15)
             
    
    # 按钮框
    btn_frame = ttk.Frame(dialog)
    btn_frame.pack(pady=(0, 10))
    
    ttk.Button(btn_frame, text="确定", style="Large.TButton", 
              command=on_confirm).pack(side=tk.RIGHT, padx=10)
    ttk.Button(btn_frame, text="取消", style="Large.TButton", 
              command=on_cancel).pack(side=tk.LEFT, padx=10)
    
    # 禁止窗口调整大小
    dialog.resizable(False, False)
    
    # 计算并设置居中位置（屏幕中央）
    def center_on_screen():
        dialog.update_idletasks()  # 确保窗口尺寸已计算
        
        # 获取屏幕尺寸
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        
        # 获取窗口尺寸
        window_width = dialog.winfo_reqwidth()
        window_height = dialog.winfo_reqheight()
        
        # 计算居中位置
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        # 设置窗口位置
        dialog.geometry(f"+{x}+{y}")
    
    # 首次居中显示
    center_on_screen()
    
    # 确保窗口显示后再次居中（防止某些系统下的位置偏移）
    dialog.after(100, center_on_screen)
    
    # 等待窗口关闭
    parent.wait_window(dialog)
    return result[0]