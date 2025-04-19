""" 创建一个隐藏的窗口，用于创建对话框 """
import tkinter as tk
from tkinter import messagebox,Toplevel,ttk
import re

def window_askyesno(parent, title, message, keywords=None):
    """
    创建一个确认对话框，包含文本和按钮
    文本区域支持关键字高亮
    增加数字高亮（颜色区分）
    :param parent: 父窗口
    :param title: 对话框标题
    :param message: 显示的消息文本
    :param keywords: 关键字字典，格式为 {关键字: 颜色}
    :return: bool，True表示确认，False表示取消
    """
    result = [None]
    
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
    dialog.transient(parent)
    dialog.grab_set()
    
    # 主容器（增加底部padding）
    main_frame = ttk.Frame(dialog, padding=(0, 0, 0, 15))  # 底部15px间距
    main_frame.pack(fill='both', expand=True)
    
    # 文本区域（上方增加10px间距）
    text_frame = ttk.Frame(main_frame, padding=(0, 10, 0, 0))
    text_frame.pack(fill='both', expand=True)
    
    text = tk.Text(text_frame,
                  wrap='word',
                  font=("Microsoft YaHei", 11),
                  padx=12, pady=8,  # 增加文本内边距
                  relief='flat',
                  highlightthickness=0,
                  width=25, # 设置窗口宽度
                  height=1)
    text.pack(fill='both', expand=True)
    
    # 按钮区域（上方增加15px间距）
    btn_frame = ttk.Frame(main_frame, padding=(0, 15, 0, 0))  # 顶部15px间距
    btn_frame.pack(fill='x', side='bottom')
    
    # 按钮容器（右对齐+外边距）
    btn_container = ttk.Frame(btn_frame)
    btn_container.pack(side='right', padx=15)  # 右侧15px外边距

    ttk.Button(btn_container, text="取消", width=8,
              command=on_cancel).pack(side='left', padx=5)
    ttk.Button(btn_container, text="确定", width=8,
              command=on_confirm).pack(side='left', padx=5)
    
    # 插入内容
    text.insert('end', message)
    text.config(state='disabled')
    
    # 关键字高亮
    if keywords:
        for word, color in keywords.items():
            start = '1.0'
            while True:
                start = text.search(word, start, stopindex='end')
                if not start: break
                end = f"{start}+{len(word)}c"
                text.tag_add(word, start, end)
                text.tag_config(word, foreground=color, 
                               font=("Microsoft YaHei", 11, "bold"))
                start = end
    # 新增：数值高亮（红色粗体）
    def highlight_numbers():
        # 匹配整数/小数/百分数/货币等（可根据需要调整正则表达式）
        number_pattern = r"""
            \b\d+\.?\d*%?\b|       # 普通数字和百分数
            \$\d+\.?\d*\b|         # 美元金额
            \b\d+\.?\d*[元美元€£]   # 其他货币
        """
        for match in re.finditer(number_pattern, message, re.VERBOSE):
            start = f"1.0+{match.start()}c"
            end = f"1.0+{match.end()}c"
            text.tag_add("number", start, end)
        # 设置高亮样式
        text.tag_config("number", foreground="blue", font=("Microsoft YaHei", 10, "normal"))
    
    highlight_numbers()  # 执行数值高亮

    # 智能尺寸调整
    def adjust_size():
        dialog.update_idletasks()
        
        # 计算文本高度（行数*行高+边距）
        line_count = int(text.index('end-1c').split('.')[0])
        text_height = line_count * 20 + 30  # 每行20px+边距
        
        # 计算总高度（文本+按钮+间距）
        total_height = text_height + btn_frame.winfo_reqheight() + 25  # 额外25px缓冲
        
        # 设置窗口尺寸
        dialog.geometry(f"{min(500, text.winfo_reqwidth()+40)}x{total_height}")
        center_window()
    
    def center_window():
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() - dialog.winfo_width()) // 2
        y = (dialog.winfo_screenheight() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")
    
    # 窗口设置
    dialog.resizable(False, False)  # 固定大小更美观
    adjust_size()
    dialog.after(100, adjust_size)  # 二次调整
    
    parent.wait_window(dialog)
    return result[0]

"""
确认窗口对话框
"""
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
