# pip install pywin32
import win32com.client
# 注意这里选择从
from config import shipment_file
from window import create_window
from db import read_db,conn
from datetime import datetime
from tkinter import messagebox
import os
from vba_replacement import EXCELProcessor
from tkinter import Tk, StringVar, OptionMenu, Button, ttk

 # 定义 workbooks 对象
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False  # 可以设置为 True 调试
excel.DisplayAlerts = False # 禁用警告
def run():
    if not os.path.exists(shipment_file):
        msg_window = create_window()
        messagebox.showerror("错误", f"文件 {shipment_file} 不存在",parent=msg_window)
        return
    wb = excel.Workbooks.Open(shipment_file)
    
    # 获取发货信息
    sql = 'select * from shipView where id=1'
    ship_dict = read_db(sql, conn)[0]
    if ship_dict['model'] == '发货':
        # 指定执行VBA文件和 function
        process_data(data=ship_dict, wb=wb)

def process_data(data, wb):

    try:
        # 定义工作表
        sheet = wb.Sheets('data')
        sm = wb.Sheets('情况说明')
        sheet1 = wb.Sheets('sheet1')

        # 修改sheet1单元格数据，让 产品种类\手动输入 从下拉菜单选择
        def set_value():
            sheet1.Cells(9,9).Value = var.get() # 产品种类
            wb.Sheets('标签').Cells(6, 8).Value = var2.get() # 手动输入开关
            root.destroy()
        
        root = Tk()
        root.title("初始设置")
        root.geometry("400x260")  # 设置窗口大小
        root.resizable(True, True)  # 允许调整窗口大小
        
        # 主框架
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # 配置网格权重
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        var = StringVar(root)
        var.set("1")  # 默认值
        options = ["1", "2", "3"]  # 下来菜单选项
        
        # 窗口标题
        style = ttk.Style()
        style.configure('Title.TLabel', 
                       font=('Arial', 15, 'bold'),
                       foreground='#004080',  # 深蓝色
                       background='#f0f0f0',  # 浅灰色背景
                       padding=10)
        label = ttk.Label(main_frame, 
                         text="发货系统设置", 
                         style='Title.TLabel')
        # 使用ttk样式
        style.configure('TButton', font=('Arial', 12), padding=5)
        style.configure('TMenubutton', font=('Arial', 12), padding=5, background='#0078d7', foreground='white')
        style.map('TMenubutton',
            background=[('active', '#106ebe')],
            foreground=[('active', 'white')]
        )
        style.configure('TLabel', font=('Arial', 12), padding=5)
        label.pack(pady=5)
        
        # 创建水平布局容器
        dropdown_frame = ttk.Frame(main_frame)
        dropdown_frame.pack(pady=20)
        
        # 包裹种类下拉菜单
        ttk.Label(dropdown_frame, text="包裹种类").pack(side='left', padx=5)
        option_menu = ttk.OptionMenu(dropdown_frame, var, options[0], *options)
        option_menu.pack(side='left', padx=10)
        
        # YES/NO 下拉菜单
        ttk.Label(dropdown_frame, text="手动重量").pack(side='left', padx=5)
        var2 = StringVar(root)
        yesno_options = ["NO","YES" ]
        yesno_menu = ttk.OptionMenu(dropdown_frame, var2,yesno_options[0], *yesno_options)
        yesno_menu.pack(side='left', padx=10)
        
        # 确定按钮
        confirm_btn = ttk.Button(main_frame, text="确定", command=set_value)
        confirm_btn.pack(pady=15)
        root.mainloop()

        # 数据处理
        if data['tax'] == 1.0:
            tax = '要退税'
            data['trade'] = "一般贸易"
            company = "上海盛傲化学有限公司"
        else:
            tax = '不退税'
            company = wb.Sheets('报关公司').Cells(4,1).Value

        if data['express'] == '空运':
            express = 'by air'
        elif data['express'] == '海运':
            express = 'by sea'
        else:
            express = data['express']
        
        sm.Cells(1,21).Value = 1  #情况说明选择值设为1
        sheet1.Cells(14,3).Value = data['chinese'] # 主产品名称

        # data 页面单元格赋值
        sheet.Cells(2, 1).Value = data['chinese']
        sheet.Cells(2, 2).Value = data['name']
        sheet.Cells(2, 3).Value = company
        sheet.Cells(2, 4).Value = tax
        sheet.Cells(2, 5).Value = express
        sheet.Cells(2, 6).Value = data['waybill']
        sheet.Cells(2, 7).Value = data['pcs']
        sheet.Cells(2, 8).Value = data['package']
        sheet.Cells(2, 9).Value = data['invoice']
        sheet.Cells(2, 10).Value = data['ask']
        sheet.Cells(2, 11).Value = data['nw_unit2']
        sheet.Cells(2, 12).Value = data['qty']
        sheet.Cells(2, 13).Value = data['nw_unit']
        sheet.Cells(2, 14).Value = data['gross']
        sheet.Cells(2, 15).Value = data['trade']
        sheet.Cells(2, 16).Value = data['place']
        sheet.Cells(2, 17).Value = data['date']
        sheet.Cells(2, 18).Value = data['total']
        sheet.Cells(2, 19).Value = data['order_id']
        
        # 运行excel 中的VBA程序
        # macro_name = "保存发票等文件.保存清关发票"
        # excel.Application.Run(macro_name)
        
        # 执行py程序(替换原有的VBA程序)
        processor = EXCELProcessor(excel=excel,wb=wb)
        processor.process()
        wb.Save()
        wb.Close()
        excel.Quit()

    except Exception as e:
        print(f"Error accessing Excel or opening workbook: {e}")
        

if __name__ == '__main__':

    run()
