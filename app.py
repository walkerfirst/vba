"""
主程序
1. 读取数据库中的发货信息
2. 打开指定的发货文件
3. 设置发货信息
4. 执行VBA程序
"""
# pip install pywin32
import sys
import win32com.client
from config import shipment_file
from window import create_window
from db import read_db,execute_db
from datetime import datetime
from tkinter import messagebox
import os
from vba_replacement import EXCELProcessor
from tkinter import Tk, StringVar, OptionMenu, Button, ttk
from DHL_bill_process import ImportDHLBill

# 定义 workbooks 对象
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False  # 可以设置为 True 调试
excel.DisplayAlerts = False # 禁用警告
wb = excel.Workbooks.Open(shipment_file)

def run():
    if not os.path.exists(shipment_file):
        msg_window = create_window()
        messagebox.showerror("错误", f"文件 {shipment_file} 不存在",parent=msg_window)
        return
    
    ship_data = fetch_ship_data()

    # 数据处理并执行vba程序
    frame_layout(ship_data)

def frame_layout(data):
    """设置主窗口布局,并传递选择的数据"""

    root = Tk()
    root.title("主窗口")
    root.geometry("400x350")
    root.resizable(False, False)

    # 添加窗口关闭协议处理
    def on_closing():
        # print("窗口关闭事件触发！")  # 检查是否执行
        root.destroy()  # 销毁窗口
        sys.exit(0)  # 退出程序

    # 绑定关闭事件
    root.protocol("WM_DELETE_WINDOW", on_closing) 

    # 运输方式映射字典
    order_options = data
    frame_data = {}
    
    # 主框架
    main_frame = ttk.Frame(root, padding="20 20 20 20")
    main_frame.pack(fill="both", expand=True)
    
    # 统一样式
    style = ttk.Style()
    style.configure('Title.TLabel', font=('Microsoft YaHei', 15, 'bold'), padding=5)
    style.configure('TLabel', font=('Microsoft YaHei', 12), padding=5)
    style.configure('TCombobox', font=('Microsoft YaHei', 11), padding=5)
    style.configure('TButton', font=('Microsoft YaHei', 10), padding=3)
    
    # 标题
    label = ttk.Label(main_frame, text="订  单  发  货  系  统", style='Title.TLabel')
    label.pack(pady=(10, 20)) # 上下间距

# 1. 订单货物名称下拉菜单（占窗口80%宽度）
    order_frame = ttk.Frame(main_frame)
    order_frame.pack(fill="x", pady=(0, 20))
    
    # ttk.Label(shipping_frame, text="运输方式：").pack(side='left', padx=(0, 10))
    var_order = StringVar()
    order_menu = ttk.Combobox(order_frame,
                               textvariable=var_order,
                               values=list(order_options.keys()),
                               state="readonly")
    order_menu.pack(side='left', fill="x", expand=True, padx=(0, 10))  # 占满剩余空间
    # 设置 current(0) 为默认选项
    order_menu.current(0)

    # 2. 包裹种类和手动重量（包裹种类占80%宽度）
    options_frame = ttk.Frame(main_frame)
    options_frame.pack(fill="x", pady=(0, 20))
    
    package_frame = ttk.Frame(options_frame)
    package_frame.pack(side='left', fill="x", expand=True)  # 包裹种类扩展
    ttk.Label(package_frame, text="货物种类：").pack(anchor='w')
    var_package = StringVar()
    package_menu = ttk.Combobox(package_frame,
                              textvariable=var_package,
                              values=["1", "2", "3"],
                              state="readonly")
    package_menu.pack(fill="x", expand=True)  # 占满剩余空间
    package_menu.current(0)

    # 客户选择（固定宽度）
    customer_frame = ttk.Frame(options_frame)
    customer_frame.pack(side='left', padx=(20, 0))  # 左边距20px
    ttk.Label(customer_frame, text="客户选择：").pack(anchor='w')
    var_weight = StringVar()
    customer_menu = ttk.Combobox(customer_frame,
                             textvariable=var_weight,
                             values=["HanChem Co., Ltd", "SIGMA-ALDRICH ISRAEL LTD"],
                             state="readonly",
                             width=25)  # 固定宽度
    customer_menu.current(0)
    customer_menu.pack()

    # 3. 按钮区域（优化布局）
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill="x", pady=(20, 0))
    
    # 使用grid布局使按钮居中并靠近
    button_frame.columnconfigure(0, weight=2)
    button_frame.columnconfigure(1, weight=2)
    button_frame.columnconfigure(2, weight=2)
    button_frame.columnconfigure(3, weight=2)
    
    # 产地证按钮
    cof_btn = ttk.Button(button_frame,
                          text="产地证",
                          command=lambda: cof_action(root,order_options[var_order.get()]),  # 传递root
                          width=6) # 固定宽度6px
    cof_btn.grid(row=0, column=0, padx=3) # 设置为第一行,第一列,左右间距3px

    # 删除按钮
    delete_btn = ttk.Button(button_frame,
                          text="删除",
                          command=lambda: delete_action(root,order_options[var_order.get()]),  # 传递root
                          width=5)
    delete_btn.grid(row=0, column=1, padx=3) 

    # 导入DHL账单按钮
    dhl_btn = ttk.Button(button_frame,
                          text="导入DHL账单",
                          command=lambda: ImportDHLBill(root),
                          width=10)
    dhl_btn.grid(row=0, column=2, padx=3)

     # 修改：在点击 "确认发货" 时才获取当前选项
    confirm_btn = ttk.Button(button_frame,
                           text="确认发货",
                           command=lambda: process_data(
                               root,
                               {
                                   'order_id': order_options[var_order.get()],
                                   'type_qty': var_package.get(),
                                   'customer': var_weight.get()
                               }
                           ),
                           width=8)
    confirm_btn.grid(row=0, column=3, padx=3) 
    root.mainloop()

def process_data(root,frame_data):
    try:
        root.destroy() # 关闭主窗口
        # 读取数据库中的发货信息
        order_id = frame_data['order_id']
        sql =f'select * from shipView where order_id = "{order_id}"'
        data = read_db(sql)[0]
        # print(data)
        # 定义工作表
        sheet = wb.Sheets('data')
        sm = wb.Sheets('情况说明')
        sheet1 = wb.Sheets('sheet1')

        sheet1.Cells(9,9).Value = frame_data['type_qty'] # 产品种类
        sheet1.Cells(2, 4).Value = frame_data['customer'] # 手动输入开关

        # 数据预处理
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
        # wb.Save()
        # 运行excel 中的VBA程序
        # macro_name = "保存发票等文件.保存清关发票"
        # excel.Application.Run(macro_name)
        
        # 执行py程序(替换原有的VBA程序)
        processor = EXCELProcessor(excel=excel,wb=wb)
        processor.process()
        wb.Save()
        wb.Close()
        excel.Quit()
        root.quit() # 退出主程序

    except Exception as e:
        print(f"Error accessing Excel or opening workbook: {e}")

def cof_action(root,order_id):
    """cof导出操作"""
    data = read_db(f'select * from shipView where order_id = "{order_id}"')[0]
    cas = data['cas']
    product = read_db(f"select hs from product where cas='{cas}'")[0]
    hs = product['hs']
    if not hs:
        messagebox.showerror("错误", f"产品 {cas} 的HS编码不存在, 请检查数据库",parent=root)
        return
    data['hs'] = hs
    from export_cof import update_cof_excel
    update_cof_excel(data)
    # root.destroy()
    # 弹出窗口提示
    messagebox.showinfo("产地证导出", f"订单: {order_id}  {data['chinese']} ({data['qty']} KG) \n \n导出成功",parent=root)

def delete_action(root,order_id):
    """删除ship表中的记录操作"""
    if order_id != '1':
        order_id = order_id
        sql = f'delete from ship where order_id = "{order_id}"'
        execute_db(sql)
        refresh_data(root)  # 删除后立即刷新
        # 弹出窗口提示
        messagebox.showinfo("提示", f"订单 {order_id} 已删除",parent=root)
    else:
        messagebox.showerror("错误", "无订单",parent=root)

def refresh_data(root):
    """清空并重建窗口（接收数据参数）"""
    # for widget in root.winfo_children():
    #     widget.destroy()
    root.destroy()  # 销毁当前窗口
    new_data = fetch_ship_data()  # 获取最新数据
    frame_layout(new_data)  # 使用传入的最新 data

def fetch_ship_data():
    """从数据库获取最新数据（示例）"""
   # 获取发货信息
    sql = 'select order_id,chinese from shipView where model="发货"'
    ship_dict = read_db(sql)
    ship_data = {}
    # 将字典列表重新调整一个新字段, 中文名(id) 为key, id 为value
    if ship_dict:
        for item in ship_dict:
            ship_data[f"{item['chinese']}({item['order_id']})"] = item['order_id']
    else:
        ship_data = {'无订单': '1'}
    return ship_data

if __name__ == '__main__':

    run()
