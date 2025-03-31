"""
功能: 读取指定路径下的csv文件,将数据保存到数据库中
1. 读取数据库中的运单号 list
2. 读取跟单记录中未有运费的订单运单号 list
3. 读取csv文件中的数据
4. 如果数据库中没有记录,则添加
5. 如果订单中有未添加的运费,则自动添加
6. 弹出提示框,显示保存的记录数和更新的运费数
"""

import os
from csv_reader import extract_columns_from_csv
from db import execute_db,read_db_list,read_db
from datetime import datetime
from tkinter import messagebox
from window import create_window
from config import csv_path

def ImportDHLBill(root):
    root.destroy() # 关闭主窗口
    if not os.path.exists(csv_path):
        print(f"错误: 路径 {csv_path} 不存在")
    else:
        # 读取数据库中的运单号 list
        waybill_list = read_db_list('select waybill from shipping_record') 
        # 读取跟单记录中未有运费的订单运单号 list
        order_waybill_list = read_db_list('select waybill from orders where freight is NULL and express="DHL" and shipping="已收到货" and status=1')
        # print(waybill_list)
        # print(order_waybill_list)
        # 读取csv文件中的数据
        data = extract_columns_from_csv(csv_path)
        i = 0
        j = 0
        for item in data:
            # 如果数据库中没有记录,则添加
            if item['waybill'] not in waybill_list:
                # 转换日期格式为 datetime
                ship_date_str = item['ship_date'].replace("'", "''")
                try:
                    ship_date = datetime.strptime(ship_date_str, "%Y/%m/%d")
                except:
                    ship_date = datetime.strptime(ship_date_str, "%Y-%m-%d")

                save_sql = f"""INSERT INTO shipping_record 
                        (waybill, bill_date, ship_date, pcs, weight, amount, import_time)
                        VALUES (
                        '{item['waybill'].replace("'", "''")}', 
                        '{item['bill_date'].replace("'", "''")}',
                        '{ship_date}',
                        {item['pcs']},
                        {item['weight']},
                        {item['amount']},
                        '{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}')"""
                execute_db(save_sql)
                i +=1
            # 如果订单中有未添加的运费,则自动添加
            if item['waybill'] in order_waybill_list:
                update_waybill = item['waybill']
                freight = item['amount']
                update_sql = f"UPDATE orders SET freight = {freight} WHERE waybill = '{update_waybill}'"
                execute_db(update_sql)
                j +=1

        # print(f"共保存 {i} ,{j} 条数据")
        msg_window = create_window()
        messagebox.showinfo("提示", f"共保存  {i}  条 记录 \n \n跟单记录中更新了  {j}  条运费", parent=msg_window)
        msg_window.destroy()
if __name__ == '__main__':
    ImportDHLBill()