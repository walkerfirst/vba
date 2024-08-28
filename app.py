# pip install pywin32
import win32com.client
# 注意这里选择从
from config import shipment_file
from db import read_db,conn
from datetime import datetime

def run():
    # 获取发货信息
    sql = 'select * from shipView where id=1'
    ship_dict = read_db(sql, conn)[0]
    if ship_dict['model'] == '发货':
        # 指定执行VBA文件和 function
        macro_name = "保存发票等文件.保存清关发票"
        run_vba_code(data=ship_dict,macro_name=macro_name)

def run_vba_code(data,macro_name):
    
    if data['tax'] == 1.0:
        tax = '要退税'
    else:
        tax = '不退税'

    if data['express'] == '空运':
        express = 'by air'
    elif data['express'] == '海运':
        express = 'by sea'
    else:
        express = data['express']
    # 定义 workbooks 对象
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # 可以设置为 True 调试
    wb = excel.Workbooks.Open(shipment_file)
    # 定义工作表
    sheet = wb.Sheets('data')
    sm = wb.Sheets('情况说明')
    sheet1 = wb.Sheets('sheet1')

    # 修改sheet1单元格数据，让产品种类初始化为1
    sheet1.Cells(9,9).Value = 1
    sheet1.Cells(14,3).Value = data['chinese']

    # 修改情况说明页面的单元格数据，让产品初始化为1
    sm.Cells(1,21).Value = 1

    # data 页面单元格赋值
    sheet.Cells(2, 1).Value = data['chinese']
    sheet.Cells(2, 2).Value = data['name']
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
    # 运行VBS程序
    excel.Application.Run(macro_name)
    wb.Save()
    wb.Close()
    # print(datetime.now())
    # excel.Quit()

if __name__ == '__main__':

    run()