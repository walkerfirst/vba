# pip install pywin32
import win32com.client
from config import excel_file
from db import read_db,conn
from datetime import datetime
def run_vba_code(excel_file,macro_name):
    print(datetime.now())
    sql = 'select * from shipView where id=1'
    data = read_db(sql, conn)[0]
    print(datetime.now())
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # 可以设置为 True 调试
    wb = excel.Workbooks.Open(excel_file)
    # 选择指定名称的工作表
    sheet = wb.Sheets('data')
    if data['tax']==1.0:
        tax = '要退税'
    else:
        tax = '不退税'
    print(datetime.now())
    # 修改数据
    sheet.Cells(2, 1).Value = data['chinese']
    sheet.Cells(2, 2).Value = data['name']
    sheet.Cells(2, 3).Value = '上海盛傲化学有限公司'
    sheet.Cells(2, 4).Value = tax
    sheet.Cells(2, 5).Value = data['express']
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
    print(datetime.now())
    excel.Application.Run(macro_name)
    wb.Save()
    wb.Close()
    print(datetime.now())
    excel.Quit()

if __name__ == '__main__':

    # 指定要打开的 Excel 文件路径
    macro_name = "保存发票等文件.保存清关发票"
    run_vba_code(excel_file,macro_name)