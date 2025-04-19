import win32com.client as win32
from config import cof_file
import pythoncom

def update_cof_excel(order_data):
    # 启动Excel应用
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    
    try:
        # 打开工作簿
        wb = excel.Workbooks.Open(cof_file)
        ws = wb.Sheets('Sheet1')
        
        # 处理数据
        pcs = int(order_data['pcs'])
        package = 'DRUM' if order_data['package'] == 'DRUM2' else order_data['package']
        if pcs > 1:
            package += "S"
        
        # 写入数据
        ws.Range('B4').Value = order_data['hs']
        ws.Range('C4').Value = order_data['chinese']
        ws.Range('D4').Value = order_data['name']
        ws.Range('E4').Value = pcs
        ws.Range('F4').Value = package
        ws.Range('G4').Value = float(order_data['gross'])
        ws.Range('J4').Value = float(order_data['gross'])
        ws.Range('K4').Value = float(order_data['qty'])
        ws.Range('S4').Value = float(order_data['ask'])
        ws.Range('R4').Value = float(order_data['ask'])
        
        # 保存并关闭
        wb.Save()
        wb.Close()
        
    except Exception as e:
        print(f"Excel操作出错: {e}")
        # 确保出错时也关闭Excel
        if 'wb' in locals():
            wb.Close(False)  # 不保存关闭
    finally:
        excel.Quit()  # 退出Excel应用
        # 强制释放 COM 对象
        del wb
        del excel
        pythoncom.CoUninitialize()