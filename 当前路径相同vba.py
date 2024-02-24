import win32com.client
import os

# 获取当前脚本所在目录
script_dir = os.path.dirname(os.path.abspath(__file__))


def run_vba_code(filename, macro_name):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # 可以设置为 True 调试
    wb = excel.Workbooks.Open(filename)
    excel.Application.Run(macro_name)
    wb.Save()
    excel.Quit()



# 指定要打开的 Excel 文件路径
filename = os.path.join(script_dir, "pdf1.0.xlsm")
macro_name = "保存发票等文件.保存清关发票"
# 执行 VBA 代码
run_vba_code(filename, macro_name)