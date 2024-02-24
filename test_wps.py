import win32com.client


def run_macro_wps(wps_file, macro_name):
    wps = win32com.client.Dispatch("Kwps.Application")  # 使用 WPS 表格的 COM 对象
    wps.Visible = False  # 设置 WPS 表格不可见

    # 打开 WPS 表格文件
    doc = wps.Documents.Open(wps_file)

    # 运行宏
    wps.Application.Run(macro_name)

    # 保存并关闭文件
    doc.Save()
    doc.Close()

    # 退出 WPS 表格程序
    wps.Quit()


if __name__ == "__main__":
    # 指定要打开的 WPS 表格文件路径和要运行的宏名称
    macro_name = "保存发票等文件.保存清关发票"
    wps_file = r"C:\Users\Administrator\PycharmProjects\VBA\shipment2.xlsm"

    # 运行宏
    run_macro_wps(wps_file, macro_name)