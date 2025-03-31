import win32com.client

def set_textbox_content(file_path):
    # 打开 Excel 应用
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # 运行时不显示 Excel 窗口
    
    # 打开工作簿
    wb = excel.Workbooks.Open(file_path)
    
    # 读取 sheet1 的 C9 和 D9 单元格内容
    ws = wb.Sheets("sheet1")
    chinese = ws.Range("C9").Value
    english = ws.Range("D9").Value
    
    # 目标工作表
    report_sheet = wb.Sheets("报关单")
    
    # 设置图片透明度
    try:
        shape = report_sheet.Shapes("公章")
        shape.Fill.Transparency = 1  # 100% 透明
    except Exception as e:
        print("Error setting image transparency:", e)
    
    # 目标工作表列表
    sheets = {
        "PL": "公司名",
        "invoice": "公司名",
        "申报要素": "公司名",
        "销售合同": "公司名",
        "情况说明fedex": "公司名"
    }
    
    for sheet_name, shape_name in sheets.items():
        try:
            sheet = wb.Sheets(sheet_name)
            shape = sheet.Shapes(shape_name)
            text_range = shape.TextFrame2.TextRange
            if sheet_name == "销售合同":
                text_range.Text = f"Supplier：\n{chinese}\n{english}"
            elif sheet_name == "情况说明fedex":
                text_range.Text = chinese
            else:
                text_range.Text = f"{chinese}\n\n{english}"
        except Exception as e:
            print(f"Error setting text in {sheet_name}: {e}")
    
    # 保存并关闭 Excel
    wb.Save()
    wb.Close()
    excel.Quit()
# 使用示例
set_textbox_content("C:\\Users\\Administrator\\工作本地\\发货\\shipment2.xlsm")