"""
Python脚本替代VBA脚本
"""

import os
import win32com.client,win32print
from config import FILE_PATH,FAPIAO_PATH,PRINTER_NAME
from datetime import datetime
from window import window_askyesno


class EXCELProcessor:
    def __init__(self,excel,wb,root):
        self.wb = wb
        self.excel = excel
        self.sheet1 = self.wb.Sheets("sheet1")
        self.pl_sheet = self.wb.Sheets("PL")
        self.lable = self.wb.Sheets("标签")
        self.root = root

    def get_cell_value(self, cell_ref):
        """获取单元格值"""
        return self.sheet1.Range(cell_ref).Value

    def set_textbox_content(self, chinese, english):
        """设置文本框内容
        新增逻辑：如果"情况说明fedex"中的内容已经是chinese，则跳过所有设置
        """
        # 目标工作表列表
        sheets = {
            "PL": "公司名",
            "invoice": "公司名",
            "申报要素": "公司名",
            "销售合同": "公司名",
            "情况说明fedex": "公司名"
        }

        # 首先检查"情况说明fedex"是否需要更新
        try:
            sheet = self.wb.Sheets("情况说明fedex")
            shape = sheet.Shapes("公司名")
            current_text = shape.TextFrame2.TextRange.Text.strip()
            if current_text == chinese.strip():
                print("情况说明fedex内容已相同，跳过设置")
                return  # 内容已相同，直接返回不执行任何设置
        except Exception as e:
            Errormsg = f"检查情况说明fedex时出错: {e}"
            window_askyesno(self.root, "错误", Errormsg)
            return

        # 正常设置所有工作表内容
        for sheet_name, shape_name in sheets.items():
            try:
                sheet = self.wb.Sheets(sheet_name)
                shape = sheet.Shapes(shape_name)
                text_range = shape.TextFrame2.TextRange
                
                if sheet_name == "销售合同":
                    text_range.Text = f"Supplier：\n{chinese}\n{english}"
                elif sheet_name == "情况说明fedex":
                    text_range.Text = chinese
                else:
                    text_range.Text = f"{chinese}\n\n{english}"

            except Exception as e:
                Errormsg = f"Error setting text in {sheet_name}: {e}"
                window_askyesno(self.root, "错误", Errormsg)

        # 保存 Excel
        self.wb.Save()

    def generate_pdf(self, sheet_name, file_path):
        """使用Excel内置功能生成PDF"""
        try:
            # 取消所有选择
            self.excel.ActiveWindow.SelectedSheets.Select(False)
            
            # 只选择目标工作表
            sheet = self.wb.Sheets(sheet_name)
            sheet.Select()
            
            # 确保只有目标工作表被选中
            if self.excel.ActiveWindow.SelectedSheets.Count > 1:
                print(f"警告：检测到多个工作表被选中，正在清除选择")
                self.excel.ActiveWindow.SelectedSheets.Select(False)
                sheet.Select()
            
            # 确保当前活动工作表是目标工作表
            if self.excel.ActiveSheet.Name != sheet_name:
                sheet.Activate()
            
            # 导出PDF
            sheet.ExportAsFixedFormat(0, file_path)  # 0 = xlTypePDF
            print(f"成功导出PDF: {file_path}")
            
            # 取消选择
            self.excel.ActiveWindow.SelectedSheets.Select(False)

        except Exception as e:
            print(f"生成PDF失败: {str(e)}")
            print(f"详细错误信息: {e.__class__.__name__}")
            if hasattr(e, 'excepinfo'):
                print(f"Excel错误代码: {e.excepinfo[5]}")

    def generate_multiple_pdf(self, sheets_to_export, output_path):
        """生成多个PDF文件并合并"""
        try:
            # 确保所有工作表取消选择
            self.excel.ActiveWindow.SelectedSheets.Select(False)
            
            # 激活第一个工作表
            first_sheet = self.wb.Sheets(sheets_to_export[0])
            first_sheet.Activate()
            
            # 选择多个工作表
            sheets = [self.wb.Sheets(sheet) for sheet in sheets_to_export]
            
            # 逐个选择目标工作表
            for sheet in sheets:
                sheet.Select(False)  # Add to selection without unselecting others
                
            # 确保正确数量的工作表被选中
            if self.excel.ActiveWindow.SelectedSheets.Count != len(sheets_to_export):
                print(f"警告：工作表选择数量不匹配，正在重新选择")
                self.excel.ActiveWindow.SelectedSheets.Select(False)
                for sheet in sheets:
                    sheet.Select(False)
            
            # 导出为PDF
            self.wb.ActiveSheet.ExportAsFixedFormat(
                Type=0,  # xlTypePDF
                Filename=output_path,
                Quality=0,  # xlQualityStandard
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=True
            )
            # print(f"成功导出组合PDF: {output_path}")
            
            # 取消选择
            self.excel.ActiveWindow.SelectedSheets.Select(False)
        except Exception as e:
            print(f"导出组合PDF失败: {str(e)}")

    def process(self):
        """主处理逻辑"""
        # 获取数据
        company = self.get_cell_value("C9") # 发件抬头
        english = self.get_cell_value("D9") # 发件抬头英文
        express = self.get_cell_value("C11") # 运输商
        model = self.get_cell_value("I11") # 贸易方式
        tax = self.get_cell_value("E14") # 退税
        tracing = int(self.get_cell_value("C12")) # 单号取整数,否则有小数点
        if tracing == 0:
            tracing = ""
        tracing = str(tracing)
        name = self.get_cell_value("C15") # 申报名称
        pcs = int(self.get_cell_value("L11")) # 包裹数量
        nw = self.pl_sheet.Range("J29").Value # 总净重
        gw = self.pl_sheet.Range("K29").Value  # 总毛重
        invoice_no = self.get_cell_value("L12") # 发票号
        package = self.get_cell_value("C20")  # 包装类型
        ask_value = self.get_cell_value("L9")  # 总报关价值
        consingee = self.get_cell_value("D2")  # 收件人
        
        # 构建确认信息
        confirm_msg = f"{tax}\n{model}\n\n"\
                    f"{company}\n{consingee}\n{tracing}  {express}\n\n"\
                     f"{name}\nUSD {ask_value}\n\n"\
                     f"{pcs}  {package}\nNET：{nw}\nG.W.：{gw}"
        
        # 定义需要高亮的关键字及颜色
        highlight_keywords = {
            "退税": "green",
            "不退税": "red",
            "DRUM": "orange",
            "CARTON": "orange",
        }

        # 显示确认对话框
        result = window_askyesno(self.root,"发货信息确认", confirm_msg,keywords=highlight_keywords)
        # 清理窗口
        if not result:
            print("用户取消运行脚本") 
            return
        
        # 更新excel 中的文本框内容
        self.set_textbox_content(chinese=company, english=english)

        # 发货为快递的情况
        if express.lower() == "dhl":
            file_name= f"{express}_{tracing}_{invoice_no}" # 给客户单据的文件名
            file_name2 = f"上海盛傲_{tracing}" # 报关用单据的文件名

            copies = self.set_labels(pcs) # 获取打印份数并设置标签打印区域

            # 显示打印确认对话框
            label_print_confirm = window_askyesno(self.root,"确认", f"是否打印DHL标签, 共 {copies} 份？",keywords={'标签':"orange"})
            if label_print_confirm:
                unit_net2 = self.wb.Sheets("data").Range("k2").Value # 单件净重2 (第二种包装)
                if unit_net2:
                    if copies > 1: # 多页并且两种标签时，先打印第一个标签, 再打印第二个标签
                        self.lable.Cells(6, 8).Value = "No" # 设置手动设置开关为NO(不显示第二种标签)
                        self.print_sheet("标签", copies=copies-1) # 设置并打印标签
                        self.lable.Cells(6, 8).Value = "YES"
                        self.print_sheet("标签", copies=1)
                    else:
                        self.lable.Cells(6, 8).Value = "YES"
                        self.print_sheet("标签", copies=1)
                else:
                    self.print_sheet("标签", copies=copies)

            file_print_confirm = window_askyesno(self.root,"确认", "是否 打印 DHL情况说明？",keywords={'情况说明':"orange"})
            if file_print_confirm:
                self.print_sheet("情况说明", copies=2) # 打印情况说明

            # 保存报关用单据
            file_list = ["invoice", "PL", "报关委托书", "报关单", "申报要素"]
            for _file in file_list:
                self.generate_pdf(_file, os.path.join(FAPIAO_PATH, f"{file_name2}_{_file}.pdf"))
            
        # 空运或海运情况
        elif express.lower() in ["by sea", "by air"]:
            file_name = f"{express}_{invoice_no}"
            # 导出PL(2)工作表
            self.generate_pdf("PL(2)",
                           os.path.join(FAPIAO_PATH, f"{file_name}_PL.pdf"))
            
            # 导出组合PDF
            sheets_to_export = ["invoice", "PL", "报关委托书", "报关单", 
                              "申报要素", "销售合同","情况说明fedex"]
            output_path = os.path.join(FAPIAO_PATH, 
                                     f"上海盛傲报关资料_{invoice_no}_{name}_{nw}KG.pdf")
            # 生成报关用PDF
            self.generate_multiple_pdf(sheets_to_export, output_path)
            # 打开生成的PDF文件
            os.startfile(output_path)
        
        # 其他情况
        else:
            file_name= f"{express}_{tracing}_{invoice_no}"

        # 保存给客户的发票
        self.generate_pdf("invoice(2)", 
                       os.path.join(FAPIAO_PATH, f"{file_name}_invoice.pdf"))
        
        # 保存商业发票
        self.generate_pdf("CI", 
                       os.path.join(FAPIAO_PATH, f"{file_name}_CI.pdf"))
        
        """处理退税"""
        if tax == "退税" and  model == "一般贸易":
            if company == "上海盛傲化学有限公司":
                # 获取订单ID
                order_id = self.wb.Sheets("data").Range("S2").Value
                if order_id:
                    order_id = int(order_id)
                else:
                    order_id = "未指定"
                # 获取当年年份
                year = datetime.now().year
                if nw > 1:
                    nw = int(nw)
                # 创建文件夹
                folder_name = f"{order_id}_{name}_{nw} KG_y{year}"
                folder_path = os.path.join(FILE_PATH, "退税", folder_name)
                os.makedirs(folder_path, exist_ok=True)
                # 导出多个PDF文件
                tax_file_list = ["invoice", "PL", "销售合同"]
                for tax_file in tax_file_list:
                    self.generate_pdf(tax_file, os.path.join(folder_path, f"{file_name}_{tax_file}.pdf"))
            else:
                window_askyesno(self.root,"错误", "退税公司 不匹配 \n \n请检查发件抬头")
    
    def set_labels(self, pcs):
        """设置标签,并返回打印份数"""
        
        # 根据包裹数量设置打印区域
        pcs = int(pcs)  # 确保pcs为整数
        copies = 1
        if pcs == 1:
            self.lable.PageSetup.PrintArea = "$A$2:$F$9"
        elif pcs == 2:
            self.lable.PageSetup.PrintArea = "$A$2:$F$18"
        else:
            N = pcs / 3
            # print(f"包裹数量: {pcs}, 计算出的N值: {N}")
            if isinstance(N,int):
                copies = N
            else:
                copies = int(N)+1
            self.lable.PageSetup.PrintArea = "$A$2:$F$29"
        return copies
                
    def print_sheet(self, sheet_name, copies=1):
        try:
            sheet = self.wb.Sheets(sheet_name)
            # 设置打印区域为A4大小
            sheet.PageSetup.Zoom = 100  # 100%缩放
            sheet.PageSetup.FitToPagesWide = 1
            sheet.PageSetup.FitToPagesTall = False
            # 设置页面为A4
            sheet.PageSetup.PaperSize = 9  # xlPaperA4 = 9
            # 设置打印方向为纵向
            sheet.PageSetup.Orientation = 1  # xlPortrait = 1
            # 设置页边距（单位：豪米?）
            sheet.PageSetup.LeftMargin = 10
            sheet.PageSetup.RightMargin = 5
            sheet.PageSetup.TopMargin = 35  # 增加顶部页边距
            sheet.PageSetup.BottomMargin = 2
            # 设置对齐方式
            sheet.PageSetup.CenterHorizontally = False  # 取消水平居中
            sheet.PageSetup.CenterVertically = False    # 取消垂直居中
            # 使用win32api打印
            try: 
                # 获取默认打印机
                default_printer = win32print.GetDefaultPrinter()
                # 检查并设置打印机
                if default_printer != PRINTER_NAME:
                    win32print.SetDefaultPrinter(PRINTER_NAME)
                # 循环打印多份
                for _ in range(copies):
                    sheet.PrintOut()
            except Exception as e:
                print(f"打印失败: {str(e)}")
        except Exception as e:
            print(f"打印时出错: {str(e)}")

if __name__ == "__main__":
    from config import shipment_file
    EXCEL_PATH = shipment_file # Excel文件路径
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False # 禁用警告
    wb = excel.Workbooks.Open(EXCEL_PATH)
    processor = EXCELProcessor(excel=excel,wb=wb)
    processor.process()
