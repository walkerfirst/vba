import os
import win32com.client
from config import shipment_file, File_PATH
import tkinter as tk
from tkinter import messagebox
from datetime import datetime

# 配置
EXCEL_PATH = shipment_file # Excel文件路径
FATHER_PATH = File_PATH       # 父路径
FAPIAO_PATH = File_PATH + '\发票'  # 发票保存路径
PRINTER_NAME = "HP LaserJet Professional M1213nf MFP" # 打印机名称

class EXCELProcessor:
    def __init__(self):
        self.excel = win32com.client.Dispatch("Excel.Application")
        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        self.wb = self.excel.Workbooks.Open(EXCEL_PATH)
        self.sheet1 = self.wb.Sheets("sheet1")
        self.invoice_sheet = self.wb.Sheets("invoice")
        self.pl_sheet = self.wb.Sheets("PL")
        self.sm_sheet = self.wb.Sheets("情况说明")
        
    def get_cell_value(self, cell_ref):
        """获取单元格值"""
        return self.sheet1.Range(cell_ref).Value

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
            return True
        except Exception as e:
            print(f"生成PDF失败: {str(e)}")
            print(f"详细错误信息: {e.__class__.__name__}")
            if hasattr(e, 'excepinfo'):
                print(f"Excel错误代码: {e.excepinfo[5]}")
            return False

    def generate_multiple_pdf(self, sheets_to_export, output_path):
        """生成多个PDF文件并合并"""
        try:
            # 取消所有选择
            self.excel.ActiveWindow.SelectedSheets.Select(False)
            
            # 选择多个工作表
            sheets = [self.wb.Sheets(sheet) for sheet in sheets_to_export]
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
            print(f"成功导出组合PDF: {output_path}")
            
            # 取消选择
            self.excel.ActiveWindow.SelectedSheets.Select(False)
        except Exception as e:
            print(f"导出组合PDF失败: {str(e)}")

    def print_sheet(self, sheet_name, copies=1):
        """
        直接打印Excel工作表
        sheet_name: 工作表名称
        copies: 打印份数
        """
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
            
            # 设置页边距（单位：厘米）
            sheet.PageSetup.LeftMargin = 1.5
            sheet.PageSetup.RightMargin = 1.5
            sheet.PageSetup.TopMargin = 1.5
            sheet.PageSetup.BottomMargin = 1.5
            
            # 设置对齐方式
            sheet.PageSetup.CenterHorizontally = False  # 取消水平居中
            sheet.PageSetup.CenterVertically = False    # 取消垂直居中
            sheet.PageSetup.LeftMargin = 1            # 设置左边距为1厘米
            sheet.PageSetup.TopMargin = 1.5             # 设置上边距为1.5厘米
            
            # 统一页眉页脚
            sheet.PageSetup.CenterHeader = ""
            sheet.PageSetup.CenterFooter = ""
            sheet.PageSetup.LeftHeader = ""
            sheet.PageSetup.LeftFooter = ""
            sheet.PageSetup.RightHeader = ""
            sheet.PageSetup.RightFooter = ""

                
            # 使用win32api打印
            try:
                import win32api
                import win32print
                
                # 获取默认打印机
                default_printer = win32print.GetDefaultPrinter()
                if default_printer != PRINTER_NAME:
                    # print(f"设置默认打印机为: {PRINTER_NAME}")
                    win32print.SetDefaultPrinter(PRINTER_NAME)
                
                # 直接打印Excel
                # print(f"开始直接打印 {sheet_name}")
                for _ in range(copies):
                    sheet.PrintOut()
                # print("打印任务已发送")
            except Exception as e:
                print(f"打印失败: {str(e)}")
        except Exception as e:
            print(f"打印时出错: {str(e)}")

    def process(self):
        """主处理逻辑"""
        # 创建确认对话框
        confirm_window = tk.Tk()
        confirm_window.overrideredirect(1)  # 完全隐藏窗口装饰
        confirm_window.withdraw()  # 立即隐藏主窗口
        
        # 提高DPI感知
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except Exception as dpi_error:
            print(f"DPI设置失败: {str(dpi_error)}")
            
        # 设置窗口缩放
        confirm_window.tk.call('tk', 'scaling', 2.0)
        
        # 配置字体
        default_font = ('Microsoft YaHei', 10)  # 使用更清晰的字体
        confirm_window.option_add('*Font', default_font)
        
        # 确保窗口完全隐藏
        confirm_window.update_idletasks()
        confirm_window.update()
        
        # 获取数据
        company = self.get_cell_value("C9") # 发件抬头
        express = self.get_cell_value("C11") # 运输商
        model = self.get_cell_value("I11") # 贸易方式
        tax = self.get_cell_value("E14") # 退税
        tracing = self.get_cell_value("C12") # 单号
        name = self.get_cell_value("C15") # 申报名称
        pcs = self.get_cell_value("L11") # 包裹数量
        nw = self.invoice_sheet.Range("G29").Value # 总净重
        invoice_no = self.invoice_sheet.Range("E5").Value # 发票号
        package = self.get_cell_value("C20")  # 包装类型
        ask_value = self.get_cell_value("L9")  # 总报关价值
        gw = self.pl_sheet.Range("K29").Value  # 总毛重
        
        # 构建确认信息
        confirm_msg = f"注意: {tax}\n发件抬头：{company}\n贸易方式：{model}\n\n" \
                     f"运输商：{express}\n单号：{tracing}\n申报名称：{name}\n\n" \
                     f"包装：{package}\n总价值：{ask_value}\n总件数: {pcs}\n" \
                     f"总净重：{nw}\n总毛重：{gw}"
        
        # 显示确认对话框
        result = messagebox.askyesno("确认", confirm_msg, parent=confirm_window)
        
        # 清理窗口
        confirm_window.destroy()
        
        if not result:
            print("用户取消运行脚本")
            return
        # 发货为快递的情况
        if express.lower() == "dhl":
            file_name= f"{express}_{tracing}_{invoice_no}" # 给客户单据的文件名
            file_name2 = f"上海盛傲_{tracing}" # 报关用单据的文件名
            # 创建隐藏窗口用于对话框
            dialog_window = tk.Tk()
            dialog_window.overrideredirect(1)
            dialog_window.withdraw()
            dialog_window.attributes('-topmost', True)  # 使窗口置顶
            
            # 显示打印确认对话框
            print_confirm = messagebox.askyesno("确认", "是否打印确认单？", parent=dialog_window)
            
            # 销毁隐藏窗口
            dialog_window.destroy()
            
            if print_confirm:
                self.print_sheet("情况说明", copies=2) # 打印情况说明
                self.print_labels(pcs) # 设置并打印标签

            self.generate_pdf("invoice", os.path.join(FAPIAO_PATH, f"{file_name2}_invoice.pdf"))
            self.generate_pdf("PL", os.path.join(FAPIAO_PATH, f"{file_name2}_PL.pdf"))
            self.generate_pdf("报关委托书", os.path.join(FAPIAO_PATH, f"{file_name2}_委托书.pdf"))
            self.generate_pdf("报关单", os.path.join(FAPIAO_PATH, f"{file_name2}_报关单.pdf"))
            self.generate_pdf("申报要素", os.path.join(FAPIAO_PATH, f"{file_name2}_申报要素.pdf"))
            
        # 空运或海运情况
        elif express.lower() in ["by sea", "by air"]:
            file_name = f"{express}_{invoice_no}"
            print(file_name)
            # 导出PL(2)工作表
            self.generate_pdf("PL(2)",
                           os.path.join(FAPIAO_PATH, f"{file_name}_PL.pdf"))
            
            # 导出组合PDF
            sheets_to_export = ["invoice", "PL", "报关委托书", "报关单", 
                              "申报要素", "情况说明fedex", "销售合同"]
            output_path = os.path.join(FAPIAO_PATH, 
                                     f"上海盛傲报关资料_{invoice_no}_{name}_{nw}KG.pdf")
            # 生成报关用PDF
            self.generate_multiple_pdf(sheets_to_export, output_path)
        else:
            file_name= f"{express}_{tracing}_{invoice_no}"

        # 保存给客户的发票
        self.generate_pdf("invoice(2)", 
                       os.path.join(FAPIAO_PATH, f"{file_name}_invoice.pdf"))
        
        # 保存商业发票
        self.generate_pdf("CI", 
                       os.path.join(FAPIAO_PATH, f"{file_name}_commercial invoice.pdf"))
        
        if tax == "要退税" and  model == "一般贸易":
            self.process_tax_refund(company, file_name, nw)


    def process_tax_refund(self, company, file_name, nw):
        """处理退税"""
        if company == "上海盛傲化学有限公司":
            # 获取订单ID
            order_id = self.wb.Sheets("data").Range("S2").Value
            if order_id:
                order_id = int(order_id)
            else:
                order_id = "未指定"
            
            # 获取申报名称
            name = self.get_cell_value("C15")
            # 获取当年年份
            year = datetime.now().year
            # 创建文件夹
            folder_name = f"{order_id}_{year}_{name}_{nw}KG"
            folder_path = os.path.join(FATHER_PATH, "退税", folder_name)
            os.makedirs(folder_path, exist_ok=True)
            
            # 导出多个PDF文件
            self.generate_pdf("销售合同",
                           os.path.join(folder_path, f"{file_name}_contract.pdf"))
            
            self.generate_pdf("invoice",
                           os.path.join(folder_path, f"{file_name}_invoice.pdf"))
            
            self.generate_pdf("PL",
                           os.path.join(folder_path, f"{file_name}_PL.pdf"))
            

    def print_labels(self, pcs):
        """打印标签"""
        sheet = self.wb.Sheets("标签")
        # 根据包裹数量设置打印区域
        try:
            pcs = int(pcs)  # 确保pcs为整数
            if pcs == 1:
                sheet.PageSetup.PrintArea = "$A$2:$G$9"
                self.print_sheet("标签")
            elif pcs == 2:
                sheet.PageSetup.PrintArea = "$A$2:$G$18"
                self.print_sheet("标签")
            else:
                copies = (pcs // 3) + 1
                sheet.PageSetup.PrintArea = "$A$2:$G$29"
                self.print_sheet("标签", copies=copies)
        except Exception as e:
            print(f"打印标签时出错: {str(e)}")

if __name__ == "__main__":

    processor = EXCELProcessor()
    processor.process()
