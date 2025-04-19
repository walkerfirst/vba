"""基本配置文件, 配置全局变量"""
import sqlite3,os

_local = True # 本地模式开关

if _local:
    # 本地路径配置
    base_dir = os.path.normpath(r"C:\Users\Administrator")  # 使用原始字符串确保路径正确
    chem_base = os.path.join(base_dir, "project") # 化学品数据库目录
else:
    # 远程路径配置
    base_dir = "Z:" 
    chem_base = "X:"

# 工作目录(生成的pdf文件 父目录)
FILE_PATH = os.path.join(base_dir, "工作")

# 发货excel文件
shipment_file = os.path.join(r"C:\Users\Administrator\工作", "发货", "shipment.xlsm")
FAPIAO_PATH = os.path.join(FILE_PATH, "发票")

# cof产地址模版导出文件
cof_file = os.path.join(FAPIAO_PATH, "cof.xlsx")

# 数据库文件
chemical_file = os.path.join(chem_base, "pharmasiAdmin", "instance", "chemical.db")

# DHL账单导入路径
csv_path = os.path.join(FILE_PATH, "发货", "快递账单")

conn = sqlite3.connect(chemical_file)

PRINTER_NAME = "HP LaserJet Professional M1213nf MFP" # 打印机名称

# 供应商字典
Supplier_DICT = {'902518-11-0':{'code':'744099904','company':'濮阳惠成电子材料股份有限公司','tel':'0393-8961801'},
                 '328-70-1':{'code':'137513392','company':'常州市仁科化工有限公司','tel':'13775109918'},
                 '3096-56-8':{'code':'744099904','company':'濮阳惠成电子材料股份有限公司','tel':'0393-8961801'},
                '109384-19-2':{'code':'563141835','company':'上海皓伯化工科技有限公司','tel':'13918007836'},
                '358-23-6':{'code':'MA0832ML8','company':'中船重工（邯郸）派瑞特种气体有限公司','tel':'18146248368'},
                '945-51-7':{'code':'563957295','company':'广东云星生物技术有限公司','tel':'18121060502'}}  

