"""基本配置文件"""
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
File_PATH = os.path.join(base_dir, "工作")

# 发货excel文件
shipment_file = os.path.join(File_PATH, "发货", "shipment.xlsm")

# cof导出文件
cof_file = os.path.join(File_PATH, "发货", "cof.xlsx")

# 数据库文件
chemical_file = os.path.join(chem_base, "pharmasiAdmin", "instance", "chemical.db")

# DHL账单导入路径
csv_path = os.path.join(File_PATH, "发货", "快递账单")

conn = sqlite3.connect(chemical_file)

