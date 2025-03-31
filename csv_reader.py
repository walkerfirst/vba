"""从CSV文件中提取指定列数据"""

import os
import csv
import chardet

def detect_encoding(file_path):
    """自动检测文件编码"""
    with open(file_path, 'rb') as f:
        rawdata = f.read(10000)
    return chardet.detect(rawdata)['encoding']

def extract_columns_from_csv(folder_path):
    """从CSV文件中提取指定列数据"""
    target_columns = ['运单号', '发件日期','账单日期', '件数', '重量', '金额合计（人民币）']
    # 存放提取的数据
    data = []
    for filename in os.listdir(folder_path):
        if not filename.lower().endswith('.csv'):
            continue
        file_path = os.path.join(folder_path, filename)
        print(f"\n正在处理文件: {filename}")
        # 检测文件编码
        encoding = detect_encoding(file_path) or 'gbk'
        with open(file_path, 'r', encoding=encoding, errors='replace') as csvfile:
            # 自动检测分隔符
            try:
                dialect = csv.Sniffer().sniff(csvfile.read(1024))
                csvfile.seek(0)
                reader = csv.DictReader(csvfile, dialect=dialect)
            except:
                csvfile.seek(0)
                reader = csv.DictReader(csvfile)  # 回退到默认读取方式
            
            # 检查目标列是否存在
            available_cols = reader.fieldnames
            missing_cols = [col for col in target_columns if col not in available_cols]
            if missing_cols:
                print(f"  → 缺少列: {', '.join(missing_cols)}")
                continue
            for row in reader:
                try:
                    item = {
                        'waybill': row['运单号'].strip(),
                        'bill_date': row['账单日期'].strip(),
                        'ship_date': row['发件日期'].strip(),
                        'pcs': int(float(row['件数'])) if row['件数'].strip() else None,
                        'weight': float(row['重量']) if row['重量'].strip() else None,
                        'amount': float(row['金额合计（人民币）'].replace(',', '')) if row['金额合计（人民币）'].strip() else None,
                    }
                    data.append(item)
            
                except ValueError as e:
                    print(f"[{filename}] 数据转换错误: {e} | 行内容: {row}")
                    continue
                
    return data

if __name__ == '__main__':
    folder_path = "./data/"  # 直接指定路径
    if not os.path.exists(folder_path):
        print(f"错误: 路径 {folder_path} 不存在")
    else:
        data= extract_columns_from_csv(folder_path)
        print(data)