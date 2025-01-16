import sqlite3
_local = False
if _local:
    shipment_file = r"C:\\Users\\Administrator\\工作本地\\发货\\shipment.xlsm"
    chemical_file = 'C:\\Users\Administrator\\project\\pharmasiAdmin\\instance\\chemical.db'
    File_PATH = r"C:\\Users\\Administrator\\工作本地"
else:
    chemical_file = 'X:\\pharmasiAdmin\\instance\\chemical.db'
    shipment_file = r"Z:\\工作\\发货\\shipment.xlsm"
    File_PATH = r"Z:\工作"
conn = sqlite3.connect(chemical_file)

