from config import conn

def read_db(sql,conn):
    # 执行 SQL 查询语句
    cursor = conn.cursor()
    cursor.execute(sql)
    data = cursor.fetchall()
    columns = [desc[0] for desc in cursor.description]  # 获取字段名称
    result = []
    for row in data:
        row_dict = dict(zip(columns, row))  # 将字段名称和值一起存储在字典中
        result.append(row_dict)
    cursor.close()
    return result

if __name__ == '__main__':
    sql = 'select * from shipView where id=1'
    data = read_db(sql,conn)
    for row in data:
        print(row)
