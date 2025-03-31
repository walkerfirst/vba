"""数据库操作模块"""

from config import conn

def read_db(sql):
    """
    将多个字段数据转换成字典的list
    """
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
def read_db_list(sql):
    """
    将单个字段数据转换成list
    """
    cursor = conn.cursor()
    try:
        cursor.execute(sql)
        return [row[0] for row in cursor.fetchall()]
        
    except Exception as e:
        print(f"查询出错: {str(e)}")
        return None
    finally:
        cursor.close()

def execute_db(sql):
    """
    执行sql语句到sqlite
    """
    cursor = conn.cursor()
    cursor.execute(sql)
    conn.commit()
    cursor.close()

if __name__ == '__main__':
    sql = 'select * from shipView where id=1'
    data = read_db(sql)
    for row in data:
        print(row)
